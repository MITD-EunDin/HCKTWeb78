using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;
using WebReport78.Models;


namespace WebReport78.Controllers
{
    public class ProtectDutyController : Controller
    {
        private readonly ILogger<ProtectDutyController> _logger;
        private readonly XGuardContext _context;
        public ProtectDutyController(ILogger<ProtectDutyController> logger, XGuardContext context)
        {
            _logger = logger;
            _context = context;
        }

        public DateTime SwapMonthDay(string dateStr)
        {
            string[] parts = dateStr.Split('/');
            if (parts.Length != 3)
                throw new ArgumentException("Định dạng ngày không hợp lệ (phải là M/d/yyyy)");

            int month = int.Parse(parts[0]);
            int day = int.Parse(parts[1]);
            int year = int.Parse(parts[2]);

            // Đảo vị trí tháng <-> ngày
            return new DateTime(year, day, month);
        }

        private DateTime ParseExcelCellAsText(ExcelRange cell, int row)
        {
            if (cell?.Value == null || string.IsNullOrWhiteSpace(cell.Text))
                throw new ArgumentException($"Ô ngày trống tại dòng {row}");

            // Nếu Excel lưu dạng số serial date
            if (cell.Value is double serialDate)
            {
                var date = DateTime.FromOADate(serialDate);
                return SwapMonthDay(date.ToString("d/M/yyyy")); // Ví dụ: 1/8/2025
            }

            // Nếu là DateTime chuẩn
            if (cell.Value is DateTime dateTime)
            {
                string formattedDate = dateTime.ToString("d/M/yyyy"); // "1/8/2025"
                string[] parts = formattedDate.Split('/');
                string swapped = $"{parts[1].Trim()}/{parts[0].Trim()}/{parts[2].Trim()}"; // ngày/tháng/năm
                return SwapMonthDay(swapped); // Ví dụ: 1/8/2025
            }

            // Nếu Excel đã lưu sẵn dạng text
            return SwapMonthDay(cell.Text.Trim());
        }

        [HttpGet] // hiển thị form mới theo month cần tìm
        public IActionResult Index(string monthYearImport)
        {
            // hiện nút lưu khi import file
            ViewData["CanSave"] = false;

            // disable nút sửa xóa khi tháng đã qua
            bool disableSave = false;

            // Nếu không có monthYearImport, lấy tháng hiện tại
            if (string.IsNullOrEmpty(monthYearImport))
            {
                monthYearImport = DateTime.Now.ToString("yyyy-MM");
            }

            var dateNow = DateTime.Now;
            if (!string.IsNullOrEmpty(monthYearImport) && DateTime.TryParseExact(monthYearImport, "yyyy-MM", null, System.Globalization.DateTimeStyles.None, out var parsedMonthYear))
            {
                // Kiểm tra tháng đã qua
                if (parsedMonthYear.Year < dateNow.Year ||
                    (parsedMonthYear.Year == dateNow.Year && parsedMonthYear.Month < dateNow.Month))
                {
                    disableSave = true;
                }

                var duties = _context.DutyShifts
                    .Where(d => d.MothYearImport.Year == parsedMonthYear.Year &&
                                d.MothYearImport.Month == parsedMonthYear.Month)
                    .OrderBy(d => d.Dateduty)
                    .ToList();

                ViewData["MonthYearImport"] = monthYearImport;
                ViewData["DisableSave"] = disableSave;

                if (!duties.Any())
                {
                    TempData["Error"] = $"Chưa có dữ liệu cho tháng {monthYearImport}.";
                }
                return View(duties);
            }

            TempData["Error"] = "Định dạng tháng/năm không hợp lệ.";
            ViewData["DisableSave"] = true; // Disable luôn nếu format sai
            return View(new List<DutyShift>());
        }


        [HttpPost]
        public IActionResult ImportFile(IFormFile file, string monthYearImport)
        {
            if (file == null || file.Length == 0 || string.IsNullOrEmpty(monthYearImport))
            {
                TempData["Error"] = "Vui lòng chọn file và tháng.";
                return RedirectToAction(nameof(Index));
            }

            if (!DateTime.TryParseExact(monthYearImport, "yyyy-MM", null, System.Globalization.DateTimeStyles.None, out var parsedMonthYear))
            {
                TempData["Error"] = "Định dạng tháng/năm không hợp lệ.";
                return RedirectToAction(nameof(Index));
            }
            var dataDuty = new List<DutyShift>();
            var dateNow = DateTime.Now;
            try
            {
                if (parsedMonthYear.Year < dateNow.Year || parsedMonthYear.Year == dateNow.Year && parsedMonthYear.Month < dateNow.Month)
                {
                    TempData["Error"] = $"Không nhập cho tháng {monthYearImport} vì tháng đã qua";
                    return RedirectToAction(nameof(Index));
                }
                else
                {

                    using (var stream = new MemoryStream())
                    {
                        file.CopyTo(stream);
                        stream.Position = 0;
                        using (var package = new ExcelPackage(stream))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            var rowCount = worksheet.Dimension.Rows;
                            for (int row = 2; row <= rowCount; row++) // bắt đầu từ dòng 2 dòng 1 mặc định là tiêu đề
                            {
                                var parsedDate = ParseExcelCellAsText(worksheet.Cells[row, 1], row);
                                var duties = new DutyShift
                                {
                                    //Id = int.Parse(worksheet.Cells[row, 1].Text),
                                    //Dateduty = DateOnly.FromDateTime(worksheet.Cells[row, 1].GetValue<DateTime>()),
                                    Dateduty = DateOnly.FromDateTime(parsedDate),
                                    //Dateduty = parsedDate,
                                    Fullname = worksheet.Cells[row, 2].Text?.Trim() ?? "",
                                    Rankk = worksheet.Cells[row, 3].Text?.Trim() ?? "",
                                    Department = worksheet.Cells[row, 4].Text?.Trim() ?? "",
                                    Phonenumber = worksheet.Cells[row, 5].Text?.Trim() ?? "",
                                    MothYearImport = new DateOnly(parsedMonthYear.Year, parsedMonthYear.Month, 1),
                                    DateCreated = DateTime.Now
                                };
                                // thêm đối tuong duty vua tao và danh sach để ty lưu và db
                                dataDuty.Add(duties);
                            }
                        }
                    }
                }
                ViewData["MonthYearImport"] = monthYearImport;
                TempData["ImportedData"] = System.Text.Json.JsonSerializer.Serialize(dataDuty); // Lưu tạm dữ liệu vào TempData
                ViewData["CanSave"] = true;
                return View("Index", dataDuty);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi import file Excel.");
                TempData["Error"] = ex.Message;
                ViewData["CanSave"] = false;
                return RedirectToAction(nameof(Index));
            }
        }


        [HttpPost]
        public IActionResult SaveImportData(string monthYearImport)
        {
            if (string.IsNullOrEmpty(monthYearImport) || !TempData.ContainsKey("ImportedData"))
            {
                TempData["Error"] = "Không có dữ liệu để lưu.";
                return RedirectToAction(nameof(Index));
            }

            // Parse monthYearImport sang DateTime để lấy Year &Month
            if (!DateTime.TryParseExact(monthYearImport, "yyyy-MM", null, System.Globalization.DateTimeStyles.None, out var parsedDate))
            {
                TempData["Error"] = "Định dạng tháng/năm không hợp lệ.";
                return RedirectToAction(nameof(Index));
            }
            try
            {
                var importData = System.Text.Json.JsonSerializer.Deserialize<List<DutyShift>>(
                (string)TempData["ImportedData"]);

                var existingDuties = _context.DutyShifts
                    .Where(d => d.MothYearImport.Year == parsedDate.Year &&
                                d.MothYearImport.Month == parsedDate.Month)
                    .ToList();
                _context.DutyShifts.RemoveRange(existingDuties);

                // Thêm dữ liệu mới
                _context.DutyShifts.AddRange(importData.OrderBy(d => d.Dateduty));
                _context.SaveChanges();

                TempData["Success"] = "Dữ liệu đã được lưu thành công.";
                ViewData["CanSave"] = false; // Ẩn nút Lưu
                return RedirectToAction(nameof(Index), new { monthYearImport });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi lưu dữ liệu vào cơ sở dữ liệu.");
                TempData["Error"] = $"Lưu dữ liệu thất bại: {ex.Message}";
                return RedirectToAction(nameof(Index), new { monthYearImport });
            }
        }

        [HttpGet]
        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(DutyShift model, string monthYearImport)
        {
            if (!ModelState.IsValid)
            {
                return PartialView("Create", model); // Trả về partial view nếu validation thất bại
            }
            var dateNow = DateTime.Now;
            try
            {
                if (!DateTime.TryParseExact(monthYearImport, "yyyy-MM", null, System.Globalization.DateTimeStyles.None, out var parsedMonthYear))
                {
                    ModelState.AddModelError("", "Định dạng tháng/năm không hợp lệ.");
                    return PartialView("Create", model);
                }

                if (parsedMonthYear.Year < dateNow.Year || parsedMonthYear.Year == dateNow.Year && parsedMonthYear.Month < dateNow.Month)
                {
                    TempData["Error"] = $"Không nhập cho tháng {monthYearImport} vì tháng đã qua";
                    return RedirectToAction(nameof(Index));
                }
                else
                {
                    model.MothYearImport = new DateOnly(parsedMonthYear.Year, parsedMonthYear.Month, 1);
                    model.DateCreated = DateTime.Now;
                    model.DateUpdate = DateTime.Now; // Cập nhật DateUpdate
                    _context.DutyShifts.Add(model);
                    _context.SaveChanges();

                    TempData["Success"] = "Thêm mới thành công.";
                    return RedirectToAction(nameof(Index), new { monthYearImport });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi thêm mới dữ liệu.");
                ModelState.AddModelError("", $"Lỗi khi thêm mới: {ex.Message}");
                TempData["Success"] = "Lỗi khi thêm mới";
                return PartialView("Create", model);
            }
        }

        [HttpGet]
        public IActionResult Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var duty = _context.DutyShifts.FirstOrDefault(d => d.Id == id);
            if (duty == null)
            {
                return NotFound();
            }

            ViewData["MonthYearImport"] = $"{duty.MothYearImport.Year}-{duty.MothYearImport.Month:D2}";
            return PartialView("Edit", duty); // Trả về partial view Edit.cshtml
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(DutyShift model, string monthYearImport)
        {
            if (!ModelState.IsValid)
            {
                return PartialView("Edit", model); // Trả về partial view nếu validation thất bại
            }
            try
            {
                if (!DateTime.TryParseExact(monthYearImport, "yyyy-MM", null, System.Globalization.DateTimeStyles.None, out var parsedMonthYear))
                {
                    ModelState.AddModelError("", "Định dạng tháng/năm không hợp lệ.");
                    return PartialView("Edit", model);
                }


                var existingDuty = _context.DutyShifts.FirstOrDefault(d => d.Id == model.Id);
                if (existingDuty == null)
                {
                    return NotFound();
                }

                existingDuty.Dateduty = model.Dateduty;
                existingDuty.Fullname = model.Fullname;
                existingDuty.Rankk = model.Rankk;
                existingDuty.Department = model.Department;
                existingDuty.Phonenumber = model.Phonenumber;
                existingDuty.MothYearImport = new DateOnly(parsedMonthYear.Year, parsedMonthYear.Month, 1);
                existingDuty.DateUpdate = DateTime.Now; // Cập nhật DateUpdate

                _context.SaveChanges();
                TempData["Success"] = "Cập nhật thành công.";
                return RedirectToAction(nameof(Index), new { monthYearImport });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi khi cập nhật dữ liệu.");
                ModelState.AddModelError("", $"Lỗi khi cập nhật: {ex.Message}");
                TempData["Error"] = "Cập nhật không thành công.";
                return PartialView("Edit", model);
            }
        }

        [HttpGet]
        public IActionResult Delete(int id)
        {
            var duty = _context.DutyShifts.FirstOrDefault(d => d.Id == id);
            if (duty == null)
            {
                return NotFound();
            }

            var monthYearImport = $"{duty.MothYearImport.Year}-{duty.MothYearImport.Month:D2}";
            _context.DutyShifts.Remove(duty);
            _context.SaveChanges();

            TempData["Success"] = "Xóa thành công.";
            return RedirectToAction(nameof(Index), new { monthYearImport });
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
