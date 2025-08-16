using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WebReport78.Models;
using OfficeOpenXml;
using System.Text.Json;
using WebReport78.Services;
using MongoDB.Driver;
using System.Globalization;
using MongoDB.Bson;
using Microsoft.EntityFrameworkCore; // Thêm để dùng CountAsync, ToListAsync

namespace WebReport78.Controllers
{
    public class InOutController : Controller
    {
        private readonly ILogger<InOutController> _logger;
        private readonly XGuardContext _context;
        private readonly MongoDbService _mongoService;
        private readonly IWebHostEnvironment _env;

        public InOutController(ILogger<InOutController> logger, XGuardContext context, MongoDbService mongoservice, IWebHostEnvironment env)
        {
            _logger = logger;
            _context = context;
            _mongoService = mongoservice;
            _env = env;
        }
        private class CameraSetting
        {
            public string source_id { get; set; }
            public string name { get; set; }
            public string type { get; set; }
        }

        private class CameraSettings
        {
            public string location_id { get; set; }
            public List<CameraSetting> cameras { get; set; }
        }

        private string ConvertTimestamp(int timestamp)
        {
            // chuyển sang datetime
            DateTime dateTime = DateTimeOffset.FromUnixTimeSeconds(timestamp)
                                              .ToLocalTime()
                                              .DateTime;

            // đinh dang dd/mm/yyyy HH:mm
            string formattedDate = dateTime.ToString("dd/MM/yyyy HH:mm");
            return formattedDate;
        }
        private long ConvertToUnixTimestamp(DateTime dateTime)
        {
            // Chuyển đổi DateTime sang Unix timestamp 
            return (long)(dateTime.ToUniversalTime() - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
        }
        public async Task<IActionResult> Index(string fromDate, string toDate, string filterType = "All", int page = 1, int pageSize = 100)
        {

            // Parse date range
            var (parsedFromDate, parsedToDate, fromTimestamp, toTimestamp) = ParseDateRange(fromDate, toDate);

            // tải config json để biết loại vào ra và đi sớm về muộn
            var cameraSettings = LoadCameraSettings();
            var checkInCamera = cameraSettings.cameras.FirstOrDefault(c => c.type == "in")?.source_id;
            var checkOutCamera = cameraSettings.cameras.FirstOrDefault(c => c.type == "out")?.source_id;
            var locationId = cameraSettings.location_id;

            // tính quân số, số khách trong ngày
            DateTime today = DateTime.Today;
            int soldierTotal = await _context.Staff.CountAsync(s => s.IdTypePerson.HasValue && (s.IdTypePerson.Value == 0 || s.IdTypePerson.Value == 2));
            int guestCount = await _context.Staff.CountAsync(s => s.IdTypePerson.HasValue && s.IdTypePerson.Value == 3 && s.DateCreated.HasValue && s.DateCreated.Value.Date == today);
            long todayFromTs = ConvertToUnixTimestamp(today);
            long todayToTs = ConvertToUnixTimestamp(today.AddHours(23).AddMinutes(59));
            int soldierCurrentToday = await CalculateCurrentSoldiers(todayFromTs, todayToTs, checkInCamera, checkOutCamera, locationId);

            // set giá trị để hiện ra table
            ViewData["SoldierTotal"] = soldierTotal;
            ViewData["SoldierCurrent"] = soldierCurrentToday;
            ViewData["GuestCount"] = guestCount;
            ViewBag.FromDate = parsedFromDate;
            ViewBag.ToDate = parsedToDate;
            ViewBag.Filter = filterType;
            ViewBag.Page = page;
            ViewBag.PageSize = pageSize;

            // Lấy và thực thi dữ liệu event log, lọc thêm theo location_id và sourceID chỉ thuộc 2 camera
            var collection = _mongoService.GetCollection<eventLog>("EventLog");
            var filter = Builders<eventLog>.Filter.And(
                Builders<eventLog>.Filter.Gte(x => x.time_stamp, fromTimestamp),
                Builders<eventLog>.Filter.Lte(x => x.time_stamp, toTimestamp),
                Builders<eventLog>.Filter.Eq(x => x.locationId, locationId),
                Builders<eventLog>.Filter.In(x => x.sourceID, new[] { checkInCamera, checkOutCamera })
            );
            var totalRecords = await collection.CountDocumentsAsync(filter);
            var data = await collection.Find(filter)
                                      .SortBy(x => x.time_stamp)
                                      .Skip((page - 1) * pageSize)
                                      .Limit(pageSize)
                                      .ToListAsync();

            ViewBag.TotalPages = (int)Math.Ceiling((double)totalRecords / pageSize);
            ProcessEventLog(data, checkInCamera, checkOutCamera, parsedFromDate, parsedToDate);

            // Lọc theo filterType
            if (filterType == "Late")
            {
                data = data.Where(x => x.type_eventLE == "L").ToList();
            }
            else if (filterType == "Early")
            {
                data = data.Where(x => x.type_eventLE == "E").ToList();
            }

            // Lưu vào session để export
            SaveToSession(data, soldierTotal, soldierCurrentToday, guestCount, filterType);

            return View(data);
        }
        private (DateTime, DateTime, long, long) ParseDateRange(string fromDate, string toDate)
        {
            DateTime today = DateTime.Today;
            DateTime parsedFromDate = today;
            DateTime parsedToDate = today.AddHours(23).AddMinutes(59);

            if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParseExact(fromDate, "yyyy-MM-ddTHH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var from))
            {
                parsedFromDate = from;
            }
            else if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParseExact(fromDate, "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out from))
            {
                parsedFromDate = from;
            }

            if (!string.IsNullOrEmpty(toDate) && DateTime.TryParseExact(toDate, "yyyy-MM-ddTHH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var to))
            {
                parsedToDate = to;
            }
            else if (!string.IsNullOrEmpty(toDate) && DateTime.TryParseExact(toDate, "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out to))
            {
                parsedToDate = to;
            }

            long fromTimestamp = ConvertToUnixTimestamp(parsedFromDate);
            long toTimestamp = ConvertToUnixTimestamp(parsedToDate);

            return (parsedFromDate, parsedToDate, fromTimestamp, toTimestamp);
        }

        private CameraSettings LoadCameraSettings()
        {
            var settingcamera = Path.Combine(_env.ContentRootPath, "settingcamera.json");
            var cameraSettingsJson = System.IO.File.ReadAllText(settingcamera);
            return JsonSerializer.Deserialize<CameraSettings>(cameraSettingsJson);
        }
        private async Task<int> CalculateCurrentSoldiers(long fromTimestamp, long toTimestamp, string checkInCamera, string checkOutCamera, string locationId)
        {
            var staffGuids = await _context.Staff
                .Where(s => s.IdTypePerson.HasValue && (s.IdTypePerson.Value == 0 || s.IdTypePerson.Value == 2))
                .Select(s => s.GuidStaff)
                .ToListAsync();

            var filter = Builders<eventLog>.Filter.And(
                Builders<eventLog>.Filter.Gte(x => x.time_stamp, fromTimestamp),
                Builders<eventLog>.Filter.Lte(x => x.time_stamp, toTimestamp),
                Builders<eventLog>.Filter.Eq(x => x.locationId, locationId),
                Builders<eventLog>.Filter.In(x => x.sourceID, new[] { checkInCamera, checkOutCamera }),
                Builders<eventLog>.Filter.In(x => x.userGuid, staffGuids)
            );

            var records = await _mongoService.GetCollection<eventLog>("EventLog")
                .Find(filter)
                .SortBy(x => x.time_stamp)
                .ToListAsync();

            var soldierStatus = new Dictionary<string, bool>();
            int soldierCurrent = 0;

            foreach (var record in records)
            {
                var userGuid = record.userGuid;
                bool isCheckIn = record.sourceID == checkInCamera;

                if (!soldierStatus.ContainsKey(userGuid))
                {
                    if (isCheckIn)
                    {
                        soldierStatus[userGuid] = true;
                        soldierCurrent++;
                    }
                    else
                    {
                        soldierStatus[userGuid] = false;
                    }
                }
                else
                {
                    if (isCheckIn && !soldierStatus[userGuid])
                    {
                        soldierStatus[userGuid] = true;
                        soldierCurrent++;
                    }
                    else if (!isCheckIn && soldierStatus[userGuid])
                    {
                        soldierStatus[userGuid] = false;
                        soldierCurrent--;
                    }
                }
            }

            return soldierCurrent < 0 ? 0 : soldierCurrent;
        }
        private void ProcessEventLog(List<eventLog> data, string checkInCamera, string checkOutCamera, DateTime fromDate, DateTime toDate)
        {
            _logger.LogInformation("Processing {Count} event logs", data.Count);
            var sources = _context.Sources.Select(s => new { s.Guid, s.Name }).ToList();
            var lateThreshold = fromDate.Date.AddHours(8).AddMinutes(30);
            var earlyThreshold = toDate.Date.AddHours(17).AddMinutes(30);

            var groupedData = data
                .GroupBy(x => new { x.userGuid, Date = DateTimeOffset.FromUnixTimeSeconds(x.time_stamp).ToLocalTime().DateTime.Date })
                .ToList();

            _logger.LogInformation("Found {GroupCount} groups of events", groupedData.Count);

            foreach (var group in groupedData)
            {
                var userGuid = group.Key.userGuid;
                var date = group.Key.Date;
                var staff = _context.Staff.FirstOrDefault(s => s.GuidStaff == userGuid);
                bool exclude = staff != null && staff.IdTypePerson == 0;
                var userRecords = group.ToList();

                _logger.LogInformation("Processing userGuid: {UserGuid}, Date: {Date}, Records: {RecordCount}, Exclude: {Exclude}", userGuid, date, userRecords.Count, exclude);

                foreach (var item in userRecords)
                {
                    //so sánh id camera và gán name
                    var source = sources.FirstOrDefault(s => s.Guid == item.sourceID);
                    item.cameraGuid = item.sourceID;
                    item.cameraName = source?.Name ?? item.sourceID;
                    if (source != null)
                    {
                        item.sourceID = source.Name;
                    }
                    item.formatted_date = ConvertTimestamp(item.time_stamp);

                    // phân loại nhận diện vào hay ra
                    if (item.cameraGuid == checkInCamera)
                        item.type_eventIO = "In";
                    else if (item.cameraGuid == checkOutCamera)
                        item.type_eventIO = "Out";
                    else
                        item.type_eventIO = "N/A";

                    item.type_eventLE = "";
                    item.IsLate = false;
                    item.IsLeaveEarly = false;

                    _logger.LogInformation("Event: userGuid={UserGuid}, time_stamp={TimeStamp}, sourceID={SourceID}, type_eventIO={TypeEventIO}", item.userGuid, item.time_stamp, item.sourceID, item.type_eventIO);
                }

                if (exclude)
                {
                    continue;
                }

                if (staff != null && staff.IdTypePerson == 2)
                {
                    // Lấy check-in đầu tiên trong ngày
                    var firstCheckIn = userRecords
                        .Where(x => x.type_eventIO == "In")
                        .OrderBy(x => x.time_stamp)
                        .FirstOrDefault();

                    // Gán mác "đi muộn" cho check-in đầu tiên nếu muộn hơn 8:30
                    if (firstCheckIn != null)
                    {
                        var checkInTime = DateTimeOffset.FromUnixTimeSeconds(firstCheckIn.time_stamp).ToLocalTime().DateTime;
                        if (checkInTime > lateThreshold)
                        {
                            firstCheckIn.type_eventLE = "L";
                            firstCheckIn.IsLate = true;
                            _logger.LogInformation("Marked as Late: userGuid={UserGuid}, checkInTime={CheckInTime}", userGuid, checkInTime);
                        }
                    }

                    // Lấy check-out cuối cùng trong ngày
                    var lastCheckOut = userRecords
                        .Where(x => x.type_eventIO == "Out")
                        .OrderByDescending(x => x.time_stamp)
                        .FirstOrDefault();

                    // Gán mác "về sớm" hoặc "check-out bình thường"
                    if (lastCheckOut != null)
                    {
                        var checkOutTime = DateTimeOffset.FromUnixTimeSeconds(lastCheckOut.time_stamp).ToLocalTime().DateTime;
                        if (checkOutTime < earlyThreshold)
                        {
                            // Kiểm tra xem có check-in nào sau check-out này nhưng trước 17:30 không
                            var hasLaterCheckIn = userRecords
                                .Any(x => x.type_eventIO == "In" &&
                                          x.time_stamp > lastCheckOut.time_stamp &&
                                          DateTimeOffset.FromUnixTimeSeconds(x.time_stamp).ToLocalTime().DateTime <= earlyThreshold);

                            if (!hasLaterCheckIn)
                            {
                                lastCheckOut.type_eventLE = "E";
                                lastCheckOut.IsLeaveEarly = true;
                            }
                            else
                            {
                                lastCheckOut.type_eventLE = "O"; // Check-out bình thường nếu có check-in sau
                            }
                        }
                        else
                        {
                            lastCheckOut.type_eventLE = "O"; // Check-out bình thường nếu sau 17:30
                        }
                    }
                }
            }
        }

        private void SaveToSession(List<eventLog> data, int soldierTotal, int soldierCurrent, int guestCount, string filterType)
        {
            HttpContext.Session.SetString("SoldierTotal", soldierTotal.ToString());
            HttpContext.Session.SetString("SoldierCurrent", soldierCurrent.ToString());
            HttpContext.Session.SetString("GuestCount", guestCount.ToString());
            HttpContext.Session.SetString("FilterType", filterType);

            var itemList = data.Select(item => new ItemModel
            {
                Name = item.Name,
                CheckTime = item.formatted_date,
                CheckType = item.type_eventIO,
                Source = item.sourceID,
                EndTime = item.type_eventLE // Lưu trạng thái đi muộn/về sớm
            }).ToList();

            HttpContext.Session.SetString("TableData", JsonSerializer.Serialize(itemList));
        }
        public List<ItemModel> getItemModel()
        {
            List<ItemModel> list = new List<ItemModel>();
            return list;
        }

        [HttpPost]
        public IActionResult ExReport(string fromDate, string toDate, string note)
        {
            try
            {
                // Lấy dữ liệu từ session
                var tableData = HttpContext.Session.GetString("TableData");
                // Lấy quân số và khách từ Session
                int soldierTotal = int.Parse(HttpContext.Session.GetString("SoldierTotal") ?? "0");
                int soldierCurrent = int.Parse(HttpContext.Session.GetString("SoldierCurrent") ?? "0");
                int guestCount = int.Parse(HttpContext.Session.GetString("GuestCount") ?? "0");
                // lấy loại lọc tìm kiếm 
                string filterType = HttpContext.Session.GetString("FilterType") ?? "";

                if (string.IsNullOrEmpty(tableData))
                {
                    return BadRequest("Không có dữ liệu để xuất.");
                }

                var data = JsonSerializer.Deserialize<List<ItemModel>>(tableData ?? "[]");

                // Định nghĩa tên file và đường dẫn mặc định
                var folder = @"D:\excel";
                var templatePath = Path.Combine(folder, "_TemplateWeb78.xlsx");

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                // tạo file mới
                var fileName = $"Attendance_Report_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";
                // biến lưu data cho vào file mới
                var stream = new MemoryStream();

                using (var package = new ExcelPackage(new FileInfo(templatePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                        return StatusCode(500, "Không tìm thấy worksheet trong file template.");

                    // phạm vi thời gian
                    var startDate = DateTime.ParseExact(fromDate, "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture);
                    var endDate = DateTime.ParseExact(toDate, "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture);

                    // ghi phạm vi thời gian
                    worksheet.Cells["C2"].Value = $"{startDate:dd-MM-yyyy HH:mm} - {endDate:dd-MM-yyyy HH:mm}";
                    // ghi quân số, số khách
                    worksheet.Cells["C3"].Value = $"{soldierCurrent} / {soldierTotal}";
                    worksheet.Cells["E3"].Value = $"{guestCount}";
                    // ghi loại tìm kiếm
                    worksheet.Cells["C4"].Value = $"{filterType}";

                    // Ghi chú
                    worksheet.Cells["C5"].Value = string.IsNullOrWhiteSpace(note) ? "Không có ghi chú" : note;

                    // ghi dữ liệu
                    for (int i = 0; i < data.Count; i++)
                    {
                        var row = i + 7;
                        worksheet.Cells[row, 1].Value = (i + 1).ToString();
                        worksheet.Cells[row, 2].Value = data[i].Name;
                        worksheet.Cells[row, 3].Value = data[i].CheckTime;
                        worksheet.Cells[row, 4].Value = data[i].CheckType;
                        worksheet.Cells[row, 5].Value = data[i].Source;
                        //worksheet.Cells[row, 6].Value = data[i].EndTime == "L" ? "Đi muộn" : data[i].EndTime == "E" ? "Về sớm" : data[i].EndTime; ;
                    }
                    package.SaveAs(stream);
                }
                // xuất file excel
                stream.Position = 0;

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Lỗi xuất file: {ex.Message}");
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}