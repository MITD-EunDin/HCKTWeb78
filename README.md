# HCKTWeb78
```
├── Controllers         
├──── InOutController.cs                   // báo cáo vào ra    
├──── ProtectDutyController.cs             // trực ban
├── Models
├──── DbWeSmart.cs                         // thay Collection này thành collection thực tế eventlog chỉ cần đổi tên              
├──── DutyShift.cs                         // tạo 1 table mới trong SQL Xguard để quản lý trực ban             
├──── ErrorViewModel.cs
├──── ItemModel.cs                         // model lưu trữ dữ liệu vào ra temp để xuất excel
├──── MongoDbService.cs                    // connect tới mongodb
├──── MongoDbSettings       
├──── Source.cs                           // thông tin camera trong xguard
├──── Staff.cs
├──── XGuardContext.cs                    // connect tới xguard
├── Views
├──── InOut  
├─────── ExportReport.cshtml
├─────── Header.cshtml
├─────── Index.cshtml
├──── ProtectDuty
├─────── Create.cshtml
├─────── Edit.cshtml
├─────── HeaderSection.cshtml
├─────── ImportFile.cshtml
├─────── Index.cshtml
├─────── Sidebar.cshtml
├──── Shared
├─────── _Layout.cshtml                  // giao diện tổng
├─────── _toastNotice.cshtml             // giaao diện thông báo thành công và lỗi
├── appsettings.json
├── efpt.config.json
├── Program.cs
└── settingcamera.json                  // thay source_id bằng source_id của camera nhận diện tương ứng
```

# Các phần cần sửa
- Sửa connectstring và database name trong appsetting.json phần MongoDb
- Sửa DbWeSmart.cs trong phần Models sao cho đúng tên Collection của Eventlog
- Thay source_id trong settingcamera.json thành souce_id của camera nhận diện tương ứng
- Trong Controllers -> InOutController.cs hàm ExReport phần [var folder = @"D:\excel"] sửa sao cho đúng đường dẫn chứa file Template Excel mẫu
