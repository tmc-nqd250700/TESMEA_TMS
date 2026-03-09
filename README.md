# TESMEA_TMS

Ứng dụng **TESMEA_TMS** là một dự án .NET (WPF/desktop) dùng để cấu hình và quản lý các thông số đo kiểm (biến tần, ống gió, cảm biến) và kịch bản cho nghiệp vụ đo kiểm quạt công nghiệp.

## Yêu cầu hệ thống

- **Hệ điều hành**: Windows 10 trở lên  
- **IDE**: Visual Studio 2019/2022 (khuyến nghị bản mới nhất)  
- **.NET**: Phiên bản tương ứng với solution (mở `.sln` để Visual Studio tự tải SDK cần thiết nếu thiếu)  

## Hướng dẫn chạy dự án

1. Mở file solution `TESMEA_TMS.sln` bằng Visual Studio.
2. Chọn cấu hình `Debug` và thiết lập project startup phù hợp (ứng dụng WPF chính).
3. **Lần chạy Debug đầu tiên**:
   - Copy file `tesmea.db` vào thư mục `runtimes` trong thư mục output (ví dụ: `bin\Debug\runtimes`) để lấy db cập nhật.
   - Đảm bảo ứng dụng có quyền đọc/ghi trên file database này.
4. Nhấn **Start Debugging** (`F5`) để chạy ứng dụng.

## Cấu trúc dự án (khái quát)

- `Services/` – Chứa các service xử lý nghiệp vụ, kết nối external app, xử lý dữ liệu (`ExternalAppService.cs`, v.v.).
- `ViewModels/` – Lớp ViewModel cho các màn hình chính (theo mô hình MVVM).
- `Views/` – Giao diện XAML của ứng dụng, gồm các màn hình chính và dialog.
- `Helpers/` – Các lớp tiện ích, xử lý dữ liệu dùng chung.

## Ghi chú phát triển

- Tuân thủ mô hình **MVVM** cho View/ViewModel.
- Hạn chế logic nghiệp vụ trong `View` (XAML code-behind); tập trung ở `ViewModels` và `Services`.
- Khi thay đổi cấu trúc database (`tesmea.db`), đảm bảo cập nhật tương ứng trong code truy cập dữ liệu.

## Đóng góp & liên hệ

- Commit message nên ngắn gọn, mô tả **mục đích thay đổi**.
- Khi thêm tính năng mới, cập nhật README nếu cần (đặc biệt là yêu cầu cấu hình/cách chạy).

Mọi thắc mắc về dự án có thể trao đổi trực tiếp trong team hoặc qua kênh liên lạc nội bộ.