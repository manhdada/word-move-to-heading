# Word Move to Heading

Add-in C# miễn phí cho Microsoft Word trên Windows, giúp di chuyển nhanh nội
dung đến một Heading khác và tự động nhận diện Outline level.

## Tính năng

- Thêm mục **Move to** thành một dòng trong menu chuột phải.
- Submenu tự động liệt kê các Heading hiện có trong tài liệu.
- Di chuyển vùng đang chọn đến cuối phần thuộc Heading đích và giữ nguyên định dạng.
- Giữ nguyên vị trí đọc sau khi di chuyển; không cuộn màn hình đến vị trí đích.
- Nút **Home → Outline tools → Nhận diện Heading** tự động gán Outline level:

  | Dạng tiêu đề | Outline level |
  |---|---:|
  | `I.`, `II.`, `III.` | 1 |
  | `1.`, `2.`, `3.` | 2 |
  | `1.1`, `1.2` | 3 |
  | `a)`, `b)`, `c)` | 4 |
  | `1.1.1` trở lên | Tăng tương ứng, tối đa 9 |

- Hỗ trợ cả số gõ trực tiếp và numbering tự động của Word.
- Toàn bộ mỗi thao tác có thể hoàn tác bằng một lần `Ctrl+Z`.

## Cài đặt

1. Mở trang [Releases](../../releases/latest).
2. Tải `WordMoveToHeading-Setup.exe`.
3. Đóng hoàn toàn Microsoft Word.
4. Chạy bộ cài và mở lại Word.

Bộ cài đăng ký add-in cho tài khoản Windows hiện tại, không yêu cầu quyền
Administrator. Hỗ trợ Word 32-bit và 64-bit.

> Bộ cài hiện chưa được ký bằng chứng thư thương mại. Windows SmartScreen có thể
> hiển thị cảnh báo; kiểm tra mã SHA-256 trên trang Release trước khi chạy.

## Cách sử dụng Move to

1. Bôi đen nội dung cần chuyển.
2. Nhấp chuột phải và chọn **Move to**.
3. Chọn Heading đích.

Nội dung được chèn ngay trước Heading kế tiếp có cấp bằng hoặc cao hơn. Nếu
không có Heading kế tiếp, nội dung được đặt ở cuối tài liệu.

## Gỡ cài đặt

Mở **Windows Settings → Apps → Installed apps → Word Move to Heading →
Uninstall**. Hãy đóng Word trước khi gỡ.

## Build từ mã nguồn

Yêu cầu:

- Windows 10/11.
- Microsoft Word desktop và Office Primary Interop Assemblies.
- .NET Framework 4.x.

Chạy PowerShell:

```powershell
./build.ps1
```

Kết quả được tạo trong thư mục `dist/`.

## Quyền riêng tư

Add-in chạy hoàn toàn trên máy, không gửi nội dung tài liệu ra Internet và không
thu thập dữ liệu sử dụng.

## Báo lỗi

Khi báo lỗi, vui lòng ghi rõ phiên bản Word, Windows, các bước tái hiện và đính
kèm ảnh chụp nếu có. Không đăng tài liệu có dữ liệu nhạy cảm.

## Giấy phép

[MIT](LICENSE)


