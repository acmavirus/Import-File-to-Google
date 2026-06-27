# Import File to Google Sheets

Ứng dụng Windows để import nội dung từ `DOCX`, `XLSX`, `XLS` và `CSV` lên Google Sheets.

## Tính năng hiện có

- Tạo Google Sheet mới từ file nguồn.
- Tách `DOCX` thành các sheet riêng cho bảng và phần văn bản.
- Import `XLSX` với dữ liệu, kích thước cột/dòng và một phần định dạng cơ bản.
- Hỗ trợ `XLS` thông qua `xlrd`.
- Hỗ trợ `CSV` với nhiều kiểu mã hóa phổ biến.
- Mở kết quả trực tiếp trên trình duyệt sau khi import xong.
- Có installer Windows và context menu chuột phải.

## Chuẩn bị Google Credentials

Ứng dụng cần file OAuth `credentials.json` kiểu Desktop App từ Google Cloud Console.

1. Tạo project trên Google Cloud Console.
2. Bật Google Sheets API.
3. Tạo OAuth 2.0 Client ID với loại `Desktop app`.
4. Tải file JSON về và đổi tên thành `credentials.json`.

Khi cài đặt bằng wizard, hãy chọn file này ở bước cấu hình Google Credentials.
Sau khi cài, file `credentials.json` sẽ được chép vào thư mục cài đặt.

## Chạy ứng dụng

- Chạy file `ImportToSheet.exe` và truyền vào đường dẫn file cần import.
- Hoặc click chuột phải lên file `.docx`, `.xlsx`, `.xls`, `.csv` rồi chọn `Import to Google Sheets`.

## Build

```powershell
python build_exe.py
```

Script sẽ tạo:

- `dist\ImportToSheet.exe`
- `dist\Setup_ImportToSheet.exe`

## GitHub Actions và Releases

- Mỗi lần push lên `main`, GitHub Actions sẽ tự build 2 file exe và lưu dưới dạng artifact.
- Khi push một tag có dạng `v*` như `v1.0.1`, workflow sẽ tự tạo GitHub Release và đính kèm 2 file exe.
- File workflow nằm tại `.github/workflows/build-release.yml`.

## Lưu ý

- Nếu không có `credentials.json` trong thư mục cài đặt, ứng dụng sẽ không thể xác thực Google.
- `.xls` cần thư viện `xlrd`.
- Một số định dạng nâng cao của Excel như công thức phức tạp, merged cells hoặc conditional formatting chưa được tái tạo đầy đủ.
