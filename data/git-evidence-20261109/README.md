# Dữ liệu công việc đã được phép push — 09/11/2026

Người dùng cho phép push tất cả dữ liệu công việc. Các nguồn/snapshot đang bị .gitignore bỏ qua được lưu tại đây, giữ nguyên SHA-256. Một số JSON gốc vượt giới hạn 100 MB của GitHub; vì vậy nội dung trùng nhau chỉ lưu một lần trong ZIP và chia thành phần tối đa 40 MiB. Không thay đổi/xóa nguồn cục bộ. Manifest ánh xạ mỗi đường dẫn về đối tượng nguyên bản.

Kiểm độc lập toàn bộ phần nén và từng tệp, không ghi nguồn:

```powershell
python scripts/restore_git_evidence_20261109.py --verify-only
```

Khôi phục vào thư mục riêng:

```powershell
python scripts/restore_git_evidence_20261109.py --output D:/UVG-restored
```

Script không ghi đè tệp có nội dung khác. Các phần phải có đủ theo manifest. Tệp thường đã được Git theo dõi vẫn nằm tại vị trí ban đầu. Thư viện phụ thuộc, cache Python, khóa Excel, NDJSON inspection, log và build tạm không phải dữ liệu bàn giao. Bộ ZIP năm Excel dành cho người dùng vẫn ở outputs/import-2025-20261109/BO_IMPORT_HD_2025_20261109.zip.
