# Nghiệm thu AMIS / meInvoice ngày 13/11/2026

File đang dùng: **TONG_HOP_HD_2026_CHO_DON_2024_2025_CHUOI.xlsx**.
Gồm 82.207 HĐ phát hành 2026 cho đơn 2024–2025 và 68.928 nhóm chuỗi.
Tab chuỗi có ngày, thanh toán từng HĐ, màu riêng cho F0–F3 và giá trị sau cùng
đã gồm VAT. Giá trị cuối chỉ hiện ở dòng đầu mỗi nhóm.

F0 là HĐ gốc; F1 là điều chỉnh/thay thế trước 2026; F2 là xử lý năm 2026;
F3 là xuất mới/xuất lại năm 2026 nối qua BHH. Đây là các giai đoạn xử lý,
không phải cấp quan hệ pháp lý. HĐ xuất lại có thể cùng liên quan nhiều chuỗi
gốc của một BHH; dùng tab HĐ 2026 khi cộng tổng các HĐ duy nhất.

`TONG_HOP_HD_2026_CHO_DON_2024_2025_RUT_GON.xlsx` là bản đối chứng để kiểm
toàn bộ giá trị tab HĐ 2026. Các Excel map và kiểm 5 cặp là bằng chứng của
bước đối chiếu trước đó, không thay cho file đang dùng.

## Dữ liệu và kiểm chứng

- `data/1311/amis/`: hai sổ đã điền số HĐ và ký hiệu theo kết quả đối chiếu.
- `data/1311/meinvoice/`: bốn ZIP nguồn do người dùng cung cấp.
- `support-data-20261113.zip`: snapshot JSON và bản trước khi sửa AMIS.
  `handoff-support-manifest.json` ghi tên, dung lượng, SHA-256 từng thành viên.
- `issued-2026-chain-book-verification.json`: kết quả kiểm công thức giá trị cuối,
  ngày, tiền, màu và bảo toàn 82.207 HĐ; SHA-256 trỏ đúng workbook cuối.
- `docs/cleanup-20261113.json`: danh sách 224 file tạm đã dọn, khoảng 1,79 GiB;
  đã kiểm 638 file nguồn và workbook cuối không đổi.

Giữ nguyên ô Thanh toán nguồn trống. Khi nguồn xác nhận Tiền hàng + VAT = 0,
giá trị cuối tính phần đó là 0 và có ghi chú tại cột Q. Loại khỏi giá trị cuối
HĐ đã bị thay thế, hủy hoặc CQT không hợp lệ. Không dùng tổng cột giá trị cuối
thay cho tổng doanh thu AMIS.

## Tái xuất đúng snapshot đã chốt

Các lệnh dưới chạy tại D:\UVG với Python có lxml, openpyxl, xlsxwriter.
Giải nén snapshot trước; không chạy lại các script điền AMIS chỉ để xuất báo cáo.
Đóng workbook trước khi chạy lệnh xuất đè.

```powershell
Expand-Archive -LiteralPath outputs/nghiem-thu-20261113/support-data-20261113.zip -DestinationPath outputs/nghiem-thu-20261113 -Force
python outputs/nghiem-thu-20261113/prepare_issued_2026_chain_view.py
python outputs/nghiem-thu-20261113/export_issued_2026_streaming.py --chains
python outputs/nghiem-thu-20261113/verify_issued_2026_chain_book.py
```

Xuất theo luồng vì Artifact Tool đã hết bộ nhớ khi lưu bộ dữ liệu lớn.
Các script lịch sử đọc thêm bằng chứng tại thư mục khác trong repo; snapshot
nén là đầu vào đã chốt để tái xuất báo cáo này. `package_handoff.py` cập nhật
gói nén sau khi kiểm chứng một snapshot mới.
