# Bàn giao phân loại hóa đơn năm 2025 — 05/11/2026

Kế toán giữ hóa đơn đúng và xử lý phần sai hoặc thiếu. Bảng phân loại có 65.494 chứng từ AMIS, tổng thanh toán 89.077.087.173đ và VAT 6.598.300.351đ. Đã đọc lại file Excel để kiểm số dòng, mã đơn dạng chữ, tiền, VAT và tổng hợp.

| Việc cần làm | Chứng từ |
|---|---:|
| Đề xuất giữ theo đối soát mã đơn và tiền | 34.560 |
| Kiểm tra và sửa liên kết AMIS–hóa đơn | 112 |
| Rà soát hóa đơn và số dư chuỗi | 327 |
| Ứng viên xuất lại hoặc bổ sung | 29.353 |
| Cần kế toán xác minh thêm | 1.142 |

Bắt đầu với tab **Sửa liên kết**. Khi một nhóm có chứng từ nghi ghép nhầm, sửa liên kết trước rồi đối chiếu lại tiền và VAT. Nhóm giữ cần so chi tiết hàng hóa. Nhóm ứng viên cần xác minh nghiệp vụ, liên kết và các hóa đơn khác trước khi lập. Tab **HĐ cần kiểm riêng** gồm 1.601 HĐ, trong đó có 25 HĐ kế toán đã ghi chú bỏ, chờ đối chiếu bằng chứng xử lý. Tổng ở trên tính theo chứng từ AMIS.

File làm việc tại máy là `outputs/phan-loai-2025-20261105/PHAN_LOAI_XU_LY_HD_2025.xlsx`. Nguồn là ảnh chụp AMIS ngày 05/11/2026, bảng kê meInvoice đến 29/10/2026 và hồ sơ bổ sung ngày 01/11/2026. Dữ liệu nguồn và file Excel chi tiết lưu cục bộ. Bản Google Sheets tạo nhầm đã được chuyển vào thùng rác. Trạng thái công việc là phân loại, chờ kế toán chốt.

Khi có đủ ảnh chụp nguồn tại máy, chạy lần lượt các script trong `scripts/`: `classify_2025_20261105.py`, `prepare_classification_workbook_20261105.py`, `write_classification_streaming_20261105.py`, `verify_classification_xlsx_20261105.py`. Kiểm tra dữ liệu mới nhất trước khi lập hóa đơn.
## Kiểm tra map tiền trên bản sao của kế toán

- Chỉ xử lý các dòng AMIS có đủ số và ký hiệu HĐ nhưng chưa lên tiền. Đã xử lý 478 dòng gồm 421 dòng thiếu tiền ở một bước trong chuỗi Ktoan, 55 dòng có HĐ năm 2026 được bổ sung từ bảng kê meInvoice và 2 dòng cần dò theo HĐ thay thế. Cột L đã được đọc lại sau khi cập nhật.
- Cột O của AMIS ghi trạng thái tại thời điểm kiểm tra. Đã kiểm 35.646 dòng dữ liệu gồm 35.559 dòng đã map tiền, 58 dòng điều chỉnh về 0, 2 dòng thay thế về 0, 1 dòng HĐ gốc 0 đ, 25 dòng thiếu số hoặc ký hiệu và 1 dòng chưa map vì hai số HĐ chung một ô.
- Dòng chưa map vẫn chờ sửa công thức L. Hai HĐ cùng ký hiệu đều có trong Ktoan, tổng tiền 200.000 đ.

## BTL 2025 trong file đối soát

- Tab BTL 2025 có 321 phiếu và 411 dòng chi tiết. Tất cả đã ghép theo Số CT bán hàng lấy trực tiếp từ AMIS Kế toán và khớp mã hàng với chứng từ bán gốc.
- Giá trị trả lại 441.952.027 đ, VAT giảm 35.356.147 đ và thanh toán giảm 477.308.174 đ. Đã đọc lại tab và kiểm tổng. Phần đối chiếu các khoản trả với hóa đơn điều chỉnh trên meInvoice còn cần thực hiện.
