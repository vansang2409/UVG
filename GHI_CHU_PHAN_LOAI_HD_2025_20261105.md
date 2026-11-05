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
