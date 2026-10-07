# UVG – đối chiếu AMIS và meInvoice

Dự án gồm toàn bộ dữ liệu nguồn từ thư mục D:\UVG, các script phân tích và kết quả hiện tại. Chỉ đọc file; không kết nối CQT, sửa dữ liệu nguồn hoặc phát hành hóa đơn.

## Tiếp tục trên máy khác

Đọc **[TIEP_TUC_UVG.md](TIEP_TUC_UVG.md)** để tiếp tục đúng bước. Bảng làm việc ưu tiên ở `outputs/uvg-kiem-tra-uu-tien/HOA_DON_CAN_KIEM_TRA_UU_TIEN.xlsx`, thống kê và cách dùng ở `outputs/uvg-kiem-tra-uu-tien/KIEM_TRA_UU_TIEN.md`.

Cài Python 3.11 trở lên, rồi chạy trong PowerShell:

```powershell
git clone https://github.com/vansang2409/UVG.git
cd UVG
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run.py --validate
.\.venv\Scripts\python.exe run.py
```

`run.py` mặc định đọc dữ liệu trong `data/` của repo, không yêu cầu ổ D trên máy mới. Có thể chọn thư mục khác bằng `--data-dir` và `--output-dir`.

## Cấu trúc

- `data/Du lieu Amis/`: hai sổ chi tiết bán hàng 2024 và tháng 1–6/2025.
- `data/Du lieu Meinvoice/`: toàn bộ chín file meInvoice đã cung cấp, gồm mẫu điều chỉnh.
- `scripts/reconcile_uvg_groups.py`: phương pháp chính, đối chiếu theo nhóm doanh thu, tháng và họ model; không dựa vào mã đơn sàn.
- `outputs/uvg-doi-chieu-nhom/UVG_DOI_CHIEU_NHOM.md`: báo cáo nhóm hiện tại; các CSV cùng thư mục có nguồn và số dòng để truy ngược.
- `outputs/uvg-doi-chieu/`: các kết quả kiểm tra và tìm ứng viên trước đó, giữ để tra cứu. Phương pháp tìm ứng viên chưa chứng minh liên kết từng giao dịch.
- `SOURCE_MANIFEST.json`: tên, dung lượng và SHA-256 của 11 file đầu vào để kiểm tra chuyển máy nguyên vẹn.

Các file `.xls` hiện có thực tế chứa định dạng OOXML; script đọc bằng openpyxl qua luồng file. XLS nhị phân thật cần chuyển đổi riêng.

Mở CSV trong Excel qua Data > From Text/CSV, chọn UTF-8 và giữ số hóa đơn ở kiểu Text để không mất số 0 đầu.

## Căn cứ và trạng thái bàn giao

AMIS do khách hàng cung cấp và đã chốt là số liệu mục tiêu. meInvoice là lịch sử hóa đơn do kế toán cũ phát hành. Chênh lệch chưa tự chứng minh bên nào sai.

AMIS tính doanh thu sau chiết khấu, trả lại và giảm giá. meInvoice loại giá trị hóa đơn đã hủy/đã bị thay thế theo trạng thái file, giữ gốc bị điều chỉnh và cộng khoản điều chỉnh hiện có. Lần theo liên kết đưa khoản sửa về kỳ hóa đơn gốc; kỳ này chưa chắc là kỳ bán hàng AMIS.

Phân bổ model từ dòng có tiền và thuế khi tổng khớp bảng tổng. Dòng thiếu tiền, thiếu lượng và khoản chưa xác định được sản phẩm được giữ riêng. Không tự suy ra biến thể Pro/Light hoặc phân bổ điều chỉnh chung tùy ý.

Đã hoàn thành bảng đối chiếu nhóm, chưa chốt danh sách sửa/phát hành. Cần tiếp tục giải thích hóa đơn thiếu liên kết, tiền hàng cộng thuế lệch tổng thanh toán, phần chưa phân bổ sản phẩm và chênh lệch kỳ. Dữ liệu CQT có thể bổ sung sau; cần xác minh lịch sử đầy đủ trước khi phát hành.

Không dùng chênh lệch tổng để xuất bù hoặc giảm toàn bộ hóa đơn. Phạm vi hiện tại là 2024–tháng 6/2025, chưa bao gồm các lần sửa sau kỳ xuất file.
