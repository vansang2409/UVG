# UVG – đối chiếu AMIS và meInvoice

Dự án gồm toàn bộ dữ liệu nguồn từ thư mục D:\UVG, các script phân tích và kết quả hiện tại. Chỉ đọc file; không kết nối CQT, sửa dữ liệu nguồn hoặc phát hành hóa đơn.

## Tiếp tục trên máy khác

Đọc **[TIEP_TUC_UVG.md](TIEP_TUC_UVG.md)** để tiếp tục đúng bước. Hai bảng hiện dùng: `outputs/ngoai-le-meinvoice-2024/NGOAI_LE_2024_THEO_CHUOI_GON.xlsx` và `outputs/hoa-don-dieu-chinh-2024/HOA_DON_DA_DIEU_CHINH_2024_BO_SUNG_TEN.xlsx`. Các báo cáo cũ đã được dọn theo yêu cầu người dùng; xem lịch sử Git nếu cần tra cứu.

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
- `scripts/prepare_adjusted_2024.py`: lọc gốc 2024 đã điều chỉnh, loại ngoại lệ và chuỗi thay thế, tính cộng dồn theo nguồn.
- `scripts/layout-tools/build-adjusted.mjs`: xuất bảng điều chỉnh theo format F0/F1/F2 gọn, không thêm bản Excel mới. Cần Node và `@oai/artifact-tool` từ runtime Codex; node_modules không đưa lên Git.
- `SOURCE_MANIFEST.json`: tên, dung lượng và SHA-256 của 11 file đầu vào để kiểm tra chuyển máy nguyên vẹn.

Các file `.xls` hiện có thực tế chứa định dạng OOXML; script đọc bằng openpyxl qua luồng file. XLS nhị phân thật cần chuyển đổi riêng.

Mở CSV trong Excel qua Data > From Text/CSV, chọn UTF-8 và giữ số hóa đơn ở kiểu Text để không mất số 0 đầu.

## Căn cứ và trạng thái bàn giao

AMIS do khách hàng cung cấp và đã chốt là số liệu mục tiêu. meInvoice là lịch sử hóa đơn do kế toán cũ phát hành. Chênh lệch chưa tự chứng minh bên nào sai.

AMIS tính doanh thu sau chiết khấu, trả lại và giảm giá. meInvoice loại giá trị hóa đơn đã hủy/đã bị thay thế theo trạng thái file, giữ gốc bị điều chỉnh và cộng khoản điều chỉnh hiện có. Lần theo liên kết đưa khoản sửa về kỳ hóa đơn gốc; kỳ này chưa chắc là kỳ bán hàng AMIS.

Phân bổ model từ dòng có tiền và thuế khi tổng khớp bảng tổng. Dòng thiếu tiền, thiếu lượng và khoản chưa xác định được sản phẩm được giữ riêng. Không tự suy ra biến thể Pro/Light hoặc phân bổ điều chỉnh chung tùy ý.

Đã hoàn thành bảng đối chiếu nhóm, chưa chốt danh sách sửa/phát hành. Cần tiếp tục giải thích hóa đơn thiếu liên kết, tiền hàng cộng thuế lệch tổng thanh toán, phần chưa phân bổ sản phẩm và chênh lệch kỳ. Dữ liệu CQT có thể bổ sung sau; cần xác minh lịch sử đầy đủ trước khi phát hành.

Không dùng chênh lệch tổng để xuất bù hoặc giảm toàn bộ hóa đơn. Phạm vi hiện tại là 2024–tháng 6/2025, chưa bao gồm các lần sửa sau kỳ xuất file.

## Quy trình năm 2024 đang dùng

- Ngoại lệ: 118 gốc, 190 hóa đơn trong chuỗi. Lọc bằng tên công ty hoặc định danh dạng MST; chưa xác thực MST.
- Nhóm đã điều chỉnh ngoài ngoại lệ: 590 gốc, 1.187 hóa đơn. Không xử lý hóa đơn thay thế trong đợt này.
- 5 chuỗi lệch số học cần kiểm tra. Riêng gốc 1C24TUV/00000574: tạm tính thanh toán trống = 0 theo xác nhận người dùng ngày 07/10/2026; ô nguồn vẫn trống, thanh toán cộng dồn 1.839.000 đồng.
- AMIS dùng cho bước xuất mới sau này, cần loại doanh thu thuộc ngoại lệ giữ nguyên. Chưa tạo import hoặc phát hành.

Tái tính và xuất bảng điều chỉnh:

```powershell
python scripts/prepare_adjusted_2024.py
node scripts/layout-tools/build-adjusted.mjs
```

Đóng file Excel trước khi xuất lại để lưu đè đúng file. JSON hỗ trợ ở scripts/data; ảnh, snapshot, file inspect và thư viện được bỏ qua bởi Git. Lệnh run.py vẫn có thể tái tạo báo cáo nhóm cũ, không phải bước hiện đang xử lý.
