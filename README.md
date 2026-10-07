# UVG – đối chiếu AMIS và meInvoice

Dự án ở `D:\UVG`, gồm dữ liệu nguồn, các script phân tích và hai report hiện tại. Phân tích và xuất report từ file, giữ nguyên dữ liệu nguồn; chưa kết nối CQT, tạo import, ký hoặc phát hành hóa đơn.

Phạm vi chốt ngày 07/10/2026: chỉ chọn F0 có ngày 01/01/2024–30/06/2025 và theo các hóa đơn liên quan đến những F0 này đến 31/12/2025. Không chọn F0 ngoài kỳ hoặc tính các lần xử lý năm 2026. Quy tắc ngoại lệ giữ nguyên; xem trạng thái hiện tại ở đầu `TIEP_TUC_UVG.md`.

Đã làm lại hai report ngày 07/10/2026 từ bộ meInvoice xuất lại đến 31/12/2025, cập nhật manifest 18 file và đối chiếu SHA-256 không đổi. Bố cục theo mẫu bảng đã điều chỉnh người dùng chọn; ngoại lệ có thêm hai cột tên người mua và MST/CCCD nguyên bản ở đầu. `outputs/` chỉ có hai Excel, không có thư mục con.

## Tiếp tục trên máy khác

Đọc **[TIEP_TUC_UVG.md](TIEP_TUC_UVG.md)** để tiếp tục đúng bước. Hai bảng hiện dùng: `outputs/NGOAI_LE_F0_2024_T6_2025.xlsx` và `outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx`. Output cũ đã được dọn theo yêu cầu. Đọc tài liệu và mở Excel thì chưa cần tái chạy phân tích. Các thay đổi cục bộ chưa commit/push sẽ chưa xuất hiện trên máy khác.

Cài Python 3.11 trở lên, rồi chạy trong PowerShell:

```powershell
git clone https://github.com/vansang2409/UVG.git
cd UVG
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe scripts/prepare_invoice_reports.py
```

Script đọc nguồn trong `data/` của repo, không yêu cầu ổ D trên máy mới. Xuất Excel cần Node và `@oai/artifact-tool` từ runtime Codex, xem lệnh phía cuối. Clone cần giữ lịch sử Git để đọc ba file liên kết cũ tại commit `9c50b0424a242695fa31182ab826613bf0fa699f`; không khôi phục các file này ra thư mục nguồn.

## Cấu trúc

- `data/Du lieu Amis/`: hai sổ chi tiết bán hàng 2024 và tháng 1–6/2025.
- `data/Du lieu Meinvoice/`: bộ xuất mới đến 31/12/2025, gồm ba file chi tiết (2024, 2025 thường, 2025 MTT), mười phần bảng kê tổng, hai file bảng điều chỉnh và một file bảng thay thế. Hai bảng điều chỉnh hiện có cùng tập liên kết; không tính trùng.
- `scripts/reconcile_uvg_groups.py`: phương pháp chính, đối chiếu theo nhóm doanh thu, tháng và họ model; không dựa vào mã đơn sàn.
- `scripts/prepare_invoice_reports.py`: đọc bộ nguồn mới, kiểm độ phủ chi tiết/bảng tổng, loại trùng liên kết, bổ sung liên kết lịch sử có provenance, chọn F0 đúng kỳ, lọc ngoại lệ và tính cộng dồn; cập nhật manifest nhưng không sửa Excel nguồn.
- `scripts/layout-tools/build-adjusted.mjs`: xuất cả hai report theo format chuẩn hai hàng tiêu đề F0/F1/F2, thêm F3 khi cần. Cần Node và `@oai/artifact-tool` từ runtime Codex; node_modules không đưa lên Git.
- `SOURCE_MANIFEST.json`: tên, dung lượng và SHA-256 của 18 file hiện tại (2 AMIS, 16 meInvoice).

Các file `.xls` hiện có thực tế chứa định dạng OOXML; script mới đọc XML bằng lxml và đọc các blob lịch sử bằng openpyxl qua luồng file. XLS nhị phân thật cần chuyển đổi riêng.

Mở CSV trong Excel qua Data > From Text/CSV, chọn UTF-8 và giữ số hóa đơn ở kiểu Text để không mất số 0 đầu.

## Căn cứ và trạng thái bàn giao

AMIS do khách hàng cung cấp và đã chốt là số liệu mục tiêu. meInvoice là lịch sử hóa đơn do kế toán cũ phát hành. Chênh lệch chưa tự chứng minh bên nào sai.

AMIS tính doanh thu sau chiết khấu, trả lại và giảm giá. meInvoice loại giá trị hóa đơn đã hủy/đã bị thay thế theo trạng thái file, giữ gốc bị điều chỉnh và cộng khoản điều chỉnh hiện có. Lần theo liên kết đưa khoản sửa về kỳ hóa đơn gốc; kỳ này chưa chắc là kỳ bán hàng AMIS.

Phân bổ model từ dòng có tiền và thuế khi tổng khớp bảng tổng. Dòng thiếu tiền, thiếu lượng và khoản chưa xác định được sản phẩm được giữ riêng. Không tự suy ra biến thể Pro/Light hoặc phân bổ điều chỉnh chung tùy ý.

Các bảng đối chiếu nhóm cũ chỉ là kết quả lịch sử, chưa tái tính AMIS theo bộ meInvoice mới. Hai report hiện tại rà chuỗi hóa đơn trong phạm vi đã chốt, chưa chốt danh sách sửa/phát hành. Cần xác minh ghi chú tiền/ngày và lịch sử đầy đủ trước khi phát hành.

Không dùng chênh lệch tổng để xuất bù hoặc giảm toàn bộ hóa đơn. Chọn F0 trong 01/01/2024–30/06/2025; theo các hóa đơn liên quan đến những F0 này đến 31/12/2025. Trạng thái nguồn là hiện tại; liên kết/ngày chứng từ quyết định phạm vi, không chỉ nhãn đã bị điều chỉnh/thay thế.

## Hai report đang dùng

- Ngoại lệ: 238 F0 (118 năm 2024, 120 tháng 1–6/2025), 326 hóa đơn trong chuỗi. Lọc bằng tên công ty hoặc định danh dạng MST; chưa xác thực MST. Có 2 chuỗi có ghi chú.
- Nhóm đã điều chỉnh ngoài ngoại lệ: 1.047 F0 (591 năm 2024, 456 tháng 1–6/2025), 2.113 hóa đơn. Không xử lý 6 chuỗi thay thế trong đợt này. Hai chuỗi có đủ F0 và ba hóa đơn liên quan được mở thêm khối F3.
- Bảng điều chỉnh có 16 chuỗi có ghi chú: 8 lệch số học, 6 có ô tiền thiếu, 1 quy ước riêng gốc 574 và 1 có ngày khác nhau giữa nguồn. Chi tiết ở cột Cần kiểm tra và tab nguồn.
- Riêng gốc 1C24TUV/00000574: tạm tính thanh toán trống = 0 theo xác nhận người dùng ngày 07/10/2026; ô nguồn vẫn trống, thanh toán cộng dồn 1.839.000 đồng. Không áp dụng quy ước này cho các ô tiền thiếu khác.
- Riêng ngày `1C25TUV/00012885`: bảng chính dùng 05/07/2025 theo xác nhận và ảnh danh sách meInvoice người dùng gửi ngày 07/10/2026. Tab nguồn giữ ngày bảng tổng 18/08/2025; giữ ghi chú lệch ngày và căn cứ xác nhận, chưa đối chiếu XML độc lập. Quy ước được lưu trong script để tái chạy nhất quán.
- AMIS dùng cho bước xuất mới sau này, cần loại doanh thu thuộc ngoại lệ giữ nguyên. Chưa tạo import hoặc phát hành.

Tái tính và xuất cả hai report:

```powershell
python scripts/prepare_invoice_reports.py
node scripts/layout-tools/build-adjusted.mjs
```

Đóng Excel trước khi xuất lại để lưu đè đúng hai file. JSON/ảnh/inspect hỗ trợ ở `scripts/data`; chúng và thư viện được bỏ qua bởi Git. Đã đối chiếu số tiền và dữ liệu nguồn với bảng tổng (ngày bảng chính 12885 áp dụng xác nhận riêng nêu trên), kiểm phạm vi, cộng dồn, ô trống, format, bộ lọc, cố định tiêu đề; 18 hash nguồn không đổi. `run.py` và `prepare_adjusted_2024.py` thuộc workflow nguồn cũ, chưa hỗ trợ bộ xuất mới; không chạy chúng để tiếp tục hai report này.
