# UVG – tiếp tục ở nhà, cập nhật 11/10/2026

Repo: https://github.com/vansang2409/UVG.git – nhánh master.

Cập nhật 12/10/2026: đã áp dụng đủ 6 điểm kế toán chốt vào cả 3 workbook chính, bổ sung 48 hóa đơn/547 dòng trên 27 gốc. Tổng hiện tại: 10.442 hóa đơn dự kiến/55.358 dòng. Cả 19 chuỗi chờ và 11 chuỗi cần chốt đã có hướng xử lý hoặc giữ nguyên; không yêu cầu trả lời lại. Đã nhận đủ 20/20 XML điều chỉnh của nhóm 19 chuỗi và giữ ZIP/XML/phản hồi gốc trong `data/Bo sung/`. Chưa nhập phần mềm, ký hoặc phát hành. Trạng thái hiện tại nằm ở đầu `TIEP_TUC_UVG.md`.

## Nhận bản mới

Nếu đã clone, chạy trong thư mục UVG:

```powershell
git status --short
git pull --ff-only origin master
git log -1 --oneline
```

Nếu có thay đổi cục bộ, nhờ Codex kiểm tra và bảo toàn trước khi pull; không reset/xóa thay đổi. Nếu chưa clone:

```powershell
git clone https://github.com/vansang2409/UVG.git
cd UVG
```

## Cho Codex tiếp tục đúng ngữ cảnh

Dán đoạn này vào chat trong project UVG:

> Đọc TIEP_TUC_UVG.md từ đầu và HUONG_DAN_MAY_O_NHA.md; ưu tiên cập nhật 12/10/2026 hơn các mục lịch sử. AMIS là mục tiêu, meInvoice là lịch sử; F0 thuộc 01/01/2024–30/06/2025. Đã chốt và áp dụng đủ 6 điểm cùng dấu 5 cặp giá âm. Import hiện tại: chưa điều chỉnh 8.410 hóa đơn/13.568 dòng; đã điều chỉnh 2.032 hóa đơn/41.790 dòng trên 1.007 chuỗi. Ngoại lệ có 244 chuỗi/338 hóa đơn liên quan; thêm 3 hồ sơ cùng sản phẩm giữ nguyên riêng. Đã rà đủ 9.695 gốc = 9.417 gốc chuẩn bị import + 247 gốc giữ nguyên + 31 gốc đã khớp. Không hỏi lại 19 chuỗi chờ/11 chuỗi cần chốt. Giữ nguồn, số 0 đầu, ô trống; đảo đúng số đã phát hành ở 8 khoản lệch. Chưa nhập/ký/phát hành. Kiểm hiện tại bằng scripts/data/apply-confirmed-20261012/verify.py --final, không chạy builder/verifier lịch sử ghi đè workbook. Đọc bước kiểm nháp phần mềm và tiếp tục theo yêu cầu tôi.

## File chính

- outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx
- outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx
- outputs/NGOAI_LE_F0_2024_T6_2025.xlsx
- outputs/kt-review-20261011-01a12992/UVG_PHUONG_AN_DA_CHOT_20261011.xlsx (phương án đã chốt và phản hồi gốc).
- outputs/kt-review-20261011-01a12992/UVG_CAC_VAN_DE_CAN_KE_TOAN_XAC_NHAN_20261011.xlsx (bản câu hỏi ban đầu, dùng đối chiếu lịch sử).

Chỉ mở file/tài liệu thì không cần tái tạo Excel. Giữ bản trước cập nhật ở `scripts/data/apply-confirmed-20261012/before/` để so sánh độc lập; ảnh xem trước, cờ duyệt và bản `staged/` chỉ giữ cục bộ. Dùng kiểm tra hiện tại:

```powershell
python scripts/data/apply-confirmed-20261012/verify.py --final
```

Cần Python 3.11+ với openpyxl theo requirements.txt; có thể dùng runtime Codex từ load_workspace_dependencies. Script tự xác định checkout, kiểm hash nguồn/phương án/bản trước cập nhật, đọc cả 3 workbook và ghi `final-verification.json`; không ghi Excel. Các audit cũ có số cố định 8.404/1.990 mô tả trạng thái trước bổ sung, không dùng để kết luận bộ hiện tại bị sai hoặc để tái tạo đè lên workbook. Giữ lịch sử Git để kiểm liên kết thay thế60→61 từ commit9c50b042. Đây là kiểm tra dữ liệu, chưa xác minh mật mã chữ ký hoặc trạng thái CQT trực tiếp.

Phần import cũ giữ ngày dự thảo 08/10 và 09/10/2026; 48 hóa đơn mới dùng 12/10/2026. Không tự đổi theo ngày máy. Bước tiếp theo trên meInvoice là nhập nháp, kiểm 8 khoản đảo lệch tổng tiền được tô vàng, tính chất dòng diễn giải, ngày và tham chiếu; xem kỹ bản nháp trước khi ký/phát hành. Không chạy lại prepare/apply/promote chỉ để kiểm tra.
