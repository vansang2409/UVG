# Tiếp tục UVG

Cập nhật 11/11/2026. Đây là trạng thái cuối đã ghi nhận; đọc lại Google Sheets trước khi sửa. Cập nhật các mục hiện có, không nối nhật ký dài vào file này.

## Việc còn chờ

1. **Hai mã đơn trong lý do BTL chưa sửa lại**. Lượt trước đã lấy mã lỗi của bảng kê meInvoice; chỉ mới điều tra và sửa Người mua hàng sau đó:

| HĐ (đều 1C25MUV) | Phiếu / chứng từ bán | Ô cột N tab BTL | Mã đang ghi → mã đúng đã đối chiếu |
| --- | --- | --- | --- |
| 00023088 | BTL00290 / BHHT2025031831 | N130:N131 | `581284749358761903` → `581234188655821869` |
| 00033304 | BTL00500 / BHHT2025040277 | N407 | `581797491018860000` → `581797491018860286` |

   - **00023088 là HĐ GIỮ LẠI**, tại `HĐ meInvoice giữ lại 2025!C28534:C28535`; thanh toán 3.039.000đ. Bảng kê lặp mã của HĐ23087 cho 23087–23090. MISA_HD_GOC A56823:B56823 và Ktoan D28260 xác nhận mã đúng; chưa xác định thao tác gây lặp.
   - 00033304 thanh toán 1.478.999đ; MISA_HD_GOC A65269:B65269 giữ mã đủ 18 số. Nguồn đã mất đuôi `286` thành `000`, phù hợp lỗi độ chính xác 15 chữ số; chưa chứng minh phát sinh ở Excel hay bước nhập nào.
2. **Kế toán đang kiểm tra BTL/BGG để import**; chưa có xác nhận mới đã import/phát hành. Sau khi KT xử lý, lấy bảng kê/liên kết mới để đối chiếu thực tế. Các Excel/ZIP cũ chưa đồng bộ các sửa mới trên Sheets.
3. Kế toán điền SKU/SL trong tab map sản phẩm; chưa tự suy mã AMIS. Nhóm 6 HĐ làm tròn và 8 HĐ khuyến mại 0đ mới được đối chiếu dữ liệu/tổng tiền, cần KT kiểm XML/bản thể hiện khi chốt nghiệp vụ.

## Bảng đang dùng

- [Xử lý HĐ 2025](https://docs.google.com/spreadsheets/d/1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA/edit) là bản làm việc hiện hành.
- [Bản sao HD 2025 — AMIS/Ktoan](https://docs.google.com/spreadsheets/d/1C5hoPDvZKB1vGFSUG9rzChcORt2L9NnsqETVHxPui9I/edit) dùng đối chiếu liên kết/chứng từ.
- [IMPORT MEINV 2024](https://docs.google.com/spreadsheets/d/1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE/edit) dùng tra bộ 2024 và mẫu lý do BTL.

| Tab trong Xử lý HĐ 2025 | gid | Trạng thái cuối |
| --- | --- | --- |
| TỔNG HỢP | 700000107 | Tổng hợp doanh thu/VAT; xem công thức và phạm vi trên live |
| Rà soát toàn bộ 2025 | 700000010 | Quyết định/chuỗi; không dùng số đếm lịch sử thay trạng thái live |
| Import HD moi | 0 | 30.263 HĐ / 35.042 dòng; số/ký hiệu phát hành tại W:X |
| Dieu chinh ve 0 | 700000101 | 30 HĐ / 38 dòng; số/ký hiệu phát hành tại AD:AE |
| Thay the ve 0 | 700000102 | 1.995 HĐ; số/ký hiệu phát hành tại AA:AB |
| Dieu chinh BTL - cho so HD | 700000103 | 321 phiếu / 410 dòng; J:M có tham chiếu HĐ gốc |
| Dieu chinh BGG - cho so HD | 700000106 | BGG00001 / 1 dòng; J:M có tham chiếu HĐ gốc |
| HĐ meInvoice giữ lại 2025 | 1330675350 | Tra HĐ giữ lại và chi tiết; không mặc định tên trên tab là tên gốc meInvoice |
| HĐ meInvoice xuất mới 2025 | 700000109 | 30.263 HĐ / 35.042 dòng; A:L chi tiết, M:N kết quả/ghi chú |
| Map SP meInvoice giữ lại | 700000110 | 419 tên SP/DV + 44 dòng diễn giải, A5:Q467 |

## BTL/BGG và map sản phẩm đã chốt

- Cột F cả hai tab BTL/BGG: **`Bán hàng cho người tiêu dùng (tên khách)`**, tên theo HĐ gốc, giữ dấu `*`. Đủ 411 dòng; F130:F131 = `Bán hàng cho người tiêu dùng (b***1902)`; F407 = `Bán hàng cho người tiêu dùng (u***fsjnkbm21w)`.
- Cột N BTL: `Điều chỉnh cho hóa đơn số ..., ký hiệu ..., ngày ... do trả lại hàng của mã đơn hàng ...`. BGG cùng phần tham chiếu, ghi giảm giá do hàng trầy xước/rỉ sét. Hai ngoại lệ mã đơn còn chờ ở bảng trên.
- Map sản phẩm đã đủ chi tiết gốc **35.733/35.733 HĐ giữ lại**, gồm công ty và biếu tặng. Đã bổ sung T1/2026 và HĐ `1C26MUV/00070743` ngày 10/08/2026; đã xóa tab **HĐ thiếu chi tiết SP** theo yêu cầu.
- Bố cục map A:Q: STT, tên meInvoice, **mã hàng meInvoice**, ĐVT, loại nội dung, nhóm HĐ, số HĐ xuất hiện, **SKU1/SL1 … SKU5/SL5**. Đã bỏ cột HĐ tham chiếu. SL là số lượng AMIS cho **1 đơn vị hàng/combo meInvoice**; 44 dòng diễn giải không map SKU. Mã nguồn thiếu thì để trống.

## Mốc phát hành đã đối chiếu

- **Bộ 2025, nguồn 11/11:** đủ 30 HĐ điều chỉnh và 1.995 HĐ thay thế về 0 ngày 10/11; bán mới 30.263 HĐ (8.000 ngày 10/11, 22.263 ngày 11/11). 30.249 HĐ khớp trực tiếp, 6 làm tròn đơn giá, 8 khuyến mại 0đ. Tiền hàng **37.067.706.702đ**, VAT **2.965.602.824đ**, thanh toán **40.033.309.526đ**. HĐ `1C26MUV/00152354` 100.000đ ngoài bộ, không cộng vào.
- **Bộ 2024, mốc 01/11:** bán mới 34.739 chứng từ; đã đối chiếu đủ 163 điều chỉnh trả hàng/219 dòng và 2.000 thay thế lô trùng, đều hợp lệ theo bảng chi tiết bổ sung. Không cộng lại lô trùng đã bị thay thế; ô thanh toán nguồn trống vẫn giữ trống.
- Báo cáo 2024 mốc 01/11 còn chênh AMIS **+2.368.960đ** (= HĐ111/khác năm +3.888.889đ và 16 nhóm −1.519.929đ); khi quay lại phải đọc live. Không coi các số dự kiến cũ là kết quả thực phát hành mới.

## Nguồn và bằng chứng

Các đường dẫn dưới tính từ `D:\UVG`; JSON kiểm chứng là mốc đã lưu, không thay cho đọc live.

- `data/1111/`: bảng kê tổng/chi tiết/liên kết mới, gồm các file bổ sung T1/2026 và HĐ00070743.
- `outputs/adjustment-reasons-20261111/`: `order-conflict-investigation.json` (2 mã đơn), `buyers-prefix-verification.json` (F), `verification.json` (lý do N), `original-buyers-source.json` (tên gốc).
- `outputs/map-retained-products-20261111/`: `august-google-readback-verification.json` là kiểm chứng map mới nhất; nguồn đầy đủ = `meinvoice-lines.json` + `added_invoice_details` trong hai file `january-supplement-update.json`, `august-supplement-update.json`. `mapping.json`/`missing.json`/`summary.json` cũ chưa bao gồm bổ sung.
- `outputs/audit-new-sales-detail-20261111/`: đối chiếu chi tiết bán mới; `outputs/fill-issued-zero-20261111/`: số HĐ điều chỉnh/thay thế đã phát hành.
- `outputs/import-2025-20261109/`: Excel/ZIP/CSV chuẩn bị trước phát hành; **không dùng để import lại hoặc coi đã đồng bộ các sửa Sheets mới**. `HUONG_DAN.md` và các tổng/STT cũ có lịch sử nhiều mốc.
- `scripts/data/post-final-detail-20261101/` và `scripts/data/check-google-revenue-report-20261101/`: kiểm chứng bộ/báo cáo 2024; xem archive nếu cần đường dẫn chi tiết.

## Quy tắc tiếp tục

- AMIS là sổ kế toán mục tiêu; meInvoice là lịch sử HĐ. Ghép **ký hiệu + số 8 chữ số**, chứng từ và bằng chứng hàng/tiền; mã đơn sàn có thể lặp hoặc mất số, không dùng riêng để tự đổi liên kết.
- Giữ nguồn, ô trống, số 0 đầu và tên che dấu; không tự cân tiền. Giữ SKU/SL, công thức, ghi chú KT và bố cục đang dùng; sửa đúng phạm vi được yêu cầu.
- Phân biệt dự thảo, import, phát hành và kiểm chứng XML/CQT. Không tự import/ký/phát hành. Mọi kết quả mới cần ghi rõ nguồn và phạm vi đã kiểm.
- Trước Git: xem `git status`, giữ thay đổi local; chỉ pull/commit/push khi được yêu cầu. Bàn giao ghi chú riêng với dữ liệu/Excel local, không tự gộp mọi file thay đổi vào commit.

Lịch sử đầy đủ trước compact: [bản lưu 11/11/2026](docs/archive/TIEP_TUC_UVG_20261111_TRUOC_COMPACT.md). Chỉ mở khi cần ca cũ; không áp dụng lại các mục chờ đã được cập nhật về sau.
