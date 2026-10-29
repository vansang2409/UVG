# Kiểm tra nguồn lập chuỗi F0 nửa cuối 2025 — 22/10/2026

## File đã hoàn tất

`outputs/HOA_DON_GOC_T7_T12_2025_F0_F1_F2.xlsx`, sheet `F0, F1, F2, ...` (37.087 gốc) và `Chi tiet nguon` (37.758 hóa đơn). Lọc cộtF của sheet chính khác trống để xem436 chuỗi có thanh toán trống; ô thanh toán nguồn vẫn trống và được tô vàng, không tự thay0. Ba cột tiền hiện tại sử dụng công thức; loại khoản đã bị thay thế, cộng các điều chỉnh còn lại, giữ tổng thanh toán trống nếu thiếu khoản đang dùng. Đã kiểm độc lập toàn bộ các dòng/khối F/số tiền/cache công thức; xem `workbook-verification.json`. Không tạo hóa đơn import hoặc ký/phát hành.

Phương thức xuất: đã tạo/tính/render bằng Artifact Tool; bước xuất lớn không hoàn tất sau thời gian chờ kéo dài. Đã dừng đúng exporter riêng và dùng bundled XlsxWriter ghi cuối với cùng cấu trúc/công thức/cache, kiểm lại bằng openpyxl chỉ đọc. Script ghi cuối: `scripts/build_h2_workbook_20261022.py`; script kiểm: `scripts/verify_h2_workbook_20261022.py`. Không tự ghi đè đầu ra hiện có.

## Cập nhật sau ảnh người dùng gửi ngày 22/10

Đã xác nhận đoạn thiếu `1C25MUV/00031054` (19/12/2025) -> `1C25MUV/00034893` (31/12/2025), biên bản BBTT00109, từ hai ảnh màn hình lưu trong `data/Bo sung/2026-10-22/`. 34893 tiếp tục bị thay thế bởi `1C26MUV/00002639` ngày13/01/2026 theo báo cáo thay thế 2026 dòng14. Bảng tổng 2026 phần1 dòng2644 xác nhận 2639 là Hóa đơn thay thế, HĐ hợp lệ, tiền hàng0/thuế0/thanh toán trống. Không tự đổi thanh toán thành0.

Đã tái chạy toàn bộ phép ghép có nguồn chứng cứ UI riêng: còn **37.087 gốc**, **662 chuỗi có liên quan**, **671 hóa đơn liên quan**, **37.758 hóa đơn trong quần thể chuỗi**; không còn trường hợp trạng thái đã bị điều chỉnh/thay thế nhưng thiếu cạnh. 34893 chuyển từ gốc tạm sang F1 của31054; 2639 là F2. Vẫn436 hóa đơn có thanh toán trống; mọi nguồn tiền/ngày đã ghép không xung đột. Đủ để dựng bảng chuỗi theo nguồn hiện có, giữ trống và đánh dấu tổng thanh toán thiếu. Các số liệu/kết luận trước ảnh ở phần dưới là lịch sử, được cập nhật bởi đoạn này.

---

Đã ghép bảng tổng xuất lại 09/10 (2024, cả năm 2025, 2026 đến 09/10), liên kết điều chỉnh 2024–2025, thay thế 2025, và hai báo cáo liên kết 2026 xuất đến 22/10 vừa tải. Nguồn nguyên bản không thay đổi; chưa tạo workbook mới, chưa sửa import, chưa nhập/ký/phát hành.

- 38.019 hóa đơn có ngày trong 01/07–31/12/2025 trên bảng tổng Excel.
- 37.088 gốc sau khi loại vai trò con trong đồ thị liên kết; gồm 36.428 mới, 549 đã bị điều chỉnh, 111 đã bị thay thế. Đây là quần thể theo nguồn hiện có, chưa loại công ty/ngoại lệ theo quyết định xử lý.
- 662 gốc có ít nhất một hóa đơn liên quan, tổng 670 hóa đơn liên quan: 509 điều chỉnh + 101 thay thế năm 2025; 51 điều chỉnh + 9 thay thế năm 2026. Tổng quần thể chuỗi 37.758 hóa đơn; chuỗi nhiều nhất hiện tìm thấy 3 hóa đơn.
- Tất cả hóa đơn đã nối đều có bảng tổng; không thấy xung đột ngày/ba khoản tiền giữa các nguồn đã ghép, không thấy đa cha hoặc chu trình.
- Một gốc `1C25MUV/00031054`, 19/12/2025, trạng thái **Hóa đơn đã bị thay thế** nhưng không có cạnh đi ra trong báo cáo thay thế 2025 hoặc hai bản 2026. Tiền hàng 1.369.444, thuế 109.555, thanh toán 1.478.999. Cần mở lịch sử hóa đơn trên meInvoice để tìm hóa đơn thay thế thực tế; không tự chọn theo mã đơn hoặc số tiền. Vì vậy chưa xác nhận đủ toàn bộ F liên quan.
- 436 hóa đơn có tiền hàng 0, thuế 0, thanh toán trống: 1 gốc và 435 hóa đơn liên quan. Giữ trống nguồn, không tự chốt thành 0; khi tính cộng dồn cần trình bày thiếu thanh toán riêng. Không cản việc lập bảng nhận diện chuỗi nhưng cản kết luận mọi tổng đã đầy đủ.
- Snapshot Google Sheets ngày 21/10 có 37.768 hóa đơn H2, ít hơn bảng tổng Excel 251 hóa đơn; 251 khóa này là gốc trên Excel, không tự bỏ chỉ vì Google không có. Ví dụ `1C25MUV/00003202`, `00003203`, 30/07/2025, trạng thái mới, gửi CQT HĐ hợp lệ. Chưa kết luận nguyên nhân thiếu ở Google.

Kết luận: đủ nguồn để dựng bản F0/F1/F2 có đánh dấu thiếu dữ liệu, chưa đủ để chốt bản đầy đủ tất cả liên kết và tổng thanh toán. Không cần xuất lại toàn bộ báo cáo 2025. Cần kiểm tra lịch sử thay thế của 31054 trước. Độ mới bảng tổng 2026 là 09/10; báo cáo liên kết đến 22/10 không thay thế được việc xác minh các thay đổi trạng thái khác sau 09/10 nếu có.

Chi tiết theo khóa/nguồn/dòng: `audit.json`; tổng hợp: `summary.json`; script tái chạy: `scripts/audit_h2_chain_readiness_20261022.py`.
