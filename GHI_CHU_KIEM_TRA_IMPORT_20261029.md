# Kiểm tra import và dữ liệu meInvoice — 29/10/2026

## Nguồn mới và phạm vi

- `data/Meinv/2024.xls`, 8 phần trong `2025/`, 23 phần trong `2026/` và 2 phần `HD_bi_dieu_chinh_2024_2026/` do người dùng xuất.
- Bộ 2026 chọn kỳ 01/01/2026–29/10/2026; 102.211 hóa đơn, không trùng ký hiệu/số trong bộ năm 2026.
- Người dùng xác nhận kế toán mới import và phát hành nhóm **đã bị điều chỉnh**. Nhóm **chưa điều chỉnh** chưa được xác nhận thực hiện.
- Chỉ đọc dữ liệu, chuẩn bị file; không nhập phần mềm, ký/phát hành hay sửa dữ liệu trên Google Drive trong phiên này.

## Nhóm đã bị điều chỉnh

- So kế hoạch Excel 1.015 chuỗi / 2.045 hóa đơn với bảng kê mới: 1.004 chuỗi khớp số dư 0 ở cả ba khoản; 8 chuỗi còn tiền; 3 chuỗi có tiền hàng/VAT bằng 0 nhưng có ô thanh toán trống.
- Có 2.032 hóa đơn năm 2026 nối vào nhóm kiểm tra. Chưa thấy 13 khoản bổ sung STT 2033–2045 nối vào 8 chuỗi còn dư: 1C24TUV/3, 478, 498, 574; 1C25TUV/40, 1348, 2377, 2378.
- Số dư tiền hàng 233.262.118đ, VAT 18.660.968đ. Một số thanh toán nguồn trống; không mặc định trống là 0 nếu chưa có căn cứ riêng.
- Google Sheets `1ColkEdug6bFDgS09lGDoXz5lPVgudMI5U3kod3TlwUE`, tab `Import dieu chinh` gid293598768, kết thúc STT2032/hàng41799. Link chưa có 13 hóa đơn bổ sung; Excel trên máy có đến2045.
- Script/kết quả: `scripts/audit_issued_adjusted_20261029.py`, `scripts/data/issued-adjusted-20261029/audit.json`. Các chẩn đoán thiếu khoản tiền 0 có thể do ô thanh toán trống, cần đọc cùng chứng cứ XML, không coi là chắc chắn thiếu hóa đơn.

## File bổ sung riêng

- `outputs/bo-sung-13-20261029/IMPORT_DIEU_CHINH_BO_SUNG_13_HD.xlsx`: một sheet `Import dieu chinh`, 24 cột; 13 hóa đơn STT2033–2045, 206 dòng, tham chiếu 8 gốc trên.
- Tổng điều chỉnh: -233.262.118đ tiền hàng, -18.660.968đ VAT, -251.923.086đ thanh toán. Giữ ngày dự thảo19/10/2026 từ bản đã chốt; kế toán kiểm tra ngày trước khi nhập.
- 4.944 ô vùng dữ liệu khớp kế hoạch19/10; số lượng/đơn giá/dòng tiền đối chiếu nguồn đã kiểm hash. Cộng vào các chuỗi hiện có về0; khoản thanh toán gốc574 dùng quy ước0 đã chốt, khoản thông tin6108/4203/4204 dùng XML0.
- STT2034 là khoản tăng để đảo giảm cũ. STT2037 chỉ giảm số lượng, tiền0; không tự loại vì tổng0.
- Hai dòng Excel75/125 của STT2035 giữ thành tiền nguồn, lệch phép nhân số lượng × đơn giá làm tròn1đ mỗi dòng. Không tự tính lại.
- `build.mjs`, `review.py`, `prepared.json`, `verification.json`, `deep-review.json` lưu cùng thư mục. Kiểm tra độc lập đạt; chưa thử import trên meInvoice.
- SHA256 file bàn giao: `e7b8a8b551a7e4f4169421d182ab2c1fd5532308bf1065e19d7906c554436daa`.

## Trạng thái CQT và HĐ1943

- Ảnh người dùng gửi cho thấy 1C25TUV/00001943 ngày29/05/2025, 645.051đ, trạng thái phát hành **TĐ không hợp lệ**. Chưa có thông báo lỗi chi tiết; không xác định nguyên nhân hay xử lý hóa đơn thay kế toán.
- HĐ này có trong bộ **chưa điều chỉnh**, STT1168, hàng5584–5585; giảm597.269đ tiền hàng +47.782đVAT. Không nằm trong file bổ sung13.
- Các bảng kê mới để trống trạng thái gửi CQT cho nhóm có mã1C24TUV/1C25TUV. Trong các gốc có kế hoạch import, chỉ1943 thiếu mãCQT. Thiếu mã là dấu hiệu tra thêm, không phải quy tắc tự kết luận toàn bộ trạng thái hợp lệ/không hợp lệ.
- Ba HĐ1C25MUV/10792,12398,16401 được nguồn ghi rõ HĐ không hợp lệ; không nằm trong hai bộ import.
- Bằng chứng: `scripts/data/invalid-status-20261029/`. Kết quả khớp tiền/liên kết không xác nhận CQT đã chấp nhận; cần kiểm trạng thái riêng trước khi dùng như số phát hành hợp lệ.

## Link bộ chưa điều chỉnh

- Drive `1zcPnKZnvgpA5XRurpjfw862C9Nr6R6iW` là Office XLSX, không phải native Sheets. Tải đúng bản người dùng gửi, giữ snapshot tại `scripts/data/check-untreated-drive-20261029/source.xlsx`; không chuyển đổi hay sửa bản online.
- Đủ8.410 hóa đơn/13.568 dòng; STT1–8410 liên tục; khớp1:1 với tabF0; không trùng gốc, không chồng nhóm đã điều chỉnh và không có quan hệ điều chỉnh mới trong bảng liên kết cho các gốc này.
- So bản cục bộ đã chốt: chỉ bổ sung3.913 ô cộtC tên khách hàng trên1.169 hóa đơn, đều từ ô trống; các cột khác và toàn bộ dữ liệuF0 không đổi. Không phát hiện dư/thiếu **so với kế hoạch đã chốt**, chưa kết luận toàn bộ danh sách đều đủ điều kiện xử lý về thuế.
- Sáu HĐ0đ: 1C24TUV/671,744; 1C25TUV/391,761,835,4748. Bảng tổng mới trống thanh toán, còn tabF0 ghi căn cứ số0 trong chi tiết và xác nhận trước đây; không coi là thiếu dòng hay chênh tiền thật.
- Script/báo cáo tại `scripts/data/check-untreated-drive-20261029/`. Báo cáo `current_money_differences` gồm6 trường hợp trống/0 này; không diễn giải thành6 khoản tiền sai.

## Việc tiếp theo

1. Kế toán xác nhận tình trạng13 khoản bổ sung, kiểm ngày nhập và khả năng phần mềm chấp nhận hóa đơn chỉ giảm số lượng0đ cùng hai dòng làm tròn.
2. Xem phản hồi truyền nhận của HĐ1943 trước khi xử lý STT1168.
3. Khi có phát hành mới, xuất lại bảng kê/liên kết rồi đối chiếu; tách số dư tiền với trạng thái CQT chấp nhận.
