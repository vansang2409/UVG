# Rà liên kết mã đơn AMIS – meInvoice – nội bộ — 21/10/2026

Người dùng xác nhận mã đơn trên meInvoice và Đối soát NB do nội bộ ghi chú, có thể sai. Không coi chúng là khóa đã được xác minh. Chỉ đọc Google Sheets, không sửa mã, công thức, dữ liệu nguồn hoặc hai bộ import.

Nguồn: Google Sheets `HD 2025`, ID `1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw`. Đã đọc đủ 65.495 dòng AMIS, 37.768 hóa đơn meInvoice, 26.550 dòng NB và65.436 dòng ánh xạ MISA_HD_GOC (không tính tiêu đề). Đọc thêm các tab review, gồm tab kế toán check mã lặp. Lưu snapshot các vùng giá trị để tái kiểm tại source/ cùng thư mục (cục bộ, không tự đưa vào Git).

## Nguyên tắc kiểm

- Nối NB với meInvoice theo số hóa đơn + ngày, kiểm tiếp người mua và từng khoản tiền hàng/thuế/thanh toán. NB không có ký hiệu; kiểm có duy nhất một hóa đơn phù hợp trước khi dùng làm bằng chứng.
- Đối chiếu mã note với chứng từ AMIS; kiểm thuế và thanh toán AMIS, ngày hạch toán, số chứng từ; không tin kết quả MATCH/INDEX đầu tiên khi có nhiều chứng từ.
- Chuẩn hóa ngày serial và ngày chữ `05-Jul-2025` đúng nghĩa. Không tự sửa ngày malformed `2709/2025` thành một ngày giả định.
- So sánh từng chỉ tiêu tiền sau ROUND_HALF_UP về đồng, dung sai phân loại tối đa1đồng/chỉ tiêu. Không thay đổi tiền nguồn. Ghi chú trong sheet không được dùng làm lệnh thay đổi dữ liệu.
- Không cộng các dòng NB hoặc AMIS để ép khớp vì chưa chứng minh chúng là chi tiết độc lập hay ghi lặp.
- Giữ trạng thái hóa đơn mới/điều chỉnh/thay thế riêng; không gọi khoản đảo khác tiền AMIS là lỗi mã đơn.
- 96 dòng của48nhóm mã lặp đã có KT ghi chú; đã giữ nguyên và kiểm số hóa đơn trong ghi chú khớp dòng nguồn. Không coi các ghi chú này là lệnh tự hủy/điều chỉnh.

## Kết quả chính

- 23.938 hóa đơn có các trường đối soát phù hợp ở cả3nguồn. Đây là mức khớp dữ liệu, không xác nhận tuyệt đối mã note đúng đơn thực tế.
- 276 hóa đơn **mới** có mã note tìm thấy trên AMIS nhưng thuế và/hoặc thanh toán không khớp. Chưa kết luận sai mã hay sai tiền, cần kiểm chứng từ gốc. Có800khoản lệch nếu tính mọi trạng thái, trong đó484là hóa đơn điều chỉnh; không đánh đồng800là lỗi.
- 629 mã note xuất hiện ở nhiều hóa đơn meInvoice,592 mã xuất hiện ở nhiều số/ngày hóa đơn NB. Có thể là xuất trùng hoặc chuỗi điều chỉnh/thay thế, không tự loại.
- Khi kiểm số/ngày/người mua/tiền trước:1.619hóađơn chỉ có NB cùng số nhưng không khớp ngày;114không khớp người mua;12không khớp tiền sau làm tròn;5khớp bộ trường nhưng còn nhiều mã NB. Bốn nhóm này không trùng nhau (tổng1.750). NB cùng số có thể thuộc ký hiệu khác, không tự kết luận ngày ghi sai.
- 11.398 hóa đơn không có số tương ứng trong NB;435khớp bộ trường NB nhưng chưa có mã tương ứng trên AMIS;247khớp NB nhưng tiền AMIS chưa khớp. Đây là tình trạng phủ/khớp dữ liệu, không phải tất cả đều là mã note sai.
- NB có615dòng ngày hóa đơn chưa đọc thành ngày hợp lệ:577trống và38ghi `2709/2025`; không tự đoán.

## Ví dụ chưa khớp

`1C25TUV/00011305`,07/07/2025, dòng meInvoice1237: mã note`250707F76E0BV1`, người mua`tranhoatqh`, thanh toán1.379.000đ. NB dòng377khớp hóa đơn/tên/bộ tiền và cũng ghi mã này. Mã đang trỏ chứng từ AMIS`BHHS2025010663`, dòng34508, thanh toán2.479.000đ,thuế183.630đ; hóa đơn chỉ thuế102.148đ. Chưa dùng ghi chú chung mã để khẳng định liên kết đúng.

## Nhóm nhận trước 01/07 nhưng hóa đơn từ 01/07/2025

- **116ứngviên khớp đối soát ba nguồn**, nhận hàng NB trước01/07, AMIS hạch toán trước01/07, hóa đơn mới từ01/07. Không có mã nào trong nhóm này bị dùng lại ở một hóa đơn khác trong bảng meInvoice07–12. Có35người mua là tên che một phần, cần giữ mức khẳng định là khớp đối soát, không coi là chứng từ giao hàng gốc.
- Số sơ bộ114được nâng thành116sau khi đọc đúng ngày chữ `05-Jul-2025` ở NB231–232: hóa đơn11052 và11053. Đây là chuẩn hóa khi đọc, không sửa sheet.
- **1trườnghợp chưa đủ AMIS**: `1C25TUV/00010101` ngày01/07/2025, NB2395ghi nhận hàng30/06/2025, nhưng mã`579400935762658864` chưa tìm được trên AMIS/MISA_HD_GOC. Không tự chọn đơn có cùng số tiền.
- Tổng117ứngviên, chưa đưa thêm vào import; không gọi116là tổng toàn bộ đơn2024–T6/2025 vì AMIS hiện tại chỉ có ngày hạch toán2025 và NB chưa chứng minh bao phủ mọi đơn.

Báo cáo chi tiết máy đọc: audit.json; số tổng: summary.json; script: scripts/audit_order_links_20261021.py. Mã note, ngày và số tiền nguồn giữ nguyên. Script cần snapshot local hoặc đọc lại Google Sheets theo source-manifest.json để tái kiểm trên máy khác. Không có thao tác ghi Google Sheets/import/ký/phát hành.
