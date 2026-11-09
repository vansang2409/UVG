# Kết quả kiểm tra bàn giao — 09/11/2026

Đã đối chiếu dữ liệu từ hàng tiêu đề đến hết 5 tab import với từng file Excel. Kiểm STT liên tục, vùng trống cuối tab và tiền hàng/VAT khớp báo cáo. BTL00449 và BTL00488 dùng số trên phiếu AMIS và chờ số HĐ bán mới theo trạng thái đã ghi trên Google Sheets. File Excel BTL, báo cáo Excel và CSV liên kết đã được đồng bộ.

| Tab import | HĐ hoặc phiếu | Dòng dữ liệu |
| --- | ---: | ---: |
| Import HD moi | 30.265 | 35.047 |
| Dieu chinh ve 0 | 27 | 35 |
| Thay the ve 0 | 1.989 | 1.989 |
| Dieu chinh BTL - cho so HD | 321 | 410 |
| Dieu chinh BGG - cho so HD | 1 | 1 |

Bộ `BO_IMPORT_HD_2025_20261109.zip` gồm bốn file import và báo cáo Excel. `IMPORT_DIEU_CHINH_BGG00001_CHO_SO_HD.xlsx` nằm cạnh bộ ZIP. Điền ký hiệu, số, ngày và mã CQT của HĐ tham chiếu trong cột J:M trước khi nhập dự thảo BTL hoặc BGG. BTL có 242 phiếu chờ số HĐ bán mới và 79 phiếu chờ kiểm tra việc trả hàng trên HĐ giữ lại.

Tab **Tổng hợp HĐ 2025** trên [Google Sheets](https://docs.google.com/spreadsheets/d/1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA/edit?gid=700000107#gid=700000107) hiển thị doanh thu thuần toàn bộ AMIS là 82.039.782.270đ chưa VAT. Từ tổng bán hàng 82.485.269.841đ, trừ 444.205.449đ của 322 phiếu trả hàng và 1.282.122đ giảm giá. Phạm vi xử lý meInvoice đối chiếu với AMIS chênh 0đ cả tiền hàng và VAT. Bằng chứng kiểm tra tại `final-import-tab-audit-20261109.json`.

---

# Lấy VAT bán mới theo AMIS 09/11 — 09/11/2026

Đã cập nhật bốn bản nháp PSC2 theo bảng bán hàng và sổ chi tiết AMIS 09/11. VAT của `BHHT2025034835`, `BHHT2025032283`, `BHHS2025018941` và `BHHT2025027227` lần lượt là 18.536đ, 195.481đ, 212.480đ và 48.033đ. File bán mới ghi tổng VAT ở dòng đầu của mỗi HĐ. Sáu chứng từ PSC2 khớp đủ tiền hàng, VAT và thanh toán với bản AMIS mới.

Phương án dự kiến khớp AMIS trong phạm vi báo cáo, gồm 80.418.470.211đ tiền hàng và 6.433.663.120đ VAT. Xem mục `amis_0911_import_vat` trong `final-summary.json` và `amis-0911-import-vat-google-verification.json` để truy dòng nguồn và STT.

---

# Giữ HĐ 00020784 cho BHDV20250074 — 09/11/2026

Kế toán đã chốt giữ `1C25MUV/00020784` cho `BHDV20250074`. HĐ khớp ngày, người mua và số tiền 150.000đ gồm VAT 11.111đ. Mã trên AMIS là `HD009935`, mã trên HĐ meInvoice là `HD009798`. `BHDV20250066` giữ HĐ `1C25MUV/00024410`.

Đã đọc lại liên kết và số tổng hợp trên Google Sheets, cập nhật báo cáo Excel và ZIP. Tiền hàng dự kiến khớp AMIS 09/11 ở 80.418.470.211đ. VAT dự kiến thấp hơn 125.759đ do bốn bản nháp tính lại VAT 8% sau khi trừ PSC2. Xem `confirmed-service-keep-20261109-google-verification.json`.

---

# Giữ thêm 5 HĐ khớp AMIS — 09/11/2026

Giữ `1C25MUV/00002368` cho `BHHT2025019975`, mã đơn `579734739357893944`. Mã đơn và tên người mua phân biệt với `00002367` của `BHHT2025019976`. Rút `00002368` khỏi lô thay thế về 0. Chuyển HĐ trùng `00012392` từ STT 1.990 sang STT 608, còn 1.989 HĐ thay thế.

Cộng lại bốn HĐ đã khớp `00012243`, `00012247`, `00012251`, `00012273`. Phương án dự kiến tăng 4.732.682đ tiền hàng và 378.613đ VAT. Sau cập nhật còn thấp hơn AMIS 138.889đ tiền hàng và 136.870đ VAT. `BHDV20250074` còn chờ xác minh HĐ riêng, tương ứng 138.889đ tiền hàng và 11.111đ VAT. Phần 125.759đ VAT còn lại thuộc bốn bản nháp tính lại 8% sau khi trừ PSC2.

Đã kiểm bộ Excel, ZIP và đọc lại số tổng hợp, chi tiết, STT trên Google Sheets. Xem `verified-keeps-20261109-google-verification.json` và mục `verified_keeps_20261109` trong `final-summary.json`.

---

# Tách tổng hợp và chi tiết chuỗi HĐ 2025 — 09/11/2026

Tab `Tổng hợp HĐ 2025` (gid 700000107) hiển thị các tổng tiền hàng, VAT và chênh lệch. Các ô lấy trực tiếp từ báo cáo nguồn. Tab `Ra soat chuoi 2025` (gid 700000010) hiển thị danh sách chi tiết, giữ nguyên số dòng và bộ lọc. Dòng tổng hợp 1–13 ở tab nguồn được ẩn để duy trì công thức và các liên kết hiện có. Đã kiểm 46 ô tổng hợp và hiển thị hai tab trên Google Sheets.

---

# Thay thế HĐ xuất trùng 00012392 về 0 — 09/11/2026

Đã chốt giữ `1C25MUV/00012263` cho `BHHT2025027643`, mã đơn `580325498726155292`. Theo xác nhận người dùng, bổ sung thay thế HĐ trùng `1C25MUV/00012392` về 0 ở STT **1.990**, số dư gốc 1.616.713đ tiền hàng, 129.337đ VAT, tổng 1.746.050đ. Hai HĐ cùng ngày, người mua và chi tiết 1 KA350 Light. Mã đơn HĐ trùng trống trên bảng kê gốc, đối chiếu qua NB và HĐ đang giữ.

Đã cập nhật Google Sheets, báo cáo Excel, CSV và ZIP. Xác minh dòng thay thế và trạng thái giữ, BHH không có dòng bán mới. Tổng phương án dự kiến không đổi. Căn cứ `confirmed_duplicate_12392` trong `final-summary.json` và `confirmed-duplicate-12392-google-verification.json`.

---

# Chốt thay thế và bán mới BHHT2025027502 — 09/11/2026

Đã chốt theo yêu cầu người dùng. Thay thế HĐ `1C25MUV/00012497` về 0 ở STT **1.230**, nối trực tiếp với `BHHT2025027502`. Dự thảo bán mới STT **29.640** theo AMIS 09/11 gồm 1 máy KA150P, tiền hàng **1.276.767đ**, VAT **102.141đ**, thanh toán **1.378.908đ**. HĐ cũ `00012497` có 655.908đ, chênh 723.000đ. HĐ `1C25MUV/00012374` thuộc chứng từ giữ `BHHT2025027561`.

Đã bỏ ghi chú chờ chốt, đồng bộ Google Sheets, Excel, CSV liên kết và ZIP. Đã kiểm cả hai STT, số tiền và báo cáo sau xử lý. Bộ file ở trạng thái dự thảo. Căn cứ `confirmed_reissue_7502` trong `final-summary.json` và `confirmed-reissue-7502-google-verification.json`.

---

# Cập nhật mốc đối chiếu theo AMIS 09/11 — 09/11/2026

Báo cáo và CSV đối soát dùng hai file AMIS trong `data/0911/`. Bảng bán hàng có 65.492 chứng từ. `BHDV20250045` đã xóa, hai BHDV sửa số lượng lên 2 khớp dự thảo. Sáu chứng từ đã bỏ dòng PSC2 trên AMIS, giảm tổng tiền hàng 2.351.000đ. Báo cáo lấy trực tiếp số mới, phần chênh do tính lại VAT 8% cho bốn HĐ là 125.759đ.

AMIS cùng phạm vi có **80.418.470.211đ tiền hàng** và **6.433.663.120đ VAT**. Phương án dự kiến **80.413.598.640đ tiền hàng** và **6.433.147.637đ VAT**, chênh **−4.871.571đ** và **−515.483đ**. Đã cập nhật Google Sheets, Excel và ZIP, kiểm hash bốn file import. Căn cứ `amis_reporting_0911` trong `final-summary.json` và `amis-0911-reporting-google-verification.json`.

---

# Giữ HĐ 00012515 cho BHHT2025027490 — 09/11/2026

Đã sửa liên kết sang `1C25MUV/00012515` cho `BHHT2025027490`, mã đơn `580369342577804306`. HĐ khớp 8 KH200, tiền hàng 4.659.259đ, VAT 372.741đ và thanh toán 5.032.000đ. Đã cập nhật AMIS, AMIS 1, Đối soát NB, Ktoan và liên kết ngược meInvoice trên Bản sao HD 2025. Cột mã đơn gốc trên bảng kê meInvoice vẫn giữ nguyên dữ liệu nguồn. Gỡ `1C25MUV/00012392` và theo dõi riêng để xác định chứng từ thực tế.

Đã đọc lại 25 ô cập nhật trên hai bảng, xác minh số lệch của BHH bằng 0 và kiểm tra Excel/ZIP. Còn **0 chứng từ AMIS chờ** trong danh sách này. Bốn file import giữ nguyên nội dung vì HĐ `00012515` đã được rút khỏi dự thảo thay thế ở lượt trước. Tổng dự kiến 80.413.598.640đ tiền hàng và 6.433.147.637đ VAT. Chênh so AMIS −7.222.571đ tiền hàng và −515.483đ VAT. Căn cứ `confirmed_keep_12515_20261109` trong `final-summary.json` và `confirmed-keep-12515-google-verification.json`.

---

# Xuất lại 5 BHH và xác minh liên kết BHH cuối — 09/11/2026

Đã lập thay thế về 0 cho 4 F0 của `BHHT2025038444`, `BHHT2025038427`, `BHHT2025038199` và `BHHT2025029713`. STT thay thế lần lượt 1.248, 1.987, 1.988, 1.989. STT bán mới 30.261–30.264, giữ số tiền AMIS.

`BHHT2025027227` dùng quy trình điều chỉnh 2024. STT 26 đảo F1 `1C25MUV/00017218` tăng 109.000đ. STT 27 giảm toàn bộ F0 `1C25MUV/00013149` 648.444đ. Hai dự thảo cùng tham chiếu F0 và chép riêng đúng Tên khách hàng/Người mua hàng F0. Số dư chuỗi từ 539.444đ về 0. STT bán mới 30.265 bỏ PSC2 109.000đ, tiền hàng 491.411đ, VAT 8% là 39.313đ, thanh toán 530.724đ.

`BHHT2025027490` vẫn chờ chốt liên kết. Mã đơn `580369342577804306` ở NB dòng 15512 giúp xác định ứng viên `1C25MUV/00012515` khớp người mua, 8 KH200 và đủ 3 khoản, thanh toán 5.032.000đ. Mã đơn trống trên HĐ gốc và không tìm được trực tiếp trong 591 file meInvoice. Liên kết hiện hành `1C25MUV/00012392` khác người mua và hàng. Tạm rút `00012515` khỏi lô thay thế STT 1.248 trong khi chờ chốt.

Đã kiểm Excel, ZIP và đọc lại 463 ô cập nhật trên Google Sheets. Tổng hiện hành **30.265 HĐ bán mới / 35.047 dòng**, **1.989 HĐ thay thế**, **27 HĐ điều chỉnh / 19 chuỗi** và **1 BHH chờ**. Báo cáo dự kiến 80.408.939.381đ tiền hàng và 6.432.774.896đ VAT. Căn cứ `final_pending_reissues_20261109` trong `final-summary.json` và `final-pending-reissues-google-verification.json`.

Kiểm bổ sung bản AMIS 09/11 ở `data/0911/`. Bốn BHH thay thế khớp các dòng tiền đã dùng. `BHHT2025027227` đã bỏ dòng PSC2 trong bản 09/11, còn 491.411đ tiền hàng và VAT nguồn 48.033đ. Dự thảo giữ đúng tiền hàng và tính lại VAT 8% là 39.313đ. Sáu BHH đều để trống ký hiệu và số HĐ ở sổ chi tiết AMIS. Tổng đối chiếu năm trong báo cáo vẫn theo bản trích đã dùng để lập bộ 2025.

---

# Giữ ba HĐ KOL/KOC và loại một chứng từ AMIS đã xóa — 09/11/2026

Giữ `1C25TUV/00013158`, `1C25TUV/00013177` và `1C25TUV/00013178` theo danh sách biếu tặng bổ sung và xác nhận của người dùng. Rút khỏi dự thảo thay thế về 0. Theo dõi riêng ở cột O của tab `Ra soat chuoi 2025` và tab Excel `MI2025 chua map`. Tổng tiền hàng 70.901.909đ, VAT 5.672.153đ, thanh toán 76.574.062đ.

`BHDV20250045` đã xóa trên AMIS theo xác nhận người dùng. Ghi dấu vết trong CSV đối soát và rút khỏi danh sách chờ. Hiện còn **1.986 HĐ thay thế về 0**, **30.260 HĐ bán mới / 35.037 dòng** và **6 chứng từ AMIS chờ**. Tổng dự kiến 80.398.644.281đ tiền hàng và 6.431.951.289đ VAT.

Đã kiểm lại 190 ô cập nhật và đủ 1.986 mã HĐ trên Sheets, đối chiếu năm Excel và năm thành viên ZIP. Bằng chứng tại `gift_keeps_and_amis_deletion_20261109` trong `final-summary.json` và `gift-keeps-and-amis-deletion-google-verification.json`.

---

# Bán mới một BHH đã về 0 và giữ HĐ 00005776 — 09/11/2026

Đã thêm `BHHT2025035946` vào dự thảo bán mới ở **STT 30.260**, dòng 35046, 1 máy KA350 Standard. Tiền hàng **2.073.147đ**, VAT **165.852đ**, thanh toán **2.238.999đ**. Chuỗi `1C25MUV/00031054 → 1C25MUV/00034893 → 1C26MUV/00002639` có số dư 0 ở cả ba khoản theo dòng 30482 tab Ktoan. Báo cáo đã nối đủ chuỗi và gộp phần theo dõi của HĐ giữa chuỗi. Không thêm lô điều chỉnh/thay thế về 0 cho BHH này.

Đã mở trực tiếp bản thể hiện HĐ `1C26MUV/00005776` ngày **27/01/2026** trên trang tra cứu chính thức MISA. HĐ có mã đơn `580221132778210811`, mặt hàng **KA150P**, số lượng **2 Cái**, tiền hàng **2.461.111đ**, VAT **196.889đ**, thanh toán **2.658.000đ**. Giữ HĐ cho `BHHT2025040425`. Chuỗi khác cùng mã đơn `1C26MUV/00038350 → 1C26MUV/00058635` đã về 0 theo nguồn. Mã tra cứu không lưu trong dữ liệu bàn giao.

Hiện có **30.260 HĐ bán mới / 35.037 dòng**, **1.989 HĐ thay thế về 0** và **7 chứng từ AMIS chờ**. Báo cáo dự kiến **80.398.644.281đ** tiền hàng và **6.431.951.289đ** VAT, chênh AMIS **−22.176.930đ** và **−1.711.831đ**. Đã kiểm 5 Excel, ZIP và đọc lại 62 ô ghi trên Google Sheets. Bằng chứng trong `verified_zero_and_keep_20261109` của `final-summary.json` và `verified-zero-and-keep-google-verification.json`.

---

# Chốt giữ 7 HĐ theo đối chiếu kế toán — 09/11/2026

Đã rà 9 chứng từ có nhiều HĐ ứng viên. Giữ 7 HĐ vì khớp mã đơn và ba khoản tiền, đã đọc lại chi tiết hàng/người mua và dòng KT ghi “xuất đúng”. BHH `BHHT2025038170` ghi hai máy và hai lõi theo combo, tên lõi KA350S trên HĐ khác cách ghi FT KA350 trên AMIS. Tổng đưa vào nhóm giữ **10.364.004đ tiền hàng, 829.121đ VAT, 11.193.125đ thanh toán**.

| BHH | HĐ giữ |
| --- | --- |
| BHHT2025038170 | 1C25MUV/00029306 |
| BHHT2025039308 | 1C25MUV/00025916 |
| BHHL2025000226 | 1C25MUV/00021270 |
| BHHT2025040437 | 1C25MUV/00021799 |
| BHHT2025029775 | 1C25MUV/00017754 |
| BHHT2025040427 | 1C25MUV/00021798 |
| BHHT2025040428 | 1C25MUV/00021797 |

Còn hai BHH trong nhóm này. `BHHT2025035946` có chuỗi `00031054 → 00034893 → 1C26MUV/00002639` đã về 0 theo tab KT, nên chưa có HĐ còn tiền để giữ. `BHHT2025040425` có `1C26MUV/00005776` khớp mã đơn và thanh toán 2.658.000đ, nhưng thiếu bảng chi tiết để kiểm 2 máy KA150P. Chuỗi cạnh tranh `00038350 → 00058635` đã về 0.

Đã ghi riêng 6 HĐ KT đánh dấu xuất trùng và 3 chuỗi KT ghi đã điều chỉnh về 0 nhưng HĐ điều chỉnh chỉ có tiền hàng/VAT 0, ô thanh toán tổng trống. Chưa xác nhận ba chuỗi này về 0. Không tạo dự thảo mới cho các HĐ này trong lượt chốt giữ.

Tổng còn **9 chứng từ AMIS chờ**. Báo cáo dự kiến **80.394.110.023đ** tiền hàng và **6.431.588.548đ** VAT, chênh AMIS lần lượt **−26.711.188đ** và **−2.074.572đ**. Bốn Excel import giữ nguyên SHA-256. Đã kiểm lại Excel báo cáo, ZIP và 77 ô Google Sheets. Quyết định lưu tại `ktoan_confirmed_keep_20261109` trong `final-summary.json` và bằng chứng trong `ktoan-confirmed-keep-google-verification.json`.

---

# Sửa hai liên kết nhầm và giữ bốn HĐ đúng — 09/11/2026

Đã gỡ map nhầm theo xác nhận người dùng. Giữ `1C25MUV/00012360` cho `BHHT2025028267` (mã đơn `580299168658785596`) và `1C25MUV/00012363` cho `BHHT2025028252` (mã đơn `580300943056209030`). Hai chuỗi này khớp tiền hàng/VAT/thanh toán **1.616.759đ / 129.341đ / 1.746.100đ** mỗi chuỗi.

| BHH | Mã đơn | HĐ riêng được giữ | Tiền hàng / VAT / Thanh toán |
| --- | --- | --- | --- |
| BHHT2025027442 | 580356480920290819 | 1C25MUV/00012483 | 628.698đ / 50.296đ / 678.994đ |
| BHHT2025027439 | 580357260801967142 | 1C25MUV/00012486 | 2.295.365đ / 183.629đ / 2.478.994đ |

Đối soát NB dòng 15480/15483 xác nhận mã đơn và tên người mua. Bảng kê và chi tiết meInvoice xác nhận mỗi HĐ đúng khớp tên sau chuẩn hóa, ngày 15/09/2025, mặt hàng, số lượng và cả ba khoản. Mỗi đơn chỉ có một ứng viên khớp cả tên trong nhóm cùng hàng và tiền đã rà. Mã đơn gốc trên bảng kê meInvoice để trống. Các bản ghi KT cùng số nhưng khác ký hiệu/tháng được phân biệt bằng ký hiệu `1C25MUV` và ngày. MeInvoice web yêu cầu đăng nhập; kiểm tra dựa trên bảng kê/chi tiết có sẵn và Google Sheets kế toán.

Đã sửa đúng hai map trên AMIS, AMIS 1, Đối soát NB, điền mã đơn đối chiếu trên Ktoan và cập nhật báo cáo. Rút `00012483` và `00012486` khỏi lô thay thế, chuyển STT `1990 → 1216` (`1C25MUV/00019659`) và `1991 → 1219` (`1C25MUV/00019221`) để giữ dãy liên tục. Hiện còn **1.989 HĐ thay thế về 0, 30.259 HĐ bán mới và 16 chứng từ chờ**. HĐ bán mới giữ nguyên file đã kiểm; hai BHH trên đã có HĐ riêng.

Báo cáo dự kiến **80.383.746.019đ** tiền hàng và **6.430.759.427đ** VAT, chênh AMIS **−37.075.192đ** và **−2.903.693đ**. Đã kiểm lại 5 Excel, ZIP và 162 ô ghi trên Google Sheets. Chi tiết trong `final-summary.json` mục `shared_invoice_map_fix_20261109` và `shared-invoice-map-google-verification.json`.

---

# Thay thế tiếp HĐ cuối chuỗi và lập lại theo AMIS — 09/11/2026

Đã bổ sung dự thảo cho **11/13 chứng từ** theo yêu cầu. STT thay thế **1.981–1.991** tham chiếu đúng HĐ thay thế còn hiệu lực trong từng chuỗi, có ngày và mã CQT nguồn. Tổng số dư được dự kiến đưa về 0: tiền hàng **17.605.890đ**, VAT **1.408.470đ**, thanh toán **19.014.360đ**. STT bán mới **30.249–30.259**, 12 dòng. Bảng đầy đủ nằm trong `continued_reissue_20261109` của `final-summary.json`.

| Chứng từ | PSC2 trừ | Tiền hàng còn lại | VAT tính lại 8% | Thanh toán |
| --- | ---: | ---: | ---: | ---: |
| BHHT2025034835 | 66.000đ | 165.704đ | 13.256đ | 178.960đ |
| BHHT2025032283 | 1.160.000đ | 1.283.519đ | 102.682đ | 1.386.201đ |
| BHHS2025018941 | 237.000đ | 2.418.998đ | 193.520đ | 2.612.518đ |

Hai chuỗi tách chờ xác minh: `BHHT2025027442` gắn HĐ cuối `1C25MUV/00012360`, còn `BHHT2025028267` đang giữ qua F0 `1C25MUV/00010850`; `BHHT2025027439` gắn HĐ cuối `1C25MUV/00012363`, còn `BHHT2025028252` đang giữ qua F0 `1C25MUV/00010855`. Hai chứng từ đang giữ có mã đơn khác với hai chứng từ chờ. Cần chốt liên kết trước khi lập phương án làm mất hiệu lực HĐ cuối.

Hiện có **30.259 HĐ bán mới / 35.036 dòng**, **1.991 HĐ thay thế về 0**, **25 HĐ điều chỉnh về 0**, **18 chứng từ AMIS chờ**. Bảng dự kiến Google còn chênh AMIS **−43.232.773đ tiền hàng** và **−3.396.300đ VAT**. Đã kiểm bộ 5 Excel, ZIP và đọc lại 680 ô đã ghi trên Google Sheets. Bộ vẫn ở trạng thái dự thảo. Phần đối soát năm trên Google Sheets là bản hiện hành; các snapshot trong `report-2025/` lưu mốc dựng trước. Bằng chứng trong `verification.json`, `zip-verification.json` và `continued-reissue-google-verification.json`.

---

# Chốt giữ 10 HĐ khớp và cập nhật bộ import — 09/11/2026

Đã xác nhận 10 HĐ hiện tại khớp tiền hàng, VAT và thanh toán với chứng từ AMIS. Cột Q nguồn AMIS ghi “Giữ HĐ hiện tại — khớp 3 khoản”. Bộ import còn **29 chứng từ chờ**, **1.980 HĐ thay thế về 0** và **30.248 HĐ bán mới**. Rút đúng 7 HĐ trong nhóm vừa chốt khỏi lô thay thế và cập nhật lại STT của 7 dòng chuyển vị trí.

| Chứng từ AMIS | HĐ giữ lại |
| --- | --- |
| BHHT2025024984 | 1C25MUV/00006219 |
| BHHT2025024919 | 1C25MUV/00006149 |
| BHHS2025013243 | 1C25MUV/00005398 |
| BHHT2025021341 | 1C25TUV/00012709 |
| BHHT2025021319 | 1C25TUV/00012729 |
| BHHT2025021665 | 1C25TUV/00012203 |
| BHHT2025021657 | 1C25TUV/00012256 |
| BHDV20250060 | 1C25MUV/00034385 |
| BHHN2025000090 | 1C25MUV/00028840 |
| BHHT2025031775 | 1C25MUV/00023087 |

Đã gỡ liên kết cũ làm 12 HĐ riêng của các chứng từ đã chốt giữ bị treo. Đưa lại 22.660.951đ tiền hàng và 1.812.875đ VAT của 12 HĐ này vào tổng dự kiến. Cộng cả 10 HĐ vừa chốt, Sheets dự kiến 80.361.754.029đ tiền hàng và 6.429.000.067đ VAT, thấp hơn AMIS lần lượt 59.067.182đ và 4.663.053đ.

Các STT thay thế sau khi dời dòng nằm trong mục `kt_keep_10_20261109` của `final-summary.json`. Cột Q nguồn và tab rà soát đã cập nhật lại STT cho 6 HĐ lệch 1–2đ ở đợt trước. Bộ import vẫn là dự thảo.

---


# Đưa HĐ lệch 1–2đ về 0 và xuất lại — 09/11/2026

Áp dụng cả chênh âm và chênh dương khi tiền hàng, VAT và thanh toán đều chênh tối đa 2đ giữa AMIS và meInvoice. Trong nhóm khách lẻ tháng 7–12, có 135 trường hợp đã có dự thảo theo quy trình này và 6 trường hợp được bổ sung. Các HĐ công ty vẫn theo tab HĐ công ty.

| Chứng từ AMIS | HĐ tham chiếu để thay thế về 0 | STT thay thế | STT bán mới |
| --- | --- | ---: | ---: |
| BHHT2025035240 | 1C25MUV/00034896 | 1982 | 30243 |
| BHHL2025000205 | 1C25MUV/00018799 | 1983 | 30244 |
| BHHT2025030189 | 1C25MUV/00022891 | 1984 | 30245 |
| BHHS2025018044 | 1C25MUV/00018130 | 1985 | 30246 |
| BHHS2025015861 | 1C25MUV/00011884 | 1986 | 30247 |
| BHHT2025024845 | 1C25MUV/00006004 | 1987 | 30248 |

Bốn trường hợp đầu tham chiếu HĐ thay thế còn hiệu lực. HĐ 00006004 được chuyển từ giữ sang thay thế về 0. Với dòng KA600FILTERFULL của BHHS2025015861, đơn giá trong file bán mới là 536.506,5đ, bằng 3.219.039đ chia cho 6 cái. Tiền hàng, VAT và thanh toán giữ đúng AMIS. Đã kiểm tra 6 chứng từ không phát sinh tham chiếu BTL hoặc BGG trong bộ hiện tại.

Sheets và bộ import có **30.248 HĐ bán mới trên 35.024 dòng**, **1.987 HĐ thay thế** và **39 chứng từ chờ**. Doanh thu dự kiến trên Sheets là 80.328.071.601đ, thấp hơn AMIS 92.749.610đ. VAT dự kiến là 6.426.305.475đ, thấp hơn AMIS 7.357.645đ. Cột Q nguồn AMIS ghi “Lệch 1–2đ · về 0 và xuất lại”. Các dòng vẫn ở trạng thái dự thảo. Chi tiết nằm trong mục `rounding_reissue_20261109` của `final-summary.json`.

---


# Đồng bộ 8 HĐ khớp sau đợt tách số cuối — 09/11/2026

Đã rút 8 HĐ khỏi lô thay thế và 4 chứng từ khỏi lô bán mới. Bộ import còn **30.242 HĐ bán mới**, **35.014 dòng**, **1.981 HĐ thay thế** và **44 chứng từ chờ**. Danh sách trong `final-summary.json`, mục `kt_final_8_20261109`. Dòng `BHHT2025027502` chờ KT chốt chênh 723.000đ trước khi dùng.

---

# Chốt giữ HĐ `1C25MUV/00006004` — 09/11/2026

Đã chấp nhận chênh 1đ VAT và thanh toán cho `BHHT2025024845`, giữ HĐ và rút khỏi lô thay thế. Lô thay thế còn **1.989 HĐ** và báo cáo còn **51 chứng từ chờ**. Xem `keep_accepted_1dong_20261109` trong `final-summary.json`.

---

# Đồng bộ 18 chứng từ KT sửa số HĐ — 09/11/2026

Đã rút 17 HĐ khớp khỏi lô thay thế về 0 và bỏ `BHHT2025024041` khỏi lô bán mới. Bộ dự thảo có 30.246 HĐ bán mới, 35.018 dòng và 1.990 HĐ thay thế. `1C25MUV/00006004` của `BHHT2025024845` lệch 1đ VAT và 1đ thanh toán, chờ KT chốt trước khi dùng. Xem `kt_remap_18_20261109` trong `final-summary.json`.

---

# Cập nhật theo số HĐ KT đã sửa — 09/11/2026

Lô thay thế hiện có **2.007 HĐ**. Đã rút 25 HĐ chuyển sang giữ và tạm rút 6 HĐ chờ xác minh. Danh sách trong `final-summary.json`, mục `kt_remap_20261109`. Còn **66 chứng từ** trong tab `Cho xac minh` của báo cáo Excel.

Bốn chứng từ `BHHT2025027505`, `BHHT2025027435`, `BHHT2025027502` và `BHHT2025027485` còn dòng trong dự thảo bán mới. Chờ KT chốt liên kết HĐ trùng trước khi dùng các dòng này. Đã ghi chú trên tab `Import HD moi` của Google Sheets.

---

# Kế toán chốt giữ 16 hóa đơn biếu tặng — 09/11/2026

Cả 16 HĐ trong danh sách biếu tặng được giữ theo xác nhận của người dùng về quyết định kế toán. Xem trạng thái tại tab `Bieu tang - KT check` và `MI2025 chua map` trong báo cáo Excel. Số tiền nguồn theo dõi riêng ở cột O của báo cáo chuỗi, nằm ngoài số doanh thu bán hàng dự kiến. Google Sheets đã đồng bộ trạng thái J4:J19, giữ nguyên cột K.

---

# Bổ sung hai chứng từ đổi sản phẩm PSC2 — 09/11/2026

- STT 30.245, `BHHT2025004408`. Bỏ PSC2 540.000đ. Dòng sản phẩm có tiền hàng 1.662.778đ, VAT 176.222đ, thanh toán 1.839.000đ.
- STT 30.246, `BHHT2025011220`. Bỏ PSC2 239.000đ. Dòng sản phẩm có tiền hàng 203.593đ, VAT 35.407đ, thanh toán 239.000đ.
- Thuế suất lưu 8% và giữ tiền VAT theo dòng AMIS người dùng xác nhận. Tiền VAT khác kết quả nhân lại với tiền hàng sau khi trừ PSC2. Kiểm cách meInvoice tiếp nhận tiền VAT trước khi sử dụng hai dòng này.
- Bản cục bộ và Google Sheets có **30.246 HĐ / 35.018 dòng** bán mới, còn 93 chứng từ chờ đều thuộc H2. Báo cáo chuỗi tách riêng 779.000đ PSC2 trong phần giải thích chênh lệch. Xem `final-summary.json`, `verification.json` và `report-2025/verification.json`. Chưa nhập, ký hoặc phát hành.

---

# Điều chỉnh theo quy cách 2024 và báo cáo 2025 trước cập nhật PSC2 — 09/11/2026

Đã sửa quy cách điều chỉnh: **25 HĐ / 33 dòng / 18 chuỗi**. Đảo riêng 7 HĐ điều chỉnh cũ còn tiền, rồi giảm toàn bộ chi tiết 18 F0; 13 điều chỉnh cũ tiền bằng 0 đã kiểm và không tạo bút toán đảo. Tất cả tham chiếu F0; tổng chỉ ghi dòng đầu mỗi HĐ. Giữ nguyên tên hàng, đơn vị, đơn giá nguồn, số lượng/tiền đảo dấu; 8 dòng khuyến mại có cờ 1. Có 13 dòng F0 đơn giá nguồn bằng 0 nhưng thành tiền khác 0: giữ đúng nguồn/quy cách 2024, không tự tính lại giá; xem source_price_flags trong adjustment-rule-2024-plan.json. Kiểm độc lập 86 dòng nguồn và 33 dòng xuất, toàn bộ 18 chuỗi về 0 sau dự thảo ở cả tiền hàng/VAT/thanh toán.

**Tên khách hàng C = H trên F0; Người mua hàng F = L trên F0, lấy riêng đúng nguyên bản, kể cả để trống.** Không dùng tên này thay thế tên kia hoặc áp quy tắc khách lẻ vào bộ điều chỉnh. Bán mới, thay thế và BTL giữ nguyên byte so với ngay trước lượt sửa này. Thay thế vẫn 2.039 HĐ; bán mới 30.244 HĐ/35.016 dòng; BTL 321 phiếu/410 dòng, chưa có tham chiếu HĐ thực tế.

Theo yêu cầu mới, đã thêm tab [Ra soat chuoi 2025](https://docs.google.com/spreadsheets/d/1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA/edit?gid=700000010#gid=700000010), theo bố cục tab báo cáo 2024 người dùng cung cấp. Có **46.194 chuỗi**, 46.208 dòng/18 cột. Số dự kiến chưa VAT: **80.193.955.311đ**, AMIS cùng phạm vi **80.408.395.929đ**, chênh **−214.440.618đ** = chênh nhóm giữ **−81.923.747đ** − 95 CT AMIS chờ **133.737.241đ** + 2 phiếu trả chờ giảm **1.220.370đ**. Các khoản công ty/chờ/16 HĐ biếu tặng tách cột O, không ép chênh bằng 0. BTL chỉ 319/321 phiếu đủ số liệu được cộng dự kiến; tất cả vẫn chờ HĐ tham chiếu. Đây là dự thảo, chưa nhập/ký/phát hành.

Đã đọc lại **1.705.021 ô trên 6 tab**: 873.277 ô import/KT và 831.744 ô báo cáo; công thức đã đối chiếu kết quả tính độc lập. Metadata 5 tab cũ giữ nguyên khi thêm báo cáo; ý kiến KT cột K được bảo vệ. Nguồn meInvoice đến 01/11/2026; chưa kiểm CQT live mới hơn. Bằng chứng: adjustment-rule-2024-verification.json, annual-report-google-verification.json, google-upload-verification.json, report-2025/verification.json, verification.json, zip-verification.json. Ảnh hiển thị ở qa/google-report-2025-final.png và qa/google-adjustment-2024-money.png.

Thứ tự tái tạo: prepare/finalize/enrich_returns/consolidate → audit/review/refine/append_unmapped → audit_adjustment_rule_2024 → apply_adjustment_2024_rule → builder chỉ số 1 và 3 → verify_adjustment_2024_rule + verify_import → prepare_revenue_report + verify_revenue_report → cập nhật Google có đọc live trước → đọc lại → package. Các verifier giữ nguyên dòng/STT của mốc bổ sung trước, unmapped-import-sequences.json và mapping điều chỉnh 18 dòng là **lịch sử trước tái dựng**, không chạy trên bản 25/33. Mapping hiện hành trong authoring.adjustment_rule_2024 và tab Excel Quy cach DC2024. Không chạy lại batch Google lịch sử; giữ mọi ghi chú live.

Người dùng đã cho phép commit/push tất cả dữ liệu công việc, rồi đóng ứng dụng và tắt máy **sau khi xác nhận push thành công**. Runtime, cache phụ thuộc và thông tin xác thực không thuộc dữ liệu bàn giao.

---

# Bộ dự thảo import hóa đơn AMIS 2025 — 09/11/2026

Bốn bộ import đã cập nhật tại [Xử lý HĐ 2025](https://docs.google.com/spreadsheets/d/1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA/edit), không đưa file báo cáo theo yêu cầu người dùng. Gồm `Import HD moi`, `Dieu chinh ve 0`, `Thay the ve 0`, `Dieu chinh BTL - cho so HD`; thêm tab `16 HD bieu tang - KT check` theo yêu cầu tiếp. Đã đọc lại đủ **872.842 ô/năm tab**; kiểm hiển thị phần bổ sung ở100%. Kết quả hiện hành trong `google-upload-verification.json`, `unmapped-google-verification.json`; các JSON mốc dịch vụ/làm tròn/tên người mua, khuyến mại, 419 đơn và 12 đơn H1 là lịch sử trước đó. Phần BTL vẫn là dự thảo chờ tham chiếu HĐ thực tế; không nhập ngay. Không chạy lại các batch ghi cũ lên Google Sheets nếu chưa đọc lại thay đổi live.

## Kết quả

- Nguồn chính: hai file AMIS xuất 08/11/2026 tại `data/0811/`, 65.493 chứng từ bán năm 2025.
- Giữ 34.758 chứng từ tháng 7–12 khớp đủ tiền hàng, VAT và thanh toán theo chuỗi hiện có. Không tạo import bán mới cho các chứng từ này.
- Loại 396 chứng từ theo danh sách `HD CTY 2025` người dùng xác nhận không xuất meInvoice. Loại theo chứng từ/chuỗi, không loại mọi khách hàng có tên công ty.
- 30.244 hóa đơn mới, 35.016 dòng: tháng 1–6 có 29.736 hóa đơn, tháng 7–12 có 508 hóa đơn. Đã gộp vào một file `IMPORT_HD_MOI_2025.xlsx`, STT liên tục 1–30.244. STT đã giao 1–30.232 giữ nguyên; hai chứng từ dịch vụ thêm STT 30.233–30.234, mười chứng từ chênh 2đ thêm STT 30.235–30.244 (12 dòng). Các liên kết/STT trong báo cáo, BTL và CSV theo file gộp.
- **2.039 HĐ thay thế về0** (567 đã giao +1.472 bổ sung chưa nối AMIS), **18 HĐ điều chỉnh về0** (7 đã giao +11 bổ sung). Giữ nguyên STT cũ; bổ sung thay thế STT568–2.039, điều chỉnh STT8–18. Chỉ thay thế chuỗi F0 đơn lẻ; chuỗi đã có điều chỉnh thì điều chỉnh số dư còn lại.
- 95 chứng từ chưa đủ điều kiện: hai chứng từ T1–T6 và 93 chứng từ T7–T12. H1 còn `BHHT2025011220` (thuế suất/phân bổ VAT giữa các dòng chưa rõ) và `BHHT2025004408` (dòng PSC2 SL 0 nhưng thành tiền 540.000đ). Có 77 chuỗi xử lý cũ chờ xác minh, giao với nhóm chứng từ chờ; không cộng hai số này. H1 không có nhóm giữ nguyên theo khớp tiền. Quy tắc xuất mới toàn bộ H1 trừ danh sách công ty vẫn áp dụng; hai chứng từ chờ là phần dữ liệu chưa hoàn tất.
- Tổng bán mới dự thảo: tiền hàng 37.033.879.723đ, VAT 2.962.708.588đ, thanh toán 39.996.588.311đ. Đây là **bán gốc trước trả lại**, không phải doanh thu ròng sau hoàn/trả.

## Bổ sung HĐ meInvoice xuất năm2025 chưa nối AMIS

Theo yêu cầu tiếp, rà đủ **46.983 thành phần chuỗi có HĐ xuất năm2025**, bao gồm mọi HĐ cha/con và liên quan năm2026 trong bộ nguồn. Đối chiếu liên kết AMIS2025, mã đơn/mã CT nguồn hiện tại và lịch sử, AMIS2024; liên kết chưa chắc hoặc mã cùng15số được giữ ngoài nhóm tự động về0. Có43.170 chuỗi đã nối/có ứng viên, 962 chuỗi gốc thuộc năm khác được bảo vệ, 395 chuỗi công ty loại, 955 chuỗi đã về0 bỏ qua. Không hiểu 46.983 là số F0 cần xử lý, hoặc số HĐ đơn lẻ.

**Bổ sung 1.472 thay thế +11 điều chỉnh**, tổng số dư thanh toán cần đưa về0 **2.560.105.903đ**. Gộp vào hai file hiện có, không tạo file import mới; bán mới và BTL giữ nguyên byte. Đã kiểm đóng đủ quan hệ tổ tiên/hậu duệ, từng khoản tiền trước/sau và toàn bộ1505 dòng nguồn cần dùng. Có12 HĐ điều chỉnh với thanh toán bảng tổng trống: bảng chi tiết có ôAB=0 rõ ràng, khớp tiền hàng/VAT; dùng bằng chứng chi tiết này, không sửa ô trống nguồn. Trong đó11 chuỗi còn tiền làm điều chỉnh, một chuỗi gồm `1C25MUV/00026722 →00027949` và `1C26MUV/00057700` đã về0 nên bỏ qua.

Người dùng chốt **16 HĐ biếu tặng/không thu tiền tách sheet hỏi kế toán**, chưa đưa vào import. Tab Google `16 HD bieu tang - KT check` (gid700000104), có nguyên HĐ/ngày/tên đơn vị/tiền/chuỗi/nguồn; **cộtK Ý kiến kế toán** để trống cho KT ghi. Đây không phải tab import. Không gửi tin nhắn cho KT hoặc tự xử lý16HĐ này. Quyết định lưu ở `unmapped-gift-confirmation.json`.

Hai ngoại lệ nguồn giữ ngoài import: **`1C25MUV/00012398`** trạng thái CQT chưa hợp lệ; **`1C25TUV/00001943`** ngoại lệ kế toán đã chốt không xử lý vì TĐ không hợp lệ. Không gộp18 trường hợp chờ của lượt rà meInvoice này vào95 chứng từ chờ AMIS ở phần trên.

Báo cáo thêm `MI2025 chua map` (kết quả từng chuỗi), `MI2025 chuoi` (F0 và các cấp liên kết, HĐ cha/thay thế, trạng thái, tiền nguồn/tiền dùng tính, căn cứ bổ sung), `Bieu tang - KT check` (16HĐ). CấpF trong báo cáo là độ sâu quan hệ, không phải số lần phát hành. Bằng chứng: `unmapped-import-verification.json`, `unmapped-google-verification.json`, `verification.json`, `zip-verification.json`; snapshot trước sửa ở `archive/before-unmapped-meinvoice-20261109/`. Nguồn meInvoice đến01/11/2026; chưa nhập/ký/phát hành hoặc kiểm CQT live sau mốc nguồn.

Tái tạo phần bổ sung: `audit_unmapped_meinvoice_2025_20261109.py` → `review_unmapped_invoice_details_20261109.py` → `refine_unmapped_meinvoice_2025_20261109.py` → `append_unmapped_invoice_import_20261109.py` → builder chỉ số1/2/3 → `verify_unmapped_invoice_import_20261109.py` và verifier chính → staging Google → kiểm live/ghi vùng bổ sung → đọc lại toàn bộ → đóng ZIP. Chính sách16HĐ được áp dụng tại bước refine. Không dùng kết quả audit cũ sau khi thay nguồn; bước append/main verifier kiểm SHA256 nguồn và không ghi ERP.

## Xác nhận dịch vụ, chênh 1–2đ và tên người mua

`BHDV20250115` và `BHDV20250104`: người dùng xác nhận SL đúng là 2, BHDV là dịch vụ. Mỗi chứng từ import 2 Lần ×138.889đ =277.778đ, VAT 22.222đ, thanh toán 300.000đ. Dữ liệu AMIS gốc SL 1 được giữ nguyên; chỉ áp dụng lớp xác nhận SL import 2 cho đúng hai mã này. Không đổi SL của BHDV/PSC khác.

Người dùng cho phép chênh tuyệt đối tối đa 2đ giữa SL × đơn giá (làm tròn đến đồng) và thành tiền dòng AMIS. Giữ nguyên SL, đơn giá, thành tiền, VAT, thanh toán; không nới tiêu chí khớp AMIS–meInvoice, thuế suất hoặc danh sách công ty. Bổ sung 10 chứng từ/12 dòng, thanh toán 21.974.002đ. Các trường hợp chênh 3đ như `BHHT2025027490`, `BHHS2025015861` vẫn chờ.

Tên đơn vị mua hàng chỉ dành cho pháp nhân/tổ chức. Khách lẻ chuyển sang Người mua hàng; nếu đã có tên người mua trên HĐ cũ thì giữ tên đó. Bốn bộ import và tab dự thảo BTL trong báo cáo đã áp dụng: bán mới 35.010 dòng khách lẻ, điều chỉnh 7, thay thế 567, BTL 410 dòng. Giữ sáu dòng của hai pháp nhân rõ tên/MST trong bán mới; không tự mở rộng danh sách 396 chứng từ loại công ty. Tiền, STT, tham chiếu HĐ và nguồn không đổi do sửa tên. Ba file xử lý cũ/BTL được xuất lại để sửa các trường tên, không còn giữ nguyên byte như các mốc trước.

Căn cứ: `service-quantity-confirmations.json`, `rounding-confirmation.json`, `buyer-field-confirmation.json`; kiểm độc lập tại `service-quantity-verification.json`, `buyer-field-verification.json`, `verification.json`. Snapshot trước gộp xác nhận là `archive/before-service-quantity-fix-20261109/`, trước sửa tên là `archive/before-buyer-field-fix-20261109/`. Chưa nhập/ký/phát hành.

## Tám chứng từ khuyến mại đã xác nhận

Người dùng xác nhận 8 chứng từ tiền bằng 0 là hàng khuyến mại. Nguồn AMIS đã có dấu ✓ ở cột Hàng khuyến mại; bước chuẩn bị trước đây bỏ sót điều kiện này nên chặn nhầm. Đã mở đủ 8 chứng từ trên AMIS live: cờ khuyến mại được tích, thuế suất 8%, SL 1, đơn giá/tiền hàng/VAT/thanh toán 0. Chỉ áp dụng đúng 8 mã trong `promotion-confirmations.json`, không tự coi mọi đơn bằng 0 là khuyến mại hoặc thuế suất 0%.

Đã thêm 8 dòng vào dự thảo mới và cột **W - Hàng khuyến mại**, giá trị **1** trên đúng 8 dòng; các dòng cũ để trống W. Khi nhập cần ghép W vào trường Hàng khuyến mại của meInvoice; R:V vẫn là các cột tra cứu cũ. Hướng dẫn MISA cho phép thêm cột và dùng giá trị 1 cho dòng khuyến mại: https://helpv4.meinvoice.vn/kb/nhap-khau-hoa-don-mot-thue-suat/ . Toàn bộ A:V của 34.994 dòng đã giao và tổng tiền giữ nguyên. Ba file điều chỉnh về 0, thay thế về 0 và BTL giữ nguyên byte; không tạo thêm xử lý HĐ cũ đã về 0.

Bằng chứng: `promotion-confirmations.json`, `promotion-verification.json`, `promotion-google-verification.json`; snapshot `archive/before-promotion-fix-20261109/`. Google chỉ thêm 8 dòng, một cột W và sửa hai ghi chú A2/A5; mở rộng bộ lọc/màu xen kẽ, giữ metadata còn lại. Chưa nhập, ký hoặc phát hành.

## Rà 12 đơn tháng 6 theo yêu cầu tiếp

Đã mở đủ 12 chứng từ trên AMIS live ngày 09/11/2026, gồm các đuôi BHHS2025010077, 0076, 0073, 0072, 0071, 0068, 0067, 0066, 0065, 0064, 0062, 0060. Tất cả ngày 27/06/2025, trạng thái Đã lập hóa đơn Chưa phát hành; ký hiệu và số HĐ trống. Đã tìm chính xác mã đơn trong 591 file bảng kê/XML meInvoice cục bộ, bảng Đối soát NB và Ktoan live: chưa có kết quả xác nhận cho cả 12 đơn. Các HĐ trên Sheet thuộc 11/07/2025 và đang liên kết đơn H2 khác; lưu chúng ở cột liên kết chưa xác minh, không dùng làm căn cứ xử lý HĐ cũ.

**HĐ thực tế của cả 12 đơn vẫn chưa xác định.** Trạng thái AMIS không chứng minh chưa phát hành ở hệ thống khác. Không tự gán HĐ theo số, tiền hoặc khách hàng chung. Đã bổ sung 12 đơn/13 dòng vào dự thảo mới theo quy tắc H1; thanh toán 11.699.000đ. Ba bộ import điều chỉnh về 0, thay thế về 0 và BTL giữ nguyên byte. Bảng `Ra soat 12 don T6` trong báo cáo ghi từng mã đơn, HĐ trên Sheet, trạng thái AMIS và URL chứng từ; báo cáo vẫn không tải lên Google theo yêu cầu trước.

Bằng chứng hiện hành: `h1-12-invoice-link-review.json`, `h1-12-invoice-source-audit.json`, `h1-12-import-verification.json`, `h1-12-google-verification.json`; snapshot trước sửa tại `archive/before-h1-12-link-review-20261109/`. Google chỉ thêm 13 dòng vào tab bán mới và sửa ghi chú A2; ba tab khác, bộ lọc, múi giờ và toàn bộ dòng cũ giữ nguyên. Chưa nhập, ký hoặc phát hành.

## Sửa phân loại chuỗi đã về 0

Điều kiện chặn chuỗi gộp trước đây chạy trước kiểm tra số dư 0, làm giữ nhầm 420 chứng từ tháng 4–5. Đã chuyển kiểm tra đủ ba khoản tiền bằng 0 và không có lỗi trạng thái lên trước điều kiện chặn việc xử lý HĐ cũ. Có 419 chứng từ đủ điều kiện được bổ sung (469 dòng, thanh toán 528.838.028đ); `BHHT2025014083` vẫn chờ vì chưa xác định duy nhất thuế suất. Không nới điều kiện cho ô trống, số dư khác 0 hoặc chuỗi có lỗi trạng thái.

Năm chuỗi liên quan là `1C25TUV/00000781`, `00000732`, `00000720`, `00000702`, `00000541`, đều đã về 0. Không tạo thêm xử lý về 0; hai file 7 điều chỉnh/567 thay thế giữ nguyên byte. Ba phiếu `BTL00356 → BHHT2025014212`, `BTL00378 → BHHT2025014037`, `BTL00407 → BHHT2025014069` chuyển sang chờ số HĐ mới; STT bán mới tương ứng 30.073, 29.973, 29.953. Đây là STT tra cứu, chưa phải số HĐ đã phát hành.

Bằng chứng: `zero-chain-correction-verification.json`; bản trước sửa ở `archive/before-zero-chain-fix-20261109/`. Google chỉ thêm 469 dòng bán và sửa 17 ô đã có (hai ghi chú cùng 15 ô liên kết/trạng thái BTL), mở rộng bộ lọc/màu xen kẽ phần bán mới, giữ bộ lọc BTL và múi giờ người dùng. Đã đọc lại toàn bộ bốn tab sau sửa.

## Trả lại hàng bán

Đã tải `Trả lại hàng bán` từ AMIS live và đọc trường **Số CT bán hàng** của đủ 322 phiếu/412 dòng. Tổng trả 444.205.449đ tiền hàng +35.536.421đ VAT =479.741.870đ thanh toán. Phiếu mới so với bộ cũ là `BTL00511`, nối về `BHHS2025023615`, trả 2.433.696đ. Toàn bộ mã hàng trả có trong chứng từ bán gốc được nối.

Giữ cách làm hiện có của dự án: bán gốc và điều chỉnh trả lại tách riêng, không tự bù vào đơn giá hay số lượng bán. Có 307 chứng từ trả hết theo mã hàng/ĐVT/SL/tiền hàng/VAT; vẫn ghi rõ để kiểm soát cả hai bước bán và trả, không coi tiền bán gốc là doanh thu ròng.

Trong `BAO_CAO_IMPORT_2025.xlsx`:

- `Lien ket BTL`: đủ 322 phiếu/412 dòng, có BHH, mã hàng, dòng AMIS và liên kết tới file/STT hóa đơn bán mới khi có.
- `Tra lai cho so HD`: mẫu dự thảo 321 phiếu/410 dòng sau khi loại một phiếu công ty, đồng thời đã tách thành file riêng `IMPORT_DIEU_CHINH_BTL_2025_CHO_SO_HD.xlsx`. **Chưa nhập phần này:** J:M còn trống để điền/xác minh ký hiệu/số/ngày/mã CQT hóa đơn thực tế cần điều chỉnh. Không dùng STT import như số hóa đơn đã phát hành. Nội dung file riêng khớp toàn bộ các ô dữ liệu trong tab dự thảo của báo cáo.
- 240 phiếu chờ số HĐ bán mới; 79 phiếu chờ đối chiếu việc trả lại trên HĐ đang giữ; hai phiếu có ngoại lệ tiền/SL trả là `BTL00449 → BHHS2025006140` và `BTL00488 → BHHS2025012114`. Giữ số nguồn, chưa coi phần vượt là khoản giảm giá được duyệt.

## Cách dùng

1. Xem `Tong hop`, `Cho xac minh`, `Loai cong ty` và `Xu ly chuoi cu` trong báo cáo. Không nhập nhóm chờ.
2. Kiểm lại lịch sử meInvoice mới hơn 01/11/2026 trước khi nhập các file xử lý về 0, để không xử lý trùng phát sinh mới. Chuỗi đã về 0 không được tạo thêm hóa đơn xử lý.
3. File thay thế dùng cột **Là dòng ghi chú/diễn giải =1**, một dòng đưa hóa đơn về 0. Các ô số lượng/tiền/thuế suất trên dòng ghi chú có thể bị meInvoice để trống sau phát hành; phải kiểm lại bản xuất thực tế.
4. File điều chỉnh về 0 dùng một dòng điều chỉnh số tiền chung, SL/đơn giá trống theo mẫu hiện có; số tiền bằng âm số dư chuỗi, tham chiếu F0. Khi vẫn cần tác động HĐ cũ, chuỗi có lịch sử thay thế phức tạp hoặc dùng chung với chứng từ cần giữ/loại được tách chờ. Chuỗi đã xác minh đủ ba khoản bằng 0 không cần xử lý về 0 thêm.
5. File bán mới duy nhất là `IMPORT_HD_MOI_2025.xlsx`. Ngày dự thảo là 09/11/2026; đổi đồng nhất theo ngày phát hành thực tế nếu dùng sau ngày này. Ngày AMIS 2025 nằm ở cột kiểm tra riêng, không phải ngày phát hành mới.
6. Sau khi có số HĐ bán thực tế, hoàn tất tham chiếu phần trả lại và kiểm chưa xử lý trước đó rồi mới nhập phần trả. Giữ cả tiền hàng, VAT, thanh toán nguồn; không tự tính lại theo số lượng × đơn giá để ép cân.

Chưa nhập, ký hoặc phát hành trên meInvoice. Không sửa hai bảng nguồn HD 2025; đã cập nhật bảng đích Xử lý HĐ 2025 theo yêu cầu. Chưa xác minh trực tiếp CQT/chữ ký XML hoặc phát sinh sau bộ nguồn 01/11/2026.

## Nguồn và phần chờ

- AMIS bản sao: https://docs.google.com/spreadsheets/d/1C5hoPDvZKB1vGFSUG9rzChcORt2L9NnsqETVHxPui9I/edit#gid=1264954379
- Danh sách công ty bản gốc: https://docs.google.com/spreadsheets/d/1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw/edit#gid=622685968
- Ktoan bản sao được đọc lại đầy đủ phạm vi 37.232 hàng gồm header và các hàng còn trống theo từng khối; bản gốc cung cấp liên kết AMIS tháng 1–6.
- Lịch sử meInvoice dùng các bảng kê 2025/2026, dữ liệu sau phát hành 29/10 và bổ sung 01/11. Không mặc định ô tiền nguồn trống bằng 0; chỉ bổ sung số dư từ Ktoan khi khớp các khoản nguồn có giá trị, có ghi căn cứ.
- Có 10 chứng từ chỉ trên Sheet nhưng không có trong AMIS mới: giữ ở tab riêng, không tự khôi phục vào import.
- Số khớp tiền không chứng minh cùng nội dung hàng hóa. CSV toàn bộ có cột kiểm nguyên văn tên/ĐVT/SL; tên rút gọn/combo có thể khác nhau, chưa tự diễn giải thành sai hàng hoặc tự xuất lại đơn đã khớp.

## Kiểm tra và tái tạo

Chạy từ gốc `D:\UVG` với Python/Node bundled:

1. `scripts/prepare_import_2025_20261109.py`
2. `scripts/finalize_import_2025_20261109.py`
3. `scripts/enrich_returns_2025_20261109.py`
4. `scripts/consolidate_import_2025_20261109.py` (giữ ánh xạ STT ổn định tại `invoice-sequences.json`)
5. `outputs/import-2025-20261109/build.mjs`
6. `scripts/verify_import_2025_20261109.py`
7. `scripts/test_import_2025_zero_chain.py`, `scripts/test_import_2025_promotions.py`; phần bổ sung dùng `scripts/verify_unmapped_invoice_import_20261109.py`. Hai verifier dịch vụ/tên người mua kiểm mốc lịch sử trước bổ sung meInvoice, không chạy cố định mốc đó trên bản mở rộng.

Script `scripts/audit_h1_12_invoice_links_20261109.py` tái kiểm nguồn và tạo dữ liệu rà liên kết 12 đơn trước khi chuẩn bị import. `scripts/verify_zero_chain_correction_20261109.py`, `scripts/verify_h1_12_import_20261109.py`, `scripts/verify_promotion_import_2025_20261109.py` và các JSON tương ứng kiểm các mốc lịch sử; không chạy verifier cố định mốc đó trên bản hiện hành. Sau kiểm file, `scripts/prepare_google_import_2025_20261109.py` tạo staging; `scripts/prepare_buyer_service_google_patch_20261109.py` chỉ chuẩn bị các vùng sửa hẹp, không tự ghi Google. Đọc live và kiểm điều kiện trước mọi lần ghi.

Thư mục `deliverables/` và ZIP bàn giao có năm file Excel: bán mới, điều chỉnh về0, thay thế về0, **điều chỉnh BTL chờ số HĐ**, báo cáo. File BTL được bổ sung riêng theo phản hồi người dùng, không chỉ nằm trong báo cáo. Các file `.inspect.ndjson` là file kiểm tra kỹ thuật được chuyển sang `qa/`. Bảy file bán mới chia tháng cũ giữ riêng trong `archive/`, không nhập đồng thời với file gộp.

Các script không ghi Google Sheets hoặc AMIS/meInvoice. Snapshot nguồn và cache là dữ liệu cục bộ phục vụ tái tạo. Không chạy lại builder cũ của năm 2024 để ghi đè bộ này. File `verification.json` ghi kết quả kiểm nguồn, toàn bộ ô import xuất ra, số tiền theo chứng từ, STT, loại công ty, chuỗi về 0 và dữ liệu trả lại.
