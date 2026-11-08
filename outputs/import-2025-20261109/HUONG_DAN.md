# Bản cuối: điều chỉnh theo quy cách 2024 và báo cáo 2025 — 09/11/2026

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
