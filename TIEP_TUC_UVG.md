# Chốt bộ bàn giao để commit/push ngày 08/10/2026

Đã kiểm tra lại bằng `scripts/verify_handoff.py`: 18 nguồn khớp manifest hiện tại; nhóm chưa điều chỉnh8404hóa đơn/13543dòng, nhóm đã điều chỉnh1600hóa đơn dự kiến/40777dòng/791chuỗi, sốhọc sau kếhoạch về0. Chưa import/ký/phát hành.

File chi tiết2024 khác commit cũ về định dạng/cấu trúcExcel, nhưng so sánh từng ô cả hai sheet không có giá trị thay đổi, kích thước cácsheet giữ nguyên. Đã cập nhật bytes/hash trongSOURCE_MANIFEST.json theo file đang bàn giao; bằng chứng `scripts/data/source-handoff-check.json` giữ hash cũ vàmới. Điều này giải quyết ghi chú “nguồn modified chưa xác minh” ở phần lịch sử phía dưới.

Bàn giao gồm nguồn,3workbook chính, bảng kiểmtra đã tạo, script/JSON căn cứ và tài liệu. Ảnh/log/snapshotExcel trunggian không đưa lênGit; các verifycũ phụthuộc snapshot không dùng trực tiếp trênmáy mới. Dùng verify_handoff.py đểkiểm tra chỉ đọc, không chạy lại builder đểghiđè kếtquảKT.

Riêng00000508: vẫn phải kiểm dòng chi tiết, tính chất dòng và phần tổng trongXML. Phương án đảoF1 rồi giảmF0 chỉ làdựthảo sốhọc, chưachốt cách sửa nội dung. Chưa mở19chuỗi chờ hoặc237chuỗi bằng0 vàoimport vì còn chờKT.

Phần “chưa commit/push” bên dưới mô tả lúc ghi nhớ trước bước đồngbộ; xem commit nhận được và xác nhậnpush trongchat để biết bộ này đã lênremote chưa. Máy ởnhà đọc HUONG_DAN_MAY_O_NHA.md và phần tổngquan ngay dưới.

---
# Tổng quan bàn giao TOÀN BỘ dự án UVG

Người dùng yêu cầu ngày 08/10/2026: ghi nhớ tất cả công việc và kinh nghiệm của dự án để qua máy khác vẫn hiểu, không chỉ ca 00000508. Tài liệu này là điểm vào chung; phần cập nhật trạng thái ngay dưới và lịch sử phía sau bổ sung chi tiết. Trạng thái mới nhất ưu tiên hơn phần lịch sử. Không hiểu “đã làm” là đã phát hành: hiện làm phân tích và chuẩn bị file.

## Mục tiêu kinh doanh và nguồn sự thật

Công ty bán máy lọc không khí qua TikTok, Shopee, Lazada và ngoài sàn. AMIS là sổ được đội kế toán làm lại, khách hàng chốt; meInvoice là lịch sử xuất của kế toán cũ. Có hóa đơn gộp và các lần điều chỉnh cũ cần rà. Mục tiêu đang làm: tách ngoại lệ, đưa nhóm ngoài ngoại lệ về 0 bằng kế hoạch điều chỉnh có kiểm tra, rồi chuẩn bị hóa đơn mới theo AMIS để tránh trùng doanh thu. Số 70.000 hóa đơn mới/10.500 điều chỉnh từng được nêu là quy mô ban đầu, không phải số đã chốt phát hành.

Bằng chứng phải tách rõ: file xuất nguồn, kết quả tính toán, XML đã ký, dữ liệu CQT, dự thảo import và hóa đơn đã phát hành. Đồng bộ thành công/có mã CQT không tự chứng minh đầy đủ hoặc đúng nội dung. Không xóa rồi nạp lại meInvoice theo suy đoán. Chưa có xác minh trực tiếp toàn bộ CQT/XML.

## Những giai đoạn đã làm

| Giai đoạn | Đã làm và bài học cần chuyển tiếp |
|---|---|
| Đọc nguồn và đối chiếu AMIS–meInvoice | Đã phân tích dữ liệu theo tháng, model, kênh khi có căn cứ. Bộ đối chiếu ban đầu có AMIS 93.194.013.744, meInvoice 95.594.027.824, chênh −2.400.014.080 đồng; đây là số lịch sử cần giải thích, không phải lệnh xuất giảm/bù và chưa tái khẳng định trên toàn bộ nguồn mới. Không ghép theo mã đơn sàn hoặc chỉ vì giống hàng/giá. |
| Rà bộ ưu tiên | Đã rà 321 hóa đơn ưu tiên trong đợt đầu: 299 thiếu liên kết, 14 thiếu tiền, 8 lệch số học (các nhóm có giao nhau); 39 mô tả có số tham chiếu chỉ là ứng viên, không phải liên kết đã chứng minh. Nhiều output lịch sử đã dọn; xem phần lịch sử/Git, không mặc định đường dẫn cũ còn tồn tại. |
| Chốt phạm vi và cập nhật nguồn | F0 01/01/2024–30/06/2025; theo chuỗi liên quan đến hết31/12/2025. Đợt nguồn mới gồm 18 file theo manifest: 2 AMIS,16 meInvoice; chi tiết phủ48.874 hóa đơn của10 phần bảng tổng. Loại trùng2 bảng điều chỉnh;8 liên kết2024 bổ sung từ Git có provenance. Số48.874 là bộ xuất, không phải sốF0 cần xử lý. |
| Tách ngoại lệ và dựng chuỗi | Nhận diện tên công ty/doanh nghiệp hoặc định danh dạngMST trong nguồn, giữ chuỗi liên quan; chưa xác thựcMST. Ngoại lệ238F0/326hóa đơn. Đã điều chỉnh ngoài ngoại lệ1047F0/2113hóa đơn. Loại6chuỗi thay thế khỏi workflow điều chỉnh; nghiệp vụ thay thế chưa xử lý. F1/F2/F3 là vị trí hóa đơn liên quan, không tự suy ra F2 sửaF1. |
| Lập kế hoạch điều chỉnh cũ | Dùng bảng tổng cho tiền từngF, cộng dồn hiện tại và sau kế hoạch. Thêm số khối mới theo thực tế từng chuỗi, không giới hạnF4–F6. Bỏ từ “sửa” ở tiêu đề. Số học bằng0 không đồng nghĩa nội dung đúng. Kế hoạch đảoF cũ rồi giảmF0 đang là dự thảo, không phải kết luận bắt buộc về pháp lý. |
| Rà237chuỗi bằng0 | Đã kiểm tra dòng bán/khuyến mãi và điều chỉnh tổng/chi tiết; noteAB đểKT rà. 185F0 có tất cả dòng thành tiền khác0 chỉ mô tả chi tiếtgốc; không kết luận điều chỉnhcũ đúng. Các thống kê kiểm tra nằm trong JSON và phần cập nhật dưới. Người dùng đã gửifile choKT, cần nhận phản hồi trước khi mở nhóm chờ. |
| Nhóm chưa điều chỉnh | Đã dựng8410F0, kiểm tra dòng0, xác nhận sáu thanh toán tổngtrống=0 theouser, tạo sheetimport. Tách6hóa đơn sốlượng0 sang sheetriêng và loại toàn bộ27dòng; còn8404hóa đơn/13543dòng. Đã rà lại mộtvòng đối chiếu nguồn, sốtiền, thamchiếu, tên, ngày, khuyếnmãi và loại trừ; final-audit.json ghi passed=true. |
| Nhóm đã điều chỉnh | Đã chuẩn bị sheetimport791chuỗi/1600hóa đơn dự kiến/40777dòng; sheetchờ19chuỗi. Kiểm kếhoạch về0 và bảo toàn cácsheet/cột trước. Không tự gộp lại nhóm237bằng0 vào đợtimport. |
| Hướng dẫn meInvoice và nghiên cứu nghiệp vụ | Đã trao đổi bước chọn mẫu, chọn trạng thái Hóa đơn điều chỉnh, ghép cột/xem trước. Kýhiệu hóađơn mới chọn theo mẫu đang áp dụng; thamchiếu gốc vẫn đúng1C24TUV/1C25TUV. Phải kiểm cấu hình môi trường thực, ngày, tùychọn tựtính trước khi nhập. Đã tìm nguồnMISA và cơquanthuế; không tự kết luận một dòng điều chỉnh tổng sai, không coi Excelchi tiết là XML. 508 là một ca minh họa, không phải toàn bộ dự án. |

## Quy tắc dữ liệu/import đã chốt trong quá trình làm

- Mẫu người dùng cung cấp: `data/Du lieu Meinvoice/Mau HD Dieu chinh.xls`. Sheet import ánh xạ24cột A:X: STT; ngày mới; tênKH; địachỉ; MST; ngườimua; email; phươngthức; cờHĐngoàihệthống; kýhiệugốc; sốgốc; ngàygốc; mãCQT; lýdo; tênhàng; đơnvị; sốlượng; đơngiá; thànhtiền; tổngtiềnhàng; thuếsuất; tổngthuế; tổngthanhtoán; hàngkhuyếnmãi. Kiểm lại tiêuđề trongworkbook trước khi ghép.
- Dòng đầu mỗi hóa đơn chứa tổng tiền hàng/thuế suất/tổng thuế/tổng thanh toán; dòng sau để trống phần tổng. Lấy phần tổng từ bảng kê tổng, không SUM chi tiết để thay số nguồn.
- Giảm gốc theo chi tiết: số lượng/thành tiền lấy dấu giảm, giữ đơn giá nguồn; thành tiền nguồn0 thì cột khuyến mãi1. Khác biệt sốlượng×đơngiá do làm tròn phải giữ tiền nguồn và kiểm tính tựđộng trong meInvoice.
- Đảo khoản điều chỉnhcũ: dùng số tổng đã kiểm của hóađơn cầnđảo; dòng diễn giải nêu khoản cũ; thamchiếu đúnggốc. Cần KT chốt bảnchất vàXML, không xem quytrình dựthảo này là bắtbuộc mọi trườnghợp.
- Không đoán phần thiếu; ngoại lệ được user xác nhận phải ghi rõ. Gốc574 có quyước riêng từ đợt trước, không áp cho các hóa đơnkhác. Gốc1314 ngày05/07/2025 đã đượcuser xác nhận theoảnh; nguồn tổng/liênkết18/08/2025 giữởnguồn.
- File nguồn có đuôi.xls nhưng bộ hiện có chứaOOXML; đọc theo nội dung thực tế, không tự đổi/sửa file nguồn. GiữSHA256/provenance theo từng đợt kiểm, không khẳng định hashcũ xác nhận trạngthái hômnay.

## Việc còn lại và cách tiếp tục

1. Nhận phản hồiKT về237chuỗi bằng0 và19chuỗi đangchờ; cập nhật ghi chú, căn cứ và quyếtđịnh theo từngchuỗi. Không chỉ dựa vào tổng0.
2. Bổ sungXML/bản hiển thị vàlịch sửCQT cho nhóm lệch/thiếu/không rõ. Đọc cả phần dòng vàtổng, phânbiệt lỗi xuấtbảng/mẫu với lỗi dữliệu đãký. Kiểm các lầnxử lý2026 trước khi thựcphát hành để tránh làmtrùng lịch sử ngoài kỳxuất.
3. Kiểm ngày/mẫu/cột/tựtính trên meInvoice thực dùng; xemtrước mộtlô nhỏ và đốichiếu sốtiền/thuế/khuyếnmãi trước khi chốt áp dụng. Chưa thao tácimport/ký/phát hành trongphiên này.
4. Chỉ chuyển nhómchờ vàoimport khi đã có kếtquả; xửlý ngoại lệ/thaythế hoặc xuấtmới AMIS làphần tiếp theo chưa hoànthành, không tự dùng tổngchênh làm sốphát hành.
5. Trước chuyể máy: đồngbộ tài liệu này,3workbook chính,script,nguồn/manifest cầnthiết vàghi chúKT. KiểmGit trạngthái, bảo toàn thayđổi; không resetpull đèfile. Memorycục bộ không thaythế bộ dựán chuyểnmáy.

Prompt để mở trên máy khác:

> Tiếp tục toàn bộ dự án UVG. Đọc TIEP_TUC_UVG.md từ đầu và README.md, ưu tiên trạng thái mới nhất hơn phần lịch sử. Kiểm tra các workbook hiện tại và ghi chú kế toán trước khi sửa. Giữ nguồn/số 0 đầu/ô trống; không tái tạo đè sheet import hoặc ghi chú. AMIS là mục tiêu, meInvoice là lịch sử; tổng bằng 0 không xác nhận pháp lý. Tiếp tục từ việc còn lại đã ghi, không chỉ xử lý 00000508. Phân biệt phân tích, dự thảo import, xác minh XML/CQT và phát hành. Kiểm tra việc đã nhận đủ thay đổi chưa commit/push trên máy trước.

---
# Bàn giao cập nhật 08/10/2026 – sau khi chuẩn bị import

Phần này thay thế các trạng thái cũ phía dưới về số file, sheet import và bước đang làm. Các số dưới đây là kết quả đã kiểm tra trong phiên làm việc, cần kiểm tra lại file nếu đã được kế toán/người dùng sửa sau đó.

## Quyết định và kinh nghiệm cần giữ

- Phạm vi F0: 01/01/2024–30/06/2025; theo hóa đơn liên quan đến 31/12/2025. Chưa xác minh các lần xử lý năm 2026 trên hệ thống sống.
- Mục tiêu: đưa nhóm ngoài ngoại lệ về 0, sau đó chuẩn bị xuất mới theo AMIS; không tự phát hành chỉ dựa vào tổng chênh lệch.
- Dùng các workbook hiện tại; thêm sheet, không tạo nhiều file báo cáo. Giữ số hóa đơn dạng text đủ 8 chữ số, giữ nguồn và ô trống. Giả định tiền trống thành 0 chỉ áp dụng các trường hợp người dùng xác nhận.
- Tổng tiền hàng, thuế, thanh toán của F0 và F cũ lấy từ bảng kê hóa đơn đã sử dụng (bảng tổng), không cộng chi tiết thay thế. Chi tiết dùng để chuẩn bị hàng hóa và phát hiện lệch; chưa đủ để phủ định XML đã ký.
- Cột Hàng khuyến mãi = 1 khi thành tiền nguồn của dòng = 0, theo quy ước import người dùng đã chốt; không suy luận chỉ từ tên sản phẩm. Đây là quy tắc xử lý file, không tự xác nhận bản chất pháp lý khuyến mãi.
- Tên khách hàng = Người mua hàng. Người mua trống thì giữ trống, không tự bịa tên. Ngày import đã dùng 08/10/2026; phải kiểm tra lại ngày khi thực tế nhập.
- F4/F5/F6 chỉ là ví dụ tên khối, số lần dự kiến phải theo từng chuỗi. Đã bỏ từ “sửa” ở tiêu đề các F. Cột AB của bảng đã điều chỉnh là ghi chú kiểm tra.
- Chuỗi cộng dồn bằng 0 chỉ xác nhận số học, không chứng minh hóa đơn điều chỉnh cũ đúng nội dung/quy định. Một dòng điều chỉnh tổng không tự động là sai.
- Khi số lượng 0: tách toàn bộ hóa đơn sang sheet kiểm tra và loại toàn bộ hóa đơn khỏi import; không chỉ bỏ dòng 0.

## File và trạng thái đang tiếp tục

1. `outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx`
   - `Chua dieu chinh`: 8.410 F0 / 13.570 dòng nguồn.
   - `Import dieu chinh`: 8.404 hóa đơn / 13.543 dòng; 1.275 dòng thành tiền 0 được đánh dấu khuyến mãi.
   - `Kiem tra SL 0`: 6 hóa đơn / 27 dòng bị loại toàn bộ khỏi import: 1C24TUV/00000774, 00000783, 00000784, 00000786; 1C25TUV/00002796, 00002797.
   - Người dùng xác nhận thanh toán tổng = 0 cho sáu trường hợp nguồn trống: 1C24TUV/00000671, 00000744; 1C25TUV/00000391, 00000761, 00000835, 00004748. Đã ghi chú, không thay ô nguồn.
   - Audit `scripts/data/untreated/final-audit.json` ghi passed=true. Có 931 dòng / 209 hóa đơn mà số lượng × đơn giá khác thành tiền nguồn (tối đa 52 đồng); giữ thành tiền nguồn và kiểm tra tùy chọn tự tính khi nhập.
2. `outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx`
   - 1.047 F0 / 2.113 hóa đơn cũ. Trạng thái cộng dồn: 804 khác 0, 237 bằng 0, 6 thiếu căn cứ tổng.
   - `Import dieu chinh`: 791 chuỗi / 1.600 hóa đơn dự kiến / 40.777 dòng. Kế hoạch hiện tại đảo các điều chỉnh cũ rồi giảm chi tiết F0; đã kiểm tra cộng dồn dự kiến về 0, chưa phải phê duyệt phát hành.
   - `Cho KT xac nhan`: 19 chuỗi giữ ngoài import; xem lý do từng dòng.
   - 237 chuỗi bằng 0 gồm 237 F0 + 237 F1 = 474 hóa đơn cũ. Đã note AB để kế toán rà. Audit ghi 196 điều chỉnh dạng tổng, 195 lệch tiền chi tiết/bảng tổng, 41 có tên hàng; các nhóm không phải kết luận pháp lý và có thể giao nhau.
   - Người dùng đã gửi file cho kế toán kiểm tra. Chờ kết quả, không tự đưa nhóm đang chờ vào import.
3. `outputs/NGOAI_LE_F0_2024_T6_2025.xlsx`: 238 F0 / 326 hóa đơn; giữ riêng.

## Trường hợp 00000508 – chưa được chốt bằng XML

`1C25TUV/00000508`, ngày 26/04/2025, điều chỉnh `1C24TUV/00000003`, ngày 10/01/2024.

- Bảng tổng: tiền hàng −356.085; thuế −28.487; thanh toán −384.572.
- Chi tiết nguồn: một dòng “Điều chỉnh giảm thành tiền hoá đơn số 00000003”, số lượng/đơn giá/tiền hàng 0; thuế và thanh toán đều −28.487.
- F0: 11.334.130 / 906.730 / 12.240.860. Cộng dồn theo bảng tổng: 10.978.045 / 878.243 / 11.856.288.
- Chưa có XML/PDF hóa đơn trong dữ liệu đã kiểm tra. Không kết luận hóa đơn sai chỉ từ Excel, không tự đảo khoản thuế lần nữa. Chuỗi này vẫn ở sheet chờ kế toán; kế hoạch đảo F1 rồi giảm F0 đang là dự kiến.
- XML cần đối chiếu cả nhóm hàng hóa và nhóm tổng tiền, tính chất dòng ghi chú/diễn giải, số đã ký/gửi CQT. Nếu chỉ xuất báo cáo sai thì sửa lỗi báo cáo; nếu dữ liệu hóa đơn sai thì xác định phần chênh lệch và xử lý tiếp theo quy định tại thời điểm thực hiện.

Các nguồn đã nghiên cứu (phải đọc lại nếu áp dụng về sau):
- MISA đã ghi nhận thiếu tiền hàng/thuế khi xuất chi tiết hóa đơn điều chỉnh giảm ở phân hệ xử lý đầu vào; sửa từ R35 ngày 03/10/2023. Chỉ là lỗi tương tự, chưa chứng minh 508 mắc đúng lỗi: https://helpv4.meinvoice.vn/kb/tinh-nang-moi-inbots/
- MISA hướng dẫn phân biệt lỗi mẫu hiển thị và XML thực sự thiếu tiền; mẫu Web mở sửa rồi lưu, Desktop liên hệ MISA, XML đã phát hành thiếu tiền thì điều chỉnh bổ sung: https://helpv4.meinvoice.vn/kb/web-hoa-don-dieu-chinh-khong-hien-thi-so-lieu-o-dong-tong-tien-hang-tien-thue-gtgt-va-dong-tong-tien-thanh-toan/
- Chỉ điều chỉnh tiền thuế có thể có tiền hàng 0, thuế âm bằng thanh toán âm. Điều chỉnh thành tiền toàn hóa đơn có thể dùng dòng Ghi chú/diễn giải: https://helpv4.meinvoice.vn/kb/cach-ghi-thong-tin-tren-hoa-don-dieu-chinh-noi-dung-ve-gia-tri-tren-hoa-don-2/
- Doanh nghiệp từng điều chỉnh sai tiếp, được cơ quan thuế giải đáp tiếp tục điều chỉnh đúng thực tế năm 2023: https://qlg.mof.gov.vn/hoidapcstc/home/cthoidap/137315 . Không mặc định dùng văn bản năm 2023 để phát hành năm 2026.

## Tiếp tục trên máy khác

- Đọc phần cập nhật này trước; mở đúng workbook hiện tại và xem ghi chú/chờ kế toán. Không tái chạy script cũ để ghi đè các sheet import/ghi chú mới.
- Đường dẫn dự án trên máy mới có thể khác D:\UVG; dùng đường dẫn tương đối trong repo. Runtime Node/Python và node_modules phải cài/tra lại trên máy đó, không mang đường dẫn cache của máy cũ sang.
- Script liên quan: `scripts/data/untreated/add-import.mjs`, `final-audit.py`; `scripts/data/check-review/prepare-adjusted-import.py`, `add-adjusted-import.mjs`, `verify-adjusted-import.py`, `zero-adjustment-review.json`. JSON lớn đọc bằng parser, không in toàn bộ.
- Xuất workbook lớn bằng artifact-tool từng khối; Node đã cần heap 8–12 GB. Python dùng đọc/đối chiếu. Kiểm tra bảo toàn sheet cũ, số tham chiếu, tiền, khuyến mãi, loại trừ và tổng dự kiến sau mỗi sửa.
- Chưa nhập vào meInvoice, ký hoặc phát hành. Sheet import là dự thảo đã kiểm tra dữ liệu, không chứng minh phát hành hợp lệ.
- Kiểm tra `git status` trước khi bàn giao. Phiên này thấy file nguồn `data/Du lieu Meinvoice/Bang_ke_chi_tiet_HD_da_su_dung_2024.xls` đang modified: chưa xác minh nguyên nhân; không reset/ghi đè hoặc khẳng định nguồn hiện khớp manifest khi chưa kiểm hash.
- Tài liệu, workbook và script mới hiện lưu cục bộ. Chưa commit/push trong lần ghi nhớ này; clone/pull máy khác chưa tự có các thay đổi. Cần đồng bộ nguyên bộ thay đổi trước khi chuyển máy.

---
# UVG – trạng thái hiện tại ngày 08/10/2026

Phạm vi người dùng chốt ngày 07/10/2026: chỉ chọn hóa đơn gốc F0 có ngày từ 01/01/2024 đến hết 30/06/2025; theo các hóa đơn liên quan đến những F0 này đến hết 31/12/2025.

Đã làm lại hai report theo nguồn meInvoice người dùng xuất đến 31/12/2025 và format bảng đã điều chỉnh được chọn làm chuẩn. Ba file chi tiết khớp đủ 48.874 hóa đơn trong 10 phần bảng kê tổng, gồm 34.910 hóa đơn `1C25MUV`. Đây là toàn bộ bộ xuất, không phải số lượng F0 cần xử lý. `SOURCE_MANIFEST.json` đã cập nhật 18 file hiện có (2 AMIS, 16 meInvoice); SHA-256 xác nhận dữ liệu nguồn không đổi trong quá trình tính và xuất report.

## Phạm vi F0 và hóa đơn liên quan đã chốt

- F0 là hóa đơn gốc đầu chuỗi, có ngày trong khoảng 01/01/2024–30/06/2025, bao gồm cả hai ngày. Không chọn thêm F0 ngoài khoảng này chỉ vì có trong bộ xuất nguồn.
- Theo liên kết từ những F0 đủ điều kiện để lấy hóa đơn điều chỉnh, thay thế và các lần xử lý tiếp nối có bằng chứng tham chiếu về cùng F0 (F1/F2 và các cấp tiếp theo nếu có). Chỉ xét hóa đơn liên quan có ngày đến hết 31/12/2025; không giới hạn chúng ở ngày 30/06/2025.
- Hóa đơn ngày 01/07/2025–31/12/2025 chỉ thuộc phạm vi khi có bằng chứng là hóa đơn liên quan đến F0 đủ điều kiện. Các lần điều chỉnh/thay thế có ngày năm 2026 nằm ngoài kỳ đang xét; giữ nguyên nguồn và không cộng chúng vào kết quả chốt đến 31/12/2025.
- Xác định phạm vi bằng ngày chứng từ và liên kết chuỗi, không chỉ bằng nhãn trạng thái hiện tại. Người dùng đã cập nhật thông tin `1C25MUV/00034730` trên danh sách HĐ từ máy tính tiền: trạng thái chuyển sang đã bị điều chỉnh; người dùng xác nhận lần điều chỉnh có ngày 03/07/2026. Hóa đơn 34730 có ngày 31/12/2025 nên không được chọn làm F0 trong phạm vi này; lần điều chỉnh ngày 03/07/2026 cũng nằm ngoài mốc chốt.
- Quy tắc ngoại lệ công ty/cá nhân có MST vẫn giữ nguyên. Việc thu thập liên kết thay thế để hiểu chuỗi không phải quyết định xử lý nghiệp vụ thay thế; bước xử lý hiện tại vẫn ưu tiên nhóm đã điều chỉnh ngoài ngoại lệ theo bàn giao.
- Người dùng đã cho phép tái tạo hai Excel và dọn output cũ. Chưa sửa dữ liệu nguồn, tạo import, ký hoặc phát hành hóa đơn.

## Bảng đang dùng

`outputs/` chỉ giữ hai Excel dưới đây, trực tiếp tại thư mục gốc, không có thư mục con. Output cũ đã được người dùng dọn sau khi lệnh xóa tự động bị chặn.

1. `outputs/NGOAI_LE_F0_2024_T6_2025.xlsx`: 238 F0 (118 năm 2024, 120 tháng 1–6/2025), 326 hóa đơn riêng biệt. Nhận diện ngoại lệ từ tên công ty/doanh nghiệp hoặc định danh dạng MST trong nguồn; chưa xác thực MST. Giữ đủ chuỗi liên quan. Bảng chính theo format chuẩn, thêm đúng hai cột đầu **Tên người mua** và **MST/CCCD chủ hộ nguyên bản**.
2. `outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx`: 1.047 F0 ngoài ngoại lệ (591 năm 2024, 456 tháng 1–6/2025), 2.113 hóa đơn riêng biệt. Chỉ chọn khi có liên kết điều chỉnh trong kỳ; loại 6 chuỗi thay thế. Bố cục 2 hàng tiêu đề phân màu F0/F1/F2/F3, mỗi khối 6 cột như mẫu. Hai F0 `1C25TUV/00000455` và `1C25TUV/00000766` có 3 hóa đơn liên quan nên cần khối F3; vị trí khối không đồng nghĩa F2 sửa F1.

Cả hai có tab `Huong dan`, `Chi tiet nguon`, bộ lọc, cố định tiêu đề, số hóa đơn đủ số 0 đầu và số tiền cộng dồn tính sẵn. Tab nguồn ghi loại/tham chiếu, trạng thái hiện tại, tên/định danh và vị trí nguồn; ô tiền thiếu vẫn trống.

Nguồn mới có 2.067 liên kết riêng biệt; hai file điều chỉnh trùng tập liên kết được loại trùng. Thêm 8 liên kết năm 2024 bị thiếu trong bộ xuất mới từ các file lịch sử Git tại commit cố định `9c50b0424a242695fa31182ab826613bf0fa699f`, kiểm SHA-256 theo manifest tại commit đó. Không khôi phục hoặc sửa các Excel nguồn đã xóa. Provenance ghi tại tab nguồn.

F0 năm 2024 tăng 590 → 591 do `1C24TUV/00000366` có liên kết điều chỉnh đến `1C25TUV/00002060` ngày 29/05/2025, dù bảng tổng vẫn ghi hóa đơn gốc là mới. Không còn F0 ngoài ngoại lệ có nhãn đã xử lý nhưng thiếu liên kết trong phạm vi kỳ xuất; điều này không xác nhận lịch sử năm 2026 hay nội dung nghiệp vụ.

## Quyết định người dùng và phần còn lại

- Mục tiêu người dùng: đưa nhóm ngoài ngoại lệ về 0, sau đó xuất mới từ AMIS; loại doanh thu đã nằm trong ngoại lệ giữ nguyên để tránh trùng. Chưa chốt nghiệp vụ phát hành.
- Chỉ xử lý nhóm đã điều chỉnh trong bước hiện tại; chưa xử lý thay thế hoặc nhóm chưa sửa.
- Bảng điều chỉnh có 16 chuỗi có ghi chú: 8 lệch số học (gốc 2024: 93, 228, 283, 285, 535; gốc 2025: 56, 102, 159), 6 chuỗi có ô tiền thiếu, quy ước riêng gốc 574 và 1 chuỗi có ngày khác nhau giữa nguồn. Gốc `1C25TUV/00001314` có hóa đơn liên quan `1C25TUV/00012885`: chi tiết ghi 05/07/2025, bảng tổng/liên kết ghi 18/08/2025. Ngày ở bảng chính đã dùng **05/07/2025** theo xác nhận của người dùng và ảnh danh sách meInvoice; tab nguồn giữ ngày bảng tổng 18/08/2025 và cả hai vị trí nguồn.
- Ngày 07/10/2026, người dùng xác nhận đã cập nhật thông tin hóa đơn `1C25TUV/00012885` trên meInvoice nhưng các mục vẫn lệch ngày; sau đó xác nhận ngày 05/07/2025 và gửi ảnh danh sách **Hóa đơn bán hàng có mã CQT** thể hiện ký hiệu 1C25TUV, số 00012885, ngày 05/07/2025. Ảnh là bằng chứng từ danh sách; chưa có XML để đối chiếu độc lập. Giữ cờ ngày giữa nguồn khác nhau, không coi lỗi đồng bộ đã giải quyết. Hai ngày đều trong kỳ hóa đơn liên quan đến 31/12/2025; phạm vi F0 và cộng dồn chuỗi 1314 (0 đồng) không đổi. Quy ước ngày bảng chính được lưu trong `USER_CONFIRMED_DATES` của script để tái chạy không mất xác nhận.
- Bảng ngoại lệ có 2 chuỗi có ghi chú: gốc `1C24TUV/00000060` có chuỗi thay thế, không cộng bản bị thay thế; gốc `1C24TUV/00000063` có ô tiền thiếu. Ghi chú chuỗi thay thế là mô tả cách cộng dồn, chưa phải chỉ định xử lý.
- Rà thêm ngày 08/10/2026: bảy hóa đơn `1C24TUV/00000063`, `1C25TUV/00000006`, `00000016`, `00006108`, `00004203`, `00004204`, `00009186` đều có số thanh toán **0 đồng ghi rõ ở cột U bảng kê chi tiết**; mọi dòng chi tiết của mỗi hóa đơn đều có ô thanh toán số 0. Bảng tổng ghi tiền hàng = 0, thuế = 0, nhưng thanh toán trống; bảng liên kết cũng trống thanh toán ở sáu hóa đơn 2025. Vì vậy đây là ô trống ở nguồn tổng, không phải thiếu giá trị thanh toán trên mọi nguồn. Chưa sửa hai Excel trong lần hỏi đáp này; report vẫn giữ ô nguồn trống và chưa dùng số thanh toán từ chi tiết để tính cộng dồn. Số 0 của từng hóa đơn không chứng minh cả chuỗi có số dư 0.
- Riêng 1C24TUV/00000574: người dùng xác nhận tạm tính thanh toán gốc trống = 0. Cộng dồn 1.839.000 đồng; giữ ô nguồn trống và ghi chú quy ước. Không áp dụng mọi ô trống.
- Đã đối chiếu mẫu hóa đơn số 3: cộng dồn 11.856.288 đồng; hai hóa đơn trong file mẫu cộng -11.856.288 đồng. Mẫu không được tính là đã phát hành.
- Bước tiếp: xác minh các ghi chú trong report mới, rồi xem phương án đưa chuỗi số 3 về 0 và nhóm đã điều chỉnh còn lại. Chưa tạo file import, ký hoặc phát hành.
- Mỗi nhóm giữ một file Excel đang dùng. Đóng Excel trước khi lưu đè; không tự tạo nhiều bản mới. Báo cáo cũ đã xóa theo yêu cầu, có thể tra lịch sử Git.

## Tám hóa đơn lệch tiền cần tiếp tục đối chiếu

Đã giải thích với người dùng ngày 08/10/2026: đây là lệch số học ngay trên từng hóa đơn trong **Bảng kê hóa đơn đã sử dụng**, giữa **cột R (Tổng tiền)** và **cột P (Doanh số bán chưa thuế) + cột Q (Tiền thuế GTGT)**. Cột P là doanh số sau chiết khấu, không phải cột N (Tổng tiền hàng). Chênh lệch dưới đây tính bằng R − (P + Q); chưa xác định số nào đúng bằng bản hóa đơn/XML và chưa chốt số tiền phải điều chỉnh.

| Ký hiệu | Số hóa đơn | P + Q (đồng) | R (đồng) | Chênh lệch có dấu (đồng) |
|---|---|---:|---:|---:|
| 1C25TUV | 00000649 | -411.999 | 411.999 | +823.998 |
| 1C25TUV | 00000561 | 348.640 | 212.640 | -136.000 |
| 1C25TUV | 00001819 | 8.209.652 | 8.207.652 | -2.000 |
| 1C25TUV | 00001167 | 671.701 | 670.701 | -1.000 |
| 1C25TUV | 00002145 | 42.329.900 | 42.329.000 | -900 |
| 1C25TUV | 00002619 | -636.718 | -636.719 | -1 |
| 1C25TUV | 00001821 | 6.157.151 | 6.157.152 | +1 |
| 1C25TUV | 00000626 | 4.695.046 | 4.695.045 | -1 |

Tổng độ lệch tuyệt đối là 963.901 đồng. Ưu tiên kiểm tra 00000649 (ngược dấu tổng tiền) và 00000561; không tự kết luận các chênh lệch 1 đồng do làm tròn. Cả tám có trong `data/Du lieu Meinvoice/Bang_ke_hoa_don_da_su_dung_2024_2025/Bang_ke_hoa_don_da_su_dung_1791380301707_1.xlsx`.

## Nhận bàn giao trên máy công ty

Trong thư mục repo trên máy công ty, chạy `git status --short`. Nếu sạch, chạy `git pull --ff-only origin master`, rồi mở thư mục đó làm project trong Codex và đọc phần trạng thái hiện tại ở đầu tài liệu này. Nếu có thay đổi cục bộ hoặc pull báo lỗi, nhờ Codex kiểm tra và giữ các thay đổi trước khi cập nhật; không reset/xóa chúng.

Tiếp tục xác minh tám lệch số học ở trên và căn cứ ngày 12885, đồng thời xem xét việc dùng số thanh toán 0 đã có trong bảng chi tiết cho bảy ô bảng tổng trống. Chưa áp dụng quy tắc bổ sung tiền này vào report/script. Sau khi xác minh các ghi chú, mới xem phương án đưa nhóm đã điều chỉnh ngoài ngoại lệ về 0 theo mục tiêu đã chốt. Chưa tạo import, ký hoặc phát hành hóa đơn.

Tái tính bằng `python scripts/prepare_invoice_reports.py`, rồi xuất cả hai report bằng `node scripts/layout-tools/build-adjusted.mjs` với Node và artifact-tool từ runtime Codex. Đã đối chiếu số tiền và dữ liệu nguồn từng hóa đơn với bảng tổng (ngày bảng chính 12885 có quy ước xác nhận riêng nêu trên), kiểm cộng dồn và quy ước 574, kiểm phạm vi, ô trống, bộ lọc, cố định tiêu đề và xem ảnh cả ba tab. Xem README.md. Script `prepare_adjusted_2024.py` và `run.py` là workflow cũ, không dùng để tái tạo hai report hiện tại.

---

## Bàn giao cũ – chỉ để tra cứu lịch sử

Các đường dẫn báo cáo cũ trong phần dưới có thể đã bị xóa. Trạng thái ở đầu tài liệu này thay thế phần bàn giao cũ.

# UVG – tiếp tục công việc trên máy ở nhà

Cập nhật ngày 06/10/2026. Repo: https://github.com/vansang2409/UVG, nhánh `master`.

Ngày 07/10/2026, dự án Git trên máy hiện tại đã được chuyển sang `D:\UVG`. Chọn thư mục này khi thêm project vào Codex. Hai thư mục `Du lieu Amis` và `Du lieu Meinvoice` ngay dưới thư mục dự án là dữ liệu gốc được giữ nguyên; Git bỏ qua chúng để tránh đưa lên hai lần. Bản sao nguồn được quản lý trong Git ở `data/`.

## Mục tiêu và quyết định đã thống nhất

Lập danh sách hóa đơn cần giữ nguyên, xuất mới, điều chỉnh hoặc thay thế, sau đó chuẩn bị Excel nhập vào meInvoice. Công ty bán máy lọc không khí qua TikTok, Shopee, Lazada và ngoài sàn, đang dùng meInvoice và USB token.

- AMIS do team làm lại sổ và khách hàng cung cấp, đã chốt là số liệu mục tiêu.
- meInvoice là lịch sử phát hành của kế toán cũ. Có hóa đơn gộp doanh thu và hóa đơn đã điều chỉnh/thay thế nhưng vẫn sai theo thông tin người dùng cung cấp.
- Không dựa vào mã đơn sàn để ghép dữ liệu. Số hóa đơn không phải mã đơn sàn.
- Đối chiếu hiện tại theo tháng, họ model và kênh khi có căn cứ nhận diện. Tổng theo nhóm chưa chứng minh một hóa đơn cụ thể phải sửa.
- Không xóa dữ liệu meInvoice. Dữ liệu CQT dùng xác minh lịch sử, không mặc định chứng minh nội dung doanh thu đúng.
- Các số 70.000 hóa đơn mới và 10.500 điều chỉnh/thay thế là quy mô công việc người dùng nêu ban đầu, chưa được xác nhận thành số lượng phải phát hành từ dữ liệu.

## Đã có gì

Trong repo có 11 file nguồn, script đối chiếu, báo cáo và bảng kiểm tra ưu tiên. Nguồn phân tích hiện tại chỉ bao gồm 2024 và tháng 1–6/2025, dù AMIS được làm lại từ 2020–2025. Chưa có dữ liệu CQT và các lần sửa sau kỳ xuất file.

| Chỉ tiêu theo bộ file hiện có | Giá trị (đồng) |
|---|---:|
| AMIS, tổng thanh toán sau giảm trừ | 93.194.013.744 |
| meInvoice, tổng thanh toán tính theo trạng thái/liên kết hiện có | 95.594.027.824 |
| AMIS trừ meInvoice | -2.400.014.080 |

Chênh lệch trên là số cần giải thích. Không lấy số này để tự xuất giảm hoặc xuất bù. AMIS dùng ngày chứng từ; meInvoice đang quy về kỳ hóa đơn gốc, có thể khác kỳ bán hàng.

Mở các file theo thứ tự:

1. `outputs/uvg-kiem-tra-uu-tien/KIEM_TRA_UU_TIEN.md`: số lượng hóa đơn riêng biệt và cách dùng bảng.
2. `outputs/uvg-kiem-tra-uu-tien/HOA_DON_CAN_KIEM_TRA_UU_TIEN.xlsx`: bảng làm việc, có bộ lọc và cột ghi kết quả.
3. `outputs/uvg-doi-chieu-nhom/UVG_DOI_CHIEU_NHOM.md`: báo cáo tổng thể, phương pháp và giới hạn.
4. `outputs/uvg-doi-chieu-nhom/SO_DU_THEO_CHUOI_HOA_DON.csv`: các hóa đơn đóng góp vào từng chuỗi.
5. `outputs/uvg-doi-chieu-nhom/DOI_CHIEU_THANG_MODEL.csv`: chênh lệch theo kỳ và họ model.

Các bảng tìm ứng viên cũ trong `outputs/uvg-doi-chieu/` chỉ dùng tra cứu. Chưa chứng minh ghép được từng giao dịch.

## Làm gì tiếp theo

### 1. Bổ sung bằng chứng cho danh sách ưu tiên

Bảng gom 299 hóa đơn thiếu liên kết nguồn, 14 hóa đơn thiếu số tiền trên bảng tổng và 8 hóa đơn tiền hàng cộng thuế lệch tổng thanh toán. Có thể một hóa đơn thuộc nhiều nhóm; số riêng biệt xem trong báo cáo ưu tiên.

Tra cứu đúng Ký hiệu + Số hóa đơn trên meInvoice. Lấy XML, bản hiển thị khi cần, danh sách và thông tin hóa đơn liên quan. Với hóa đơn thiếu liên kết, cần xác định hóa đơn gốc, các lần điều chỉnh/thay thế và bản hiện hành. Nếu có lần sửa sau tháng 6/2025, lấy thêm cả lịch sử này.

Đặt file bổ sung trong `data/Bo sung/meInvoice/`. Tạo thư mục khi có file. Trong Excel, ghi đường dẫn tương đối vào cột File bằng chứng đã bổ sung, rồi ghi kết quả. Giữ nguyên các cột dữ liệu nguồn; điền các cột làm việc phía cuối. Không tự điền tiền thiếu bằng 0 hoặc tự suy ra liên kết.

Có thể làm mẫu 10–20 hóa đơn trước để thống nhất cách đọc XML và ghi kết quả, sau đó xử lý phần còn lại theo lô. Đây là kiểm tra hồ sơ, chưa import hay phát hành.

### 2. Giải thích sản phẩm và lệch kỳ

Sau nhóm ưu tiên, xử lý 1.176 hóa đơn chưa phân bổ chắc chắn tiền theo sản phẩm, phần chưa nhận diện model và các tháng chênh lệch lớn. Một hóa đơn có thể có nhiều vấn đề. Không phân bổ tùy ý khoản điều chỉnh chung vào sản phẩm, không suy ra biến thể Pro/Light và không coi kênh chưa xác định là ngoài sàn.

Đối chiếu AMIS sau chiết khấu, trả lại, giảm giá với số dư hóa đơn gốc và các lần sửa. Muốn kết luận theo từng hóa đơn, cần tài liệu chứng minh những giao dịch nào đã nằm trong hóa đơn gộp đó. Nếu thiếu căn cứ thì để Chưa đủ căn cứ, không ép ghép những giao dịch chỉ giống hàng và giá.

### 3. Xác minh lịch sử CQT trước khi phát hành

Có thể bổ sung CQT sau trong quá trình phân tích. Trước khi chốt phát hành, cần đối chiếu lịch sử đầy đủ cùng MST/phạm vi thời gian, gồm các lần sửa sau kỳ xuất dữ liệu. Đặt file trong `data/Bo sung/CQT/` khi có.

Đối chiếu ký hiệu, số hóa đơn, ngày, trạng thái, tiền hàng, thuế, thanh toán và liên kết điều chỉnh/thay thế. Trạng thái truyền thành công không chứng minh AMIS và hóa đơn khớp nội dung.

### 4. Chốt bảng xử lý và tạo Excel import

Chỉ sau khi giải thích được lịch sử và giá trị, mới chốt từng trường hợp: Giữ nguyên / Xuất mới / Điều chỉnh / Thay thế / Chưa đủ căn cứ. Với trường hợp sửa, phải chỉ ra hóa đơn tham chiếu, lý do và nội dung đúng dựa trên bằng chứng.

Khi đó mới chọn mẫu meInvoice đúng trên môi trường công ty thực dùng, ánh xạ các trường và kiểm tra file import. Mẫu trên testapp3 chưa chứng minh cấu hình môi trường thật. Cần kiểm tra quy định áp dụng tại thời điểm xử lý trước khi chọn nghiệp vụ sửa và thời điểm phát hành.

## Mở dự án trên máy khác

Lần đầu:

```powershell
git clone https://github.com/vansang2409/UVG.git
cd UVG
```

Nếu đã clone và chưa có thay đổi cục bộ:

```powershell
git pull --ff-only origin master
```

Nếu đã ghi kết quả trong Excel, lưu bản làm việc trước khi pull. Không dùng lệnh reset để bỏ thay đổi. Mở Markdown và Excel ngay được, không cần chạy lại toàn bộ phân tích.

Nếu cần chạy lại, cài Python 3.11+ rồi làm theo `README.md`. Để tái tạo bảng ưu tiên từ CSV đối chiếu hiện tại:

```powershell
py scripts/prepare_priority_review.py
```

Lệnh này tái tạo CSV và thống kê, không cập nhật bản Excel có ghi chú. Không chạy đè lên ghi chú đã nhập trong CSV. Khi bổ sung dữ liệu nguồn, cần cập nhật xử lý và kiểm tra lại trước khi dùng kết quả mới.

## Nhắn Codex ở nhà

> Tôi đang tiếp tục dự án UVG. Hãy đọc TIEP_TUC_UVG.md, README.md và báo cáo kiểm tra ưu tiên trước. AMIS là số liệu khách hàng đã chốt; không ghép theo mã đơn sàn. Tiếp tục kiểm tra danh sách ưu tiên bằng các file bổ sung trong data/Bo sung, giữ nguyên dữ liệu nguồn và ghi rõ bằng chứng cho từng kết luận. Chưa tạo chỉ định phát hành khi còn thiếu lịch sử hoặc chưa giải thích được chênh lệch.

## Trạng thái thực hiện

### Tiếp tục ngày 07/10/2026 – rà soát bộ file hiện có

Đã rà 321/321 hóa đơn ưu tiên và xác nhận SHA-256 của 11 file nguồn khớp bản bàn giao. Kết quả ở `outputs/uvg-kiem-tra-uu-tien/ra-soat-2026-10-07/KET_QUA_KIEM_TRA.md`; bảng 321 dòng ở `KET_QUA_RA_SOAT_321.csv` trong cùng thư mục, giữ nguyên 21 cột cũ và thêm căn cứ/kết quả phía cuối. Excel và CSV ưu tiên ban đầu giữ nguyên.

299 hóa đơn vẫn thiếu liên kết trong các bảng DS; 39 hóa đơn trong nhóm này có số tham chiếu trong diễn giải, mới là ứng viên tra cứu, chưa xác nhận ký hiệu. 14 hóa đơn thiếu tiền chưa có bảng tổng khác bổ sung đủ giá trị. 8 lệch số học được xác nhận lại; `1C25TUV/00000649` có tiền hàng/thuế âm nhưng thanh toán dương cần ưu tiên XML. Có 4 hóa đơn không có dòng chi tiết. Chưa có dữ liệu trong `data/Bo sung`.

Bước tiếp theo: lấy XML cho lô 22 hóa đơn trong `outputs/uvg-kiem-tra-uu-tien/ra-soat-2026-10-07/LO_22_CAN_LAY_XML.csv`, gồm 8 lệch tiền và 14 thiếu tiền; lấy cả XML tham chiếu và lịch sử sửa sau 06/2025. Sau đó kiểm tra 39 hóa đơn có số tham chiếu trong diễn giải. Chưa thay đổi tổng đối chiếu, chưa chốt nghiệp vụ xử lý, chưa import/ký/phát hành. Script rà soát: `scripts/audit_priority_sources.py`.

Đã phân tích file và lập danh sách cần kiểm tra. Chưa kết nối CQT, chưa xác minh từng hóa đơn bằng XML, chưa chốt số lượng cần phát hành, chưa import, ký hay phát hành hóa đơn. Xem commit mới nhất trên GitHub để biết phiên bản tài liệu đã nhận.

Các file bàn giao đã được commit trong dự án. Ngày 07/10/2026, quyền chạy lệnh Git ngoài sandbox đã cho phép Codex hoàn tất bước commit. Trên máy ở nhà, nhận phiên bản mới bằng clone/pull như hướng dẫn ở trên, rồi kiểm tra:

```powershell
git log -1 --oneline
Test-Path TIEP_TUC_UVG.md
Test-Path outputs/uvg-kiem-tra-uu-tien/HOA_DON_CAN_KIEM_TRA_UU_TIEN.xlsx
```

Commit bàn giao có thông điệp `Add invoice priority review and home continuation guide`. Hai lệnh Test-Path cần trả về True. Commit dự án nền trước đó là `4e78390`.
