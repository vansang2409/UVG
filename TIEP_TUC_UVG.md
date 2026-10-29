# Cập nhật báo cáo Google Sheets theo phát hành thực tế — 30/10/2026

Theo yêu cầu cập nhật link báo cáo, đã sửa trực tiếp tab **Ra soat chuoi 2024**, gid **700000010**, trong Google Sheet **IMPORT MEINV 2024** (`1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE`). Giữ bố cục ba cột tóm tắt và phần 778 dòng chi tiết thu gọn.

- Đã kiểm lại toàn bộ 778 gốc trên bảng kê hiện hành và liên kết cũ đã xác minh: **654 gốc/1.253 HĐ điều chỉnh** và **6 chuỗi thay thế** đều đã phát hành, về 0 đủ ba khoản. 118 dòng ngoài diện xử lý gồm 116 gốc còn tiền, một HĐ hàng tặng 0đ và HĐ78 đã hủy; không đưa các dòng này về 0 bằng thao tác sửa báo cáo.
- Cập nhật A1:C13 để phân biệt phần cũ đã kiểm thực tế với phần bán mới/trả hàng còn dự tính. Đổi nhãn J/K/N và cập nhật N15:N792, P15:R792 bằng kết quả, liên kết, nguồn/dòng phát hành mới. Không đổi số tiền, công thức hay ghi vào các tab import.
- Tiền cũ giữ **622.963.131đ**, dự tính sau bán mới/trả hàng **49.930.540.760đ**, AMIS **49.928.171.800đ**, chênh **2.368.960đ** giữ nguyên. Chênh này là HĐ111 khác kỳ +3.888.889đ và 16 nhóm khác tiền hàng -1.519.929đ, khác với số dư các chuỗi xử lý đã về0. Phần bán mới AMIS 49.533.443.555đ và 163 phiếu trả hàng -225.865.926đ chưa được đối chiếu phát hành trong đợt kiểm này, không gọi tổng dự tính là doanh thu đã phát hành thực tế.
- Đã đọc lại **16.992 ô** trên tab chính và tab chênh: đúng **3.138 ô** cần cập nhật, không thay đổi ngoài phạm vi, không lỗi công thức; **1.163 kết quả công thức** giữ nguyên. Đã xem ảnh Google Sheets trên Chrome ở100%, bố cục tóm tắt hiển thị đủ. Hash 40 file nguồn không đổi.
- Tái lập nội dung: `scripts/prepare_report_post_issue_20261030.py`; căn cứ `scripts/data/report-post-issue-20261030/{before-cells,after-cells,prepared,requests,google-verification}.json`. Các `batch-*.json` chỉ là yêu cầu chia nhỏ tạm, được bỏ qua bởi Git. Dùng snapshot/kiểm tra bản live trước khi gửi lại requests, tránh ghi đè thay đổi về sau của người dùng. Phạm vi bàn giao Git đợt này gồm script chuẩn bị, năm JSON căn cứ/kiểm chứng, ghi chú và quy tắc bỏ qua batch tạm; không thay đổi workbook nguồn hoặc dữ liệu import.

---

# Chốt kiểm tra sau phát hành và bàn giao Git — 30/10/2026

Người dùng xác nhận KT đã import/phát hành và tải nguồn mới tại `data/Meinv_SAU_PHAT_HANH_20261029/`. Cập nhật30/10/2026 đã đối chiếu chỉ đọc đủ40file: 33phần bảng kê (159.589HĐ duy nhất), 5phần điều chỉnh (13.400liên kết), 2file thay thế (468liên kết). Bộ đầu chọn kỳ01/01/2024–29/10/2026; file bổ sung xuất theo ngàyHĐ thay thế29/10/2026. Đây là kết quả từ dữ liệu phát hành mới, thay cho các dự tính lịch sử bên dưới khi đánh giá đã về0 hay chưa.

- **Người dùng đã xác nhận KT loại `1C25TUV/00001943` vì TĐ không hợp lệ**, kèm ảnh trạng thái ngày29/10. STT1168 có trong file import online/cục bộ nhưng KT chủ động bỏ khỏi đợt phát hành. Tách giá trị nguồn597.269đ +47.782đVAT =645.051đ sang loại trừ; không còn ghi là thiếu xử lý, không sửa workbook/online/giá trị nguồn.
- **Đã chốt về0 đủ3khoản trong phạm vi KT thực hiện.** Ngày30/10 người dùng xác nhận trực tiếp **“trống là 0 á”** cho các ô tổng thanh toán đang treo. Áp dụng đúng24ô/15chuỗi đang rà (9điều chỉnh+6thay thế), giữ nguyên ô nguồn trống và lưu căn cứ `scripts/data/post-issue-20261029/confirmed-blank-payment-zero.json`; không áp dụng tùy tiện cho ô trống khác. Hai bộ điều chỉnh **9.424/9.424gốc** khớp0 (Đã điều chỉnh1.015/1.015, Chưa điều chỉnh8.409/8.409 sau loại1943), cộng6thay thế thành **9.430/9.430chuỗi trong phạm vi về0**. Không còn dòng cần xử lý.
- **13HĐ bổ sung đã xuất hiện**: `1C26MUV/00110692–00110704`, ngày29/10, đúng liên kết8gốc, đều ghi HĐ hợp lệ và khớp3khoản. `00110696` (STT2037) thanh toán nguồn trống được tính0 theo xác nhận30/10. 8chuỗi không còn dư tiền hàng/VAT233.262.118đ/18.660.968đ như lần kiểm trước.
- Tìm đủ **10.454/10.454HĐ điều chỉnh trong phạm vi** sau loạiSTT1168/gốc1943; toàn bộ HĐ năm2026 này có trạng thái nguồn HĐ hợp lệ. Chưa tra cứuCQT trực tiếp hoặc kiểm chữ ký/chi tiết dòng phát hành mới.
- **Đã xác nhận đủ6liên kết thay thế ngày30/10** qua file người dùng gửi `Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_1793293386111.xls`, lưu bản giốnghash ở `data/Meinv_SAU_PHAT_HANH_20261029/THAY_THE/`. Cặp kýhiệucũ1C25TUV → mới1C26MUV: **2601→110705;2610→110706;2602→110707;2603→110708;2604→110709;2617→110710** (sốHĐ chuẩn8chữsố). Ngàycũ31/05/2025, ngàymới29/10/2026; tổng6cũ371.344.277đ tiền hàng +29.707.543đVAT =401.051.820đ. Cả6mới tiền hàng/VAT0, thanh toán nguồn trống tính0 theo xác nhận30/10, bảng kê tổng ghiHĐ hợp lệ. Cả6chuỗi đã về0 đủ3khoản.
- Báo cáo/danh sách: `outputs/kiem-tra-sau-phat-hanh-20261029/KET_QUA_KIEM_TRA.md`, `CAC_TRUONG_HOP_CAN_KIEM_TRA.csv` (0dòng, chỉ tiêu đề), `LOAI_TRU_THEO_XAC_NHAN_KT.csv` (1gốc1943), `TAT_CA_CHUOI.csv` (9.431dòng, gồm9.430về0 +1loại trừ). Script `scripts/{audit_post_issue_20261029,verify_post_issue_20261029}.py`, căn cứ `scripts/data/post-issue-20261029/{audit,summary,verification,accountant-exclusions,confirmed-blank-payment-zero}.json`.
- Kiểm độc lập bằng openpyxl đã đọc lại 20.925 dòng hóa đơn ở 6 file nguồn, kiểm số dư/tình trạng từng nhóm, CSV và hash nguồn/kế hoạch: đạt. Không sửa workbook nguồn, import/phát hành lại hay ghi dữ liệu lên phần mềm.
- Phạm vi bàn giao Git theo yêu cầu ngày 30/10: 40 file nguồn mới, hai script đối soát/kiểm chứng, báo cáo và danh sách chi tiết, JSON căn cứ gồm loại trừ HĐ 1943 và xác nhận thanh toán trống bằng 0. Không đưa thư mục tạm, ảnh clipboard, bản tải kiểm tra Google Drive hoặc thông tin xác thực vào commit. Chạy lại `scripts/verify_post_issue_20261029.py` để kiểm nguồn và kết quả; `scripts/audit_post_issue_20261029.py` để tái lập báo cáo từ nguồn và các xác nhận đã chốt.

---

# File import thay thế về 0 cho 6 chuỗi — 29/10/2026

Người dùng cung cấp `data/Mau_hoa_don_GTGT_MTT_1_thue_suat_Thay_the.xls` và yêu cầu tạo import, chưa nhập meInvoice. Đã tạo đúng một file bàn giao: `outputs/amis-2024-theo-btl-bhh-20261026/IMPORT_THAY_THE_VE_0_6_CHUOI_2024.xlsx`, sheet `Hóa đơn GTGT`, đúng 22 cột A:V, tiêu đề hàng9.

- 6 tham chiếu HĐ cuối `1C25TUV/00002601, 00002610, 00002602, 00002603, 00002604, 00002617`, đều ngày31/05/2025; lần lượt thuộc gốc2024 số80,87,116,117,118,184. Số dòng tương ứng21,12,10,14,28,96; tổng181dòng.
- Giữ nguyên từng tên hàng và ĐVT của HĐ cuối. Đưa SL, đơn giá, thành tiền về0 ở từng dòng; VAT0 ở dòng đầu mỗi HĐ, thuế suất8 theo nguồn. Thông tin người mua và TM/CK giữ nguồn, tên đơn vị/MST nguồn trống giữ trống. CộtK “HĐ bị thay thế thuộc hệ thống khác” để trống vì HĐ cùng hệ thống, cộtO mãCQT trống vì nguồn không cung cấp. Không tự điền email/điện thoại/CCCD.
- Ngày hóa đơn thay thế mới (cộtB) để trống, cần KT chọn trước khi nhập. Tổng HĐ bị thay thế371.344.277đ tiền hàng +29.707.543đVAT =401.051.820đ; bản thay thế mới0đ. Không ghi số âm của hóa đơn điều chỉnh vào mẫu thay thế.
- Đã đối chiếu tổng chi tiết với tổng6HĐ, giữ số tham chiếu dạngtext8chữsố, đọc lại đủ181dòng với dữ liệu chuẩn bị, kiểm không còn dòng ví dụ hay công thức mẫu. Đã xem ảnhrender vùng tham chiếu/tiền/những dòng cuối; nguồn và mẫu không đổi. Bản xuất đầu có công thức mẫu rỗng còn sót nên builder dựng vùng nhập sạch theo đúng22tiêuđề và hướng dẫn mẫu; bản cuối đã kiểm đạt, SHA256 `749d2f2b3c8a486a0a791d07f726104e5baaa67435d490010652d92f352819ad`.
- Builder/kiểm chứng: `scripts/{prepare_replacement_zero_2024.py,build_replacement_zero_2024.mjs,verify_replacement_zero_2024.py}`, căn cứ `scripts/data/replacement-zero-20261029/{prepared,verification}.json`. Chưa thử nhập meInvoice, chưa ký/phát hành; báo cáo doanh thu vẫn là dự tính. Không sửa các file import/báo cáo khác; chưa commit/push đợt này.

---

# Làm gọn cách xem mục 2 — 28/10/2026

Theo phản hồi “khó xem quá”, đã sửa trực tiếp hai tab hiện tại, giữ nguyên số liệu:

- `Hang hoa theo thang 2024`: hai bảng riêng, mỗi bảng 4 cột. Doanh thu A3:D17; giá vốn A19:D34. Chi tiết 282 dòng chuyển xuống A38:N319, tổng hàng 320; khối hàng 37–320 thu gọn, mở dấu + khi cần xem mã hàng/số lượng.
- `Hang hoa can doi chieu`: A3:D8 liệt kê 5 việc cần bổ sung. 55 dòng chi tiết chuyển xuống A12:I66, khối hàng 11–66 thu gọn.
- Đã so toàn bộ 282 dòng và 55 dòng với trước sửa: giống hoàn toàn, kể cả kết quả công thức. Tổng DT AMIS 49.928.171.800đ, HĐ dự tính 49.930.540.760đ, chênh 2.368.960đ. Giá vốn đã có nguồn 24.542.992.457đ; vẫn thiếu giá vốn HĐ63/111, không kết luận đã khớp đầy đủ.
- Giữ chữ 11pt, số tiền phân cách hàng nghìn, tiêu đề xuống dòng; kiểm lại giá trị và định dạng qua Google Sheets API. Chưa xác minh ảnh render do kết nối trình duyệt không khả dụng như ghi chú trước. Không sửa import, không tạo thêm Excel bàn giao, chưa commit/push.

---

# Mục 2 — Hàng xuất bán theo tháng và giá vốn — 28/10/2026

Người dùng yêu cầu thống kê hàng xuất bán theo tháng từ sổ kế toán và hóa đơn, kèm giá vốn; đã chọn rõ **bộ hóa đơn dự tính sau xử lý**, không phải hóa đơn đã phát hành thực tế. Đã thêm đúng2tab trong cùng Google Sheet `IMPORT MEINV 2024`, không tạo Excel bàn giao:

- `Hang hoa theo thang 2024` (gid700000020), A1:N305: tổng hợp12tháng tại A5:H18; chi tiết **282 dòng tháng/mã hàng/ĐVT** tại A22:N305, có số lượng, doanh thu, giá vốn hai phía và chênh. Tách ĐVT, không cộng chung Cái/Chiếc/Sợi/Bộ/Lần; giữ cả dòng dịch vụ để tổng doanh thu khớp nguồn.
- `Hang hoa can doi chieu` (gid700000021), A1:I59: **55 dòng vấn đề**, gồm2HĐ thiếu giá vốn,26dòng hàng/combo chưa xác định mã SKU duy nhất,17dòng tiền điều chỉnh chưa phân bổ SKU,10dòng khácĐVT HĐ/AMIS. Một HĐ có thể có nhiều vấn đề, không cộng tổng tiền của tab này.

Phân tháng: bán mới và trả lại theo ngày chứng từ AMIS; HĐ cũ giữ lại theo tháng HĐ gốc2024, gồm các HĐ điều chỉnh/thay thế liên quan năm sau. Đây là phân tháng của phương án hiện tại, không gọi là báo cáo ngày phát hành thực tế. Doanh thu đều chưa VAT. Tổng AMIS **49.928.171.800đ**, HĐ dự tính **49.930.540.760đ**, chênh **2.368.960đ**, khớp báo cáo trước.

Giá vốn lấy nguyên cột AN của sổ AMIS (giá vốn trả hàng mang dấu âm). Giá vốn phần **đã có nguồn đối chiếu: 24.542.992.457đ**: bán mới24.342.047.362đ + HĐ cũ theo115BHH310.452.688đ − trả hàng109.507.593đ. Mỗi BHH chỉ dùng một lần, không cộng lặp khi gộp/tách hóa đơn. Giá vốn AMIS cả năm cũng24.542.992.457đ; giá vốn hóa đơn dự tính **chưa đầy đủ** vì HĐ111 (đơn BHH2023005090 thuộc2023) và HĐ63 (hàng tặng0đ, thiếu BHH) chưa có nguồn giá vốn. Không coi thiếu dữ liệu là0đ; tổng chênh GV cả năm để trống. Các ô giá vốn chưa phân bổ theo SKU/ĐVT để trống hoặc chỉ hiện phần đã ghép và đánh dấu; không tự chia theo doanh thu.

HĐ cũ cần thống kê hàng gồm **117gốc/186HĐ**, bao gồm cả HĐ63 có tiền0 nhưng SL1 hàng tặng (khác số116gốc còn tiền trong báo cáo doanh thu). Chi tiết HĐ cũ đều trống mã hàng; ghép qua danh sách BHH đã xác nhận, tên chuẩn hóa trùng, model/loại hàng duy nhất trong cùng BHH, hoặc BHH chỉ có một SKU. Dòng combo nhiều hàng, nhiều ứng viên, khácĐVT hoặc điều chỉnh tiền cho nhiềuSKU giữ riêng; không dùng ghép gần đúng hoặc tự quy đổi. Có **4.739.490đ** điều chỉnh tiền chưa phân bổ được theoSKU, vẫn cộng đủ vào doanh thu ở dòng riêng để tổng khớp.

Kiểm độc lập trực tiếp sổ gốc **44.578dòng**, từng tháng/mã/ĐVT, đủ30SKU AMIS; kiểm số tổng, dấu giá vốn trả hàng, sốBHH ghép và30hash nguồn. Đã đọc lại2tab:12tháng/282dòng/55vấnđề; không lỗi công thức, tổng DT/GV khớp. Đã kiểm định dạng native quaAPI, sửa định dạng số lượng về số nguyên để không còn dấu phẩy thừa. Chrome đã mất kết nối nên không chụp UI được; exportXLSX thành công nhưng tải bản xem thử bị lỗi TLS, không vô hiệu bảo mật. Không ghi nhận đã xem ảnh render. Không thay đổi các tab/bộ import hay ghi sổ/phát hành.

Mã và chứng cứ: `scripts/prepare_monthly_goods_2024.py`, `scripts/build_monthly_goods_requests.py`, `scripts/verify_monthly_goods_2024.py`, `scripts/data/monthly-goods-20261028/{prepared,verification,google-verification,layout}.json`. Phần thăm dò lưuprobe.json. Chưa commit/push.

---

# Hiển thị chi tiết HĐ 111 và 16 nhóm lệch — 28/10/2026

Theo yêu cầu người dùng, đã bỏ cả6dòng ghi chú cũ (hàng7–12) của `Doi chieu lech 2024`, thay bằng phần chi tiết nhìn thấy ngay, không phải mở nhóm ẩn:

- A8:F11: HĐ **1C24TUV/00000111**, chứng từ **BHH2023005090**, khách hàng AKIA; tiền hàng HĐ3.888.889đ, tiền AMIS trong phạm vi2024 bằng0, chênh+3.888.889đ. Ghi rõ HĐ lập15/03/2024 nhưng danh sách ghi đơn bán thuộc2023; không hiểu số0 là đơn năm2023 có giá trị0.
- A13:F31: bảng16nhóm có HĐ meInvoice gốc, mã BHH, tên khách hàng, tiền HĐ sau đủ điều chỉnh, tiền AMIS2024, chênh từng nhóm. Tổng chênh F31 **-1.519.929đ**.
- F32 tổng hai phần **2.368.960đ**. A33:F33 giải thích riêng HĐ64 có2609 giảm và9375 cộng lại482.407đ.

Phần tóm tắt A1:F6 vẫn giữ tổng chênh. Khối nguồn kiểm tra118dòng giữ nguyên giá trị và công thức, chuyển xuống header34/data35–152, thu gọn34–152; bộ lọc chuyển đúng khối này. Đã so118dòng sau với trước khớp toàn bộ; kiểm17trường hợp và tổngkhớp, không lỗi công thức; báo cáo doanh thu chính không đổi. Đã xem hiển thị, tăng chiều cao2dòng tên khách dài; chỉ cố định2hàng tiêu đề để không ghim nhầm tiêu đề tóm tắt trên bảng chi tiết. Không thêm tab/file Excel hoặc sửa import. Builder đã đổi cùng bố cục mới.

---

# Bỏ cùng dòng bốn cặp cũ ở tab chênh — 28/10/2026

Đã xóa thêm dòng7 "4 cặp cũ: số còn giữ lại" trong `Doi chieu lech 2024` theo ảnh người dùng. Giữ các ghi chú còn lại và118 dòng chi tiết; header nay hàng13, dữ liệu14–131, nhóm thu gọn13–131. Tổng chênh B6 vẫn2.368.960đ. Builder đã bỏ dòng và sửa mốc/công thức để không tạo lại. Tab báo cáo chính đã bỏ dòng tương tự ở lần trước; không đổi import hay số nguồn.

---

# Bỏ dòng bốn cặp cũ bằng 0 khỏi báo cáo — 28/10/2026

Theo yêu cầu người dùng đã xóa đúng dòng13 "4 cặp cũ: số còn giữ lại" của tab `Ra soat chuoi 2024`. Dòng giải thích chênh còn lại dịch lên13, dùng B13=B12; vẫn2.368.960đ, doanh thu sau dự tính vẫn49.930.540.760đ. Header chi tiết nay hàng14, dữ liệu15–792, nhóm chi tiết vẫn thu gọn. Đã đổi builder cùng cấu trúc để không tạo lại dòng đã bỏ; không đổi số liệu hay tab import.

---

# Báo cáo doanh thu đã tính lại sau pull — 28/10/2026

**Trạng thái hiện hành, thay thế các số dự tính cũ bên dưới.** Người dùng yêu cầu tính lại theo file vừa pull từ commit **9cb2cca** và cập nhật Google Sheets. Đã đọc lại toàn bộ nguồn, hai bộ import điều chỉnh cũ, bộ bán AMIS và phần trả hàng hiện tại. Giữ phạm vi **gốc năm 2024 và các điều chỉnh/thay thế liên quan năm sau**; số sau là dự tính khi hoàn tất nhập hóa đơn và phần KT thay thế bằng 0, chưa xác minh đã phát hành thực tế.

| Khoản | Doanh thu chưa VAT | VAT | Thanh toán |
|---|---:|---:|---:|
| Trước xử lý | 47.402.157.888 | 3.792.174.626 | 51.194.328.614 |
| Điều chỉnh về 0 theo file mới | -46.407.850.480 | -3.712.630.027 | -50.120.476.607 |
| Thay thế bằng 0, KT xử lý riêng | -371.344.277 | -29.707.543 | -401.051.820 |
| Hóa đơn cũ còn giữ lại | 622.963.131 | 49.837.056 | 672.800.187 |
| Bán mới AMIS | 49.533.443.555 | 3.962.675.787 | 53.496.119.342 |
| Điều chỉnh trả hàng | -225.865.926 | -18.069.275 | -243.935.201 |
| **Sau hoàn tất phương án, dự tính** | **49.930.540.760** | **3.994.443.568** | **53.924.984.328** |
| Sổ AMIS | 49.928.171.800 | 3.994.254.050 | 53.922.425.850 |
| **Chênh còn lại** | **2.368.960** | **189.518** | **2.558.478** |

Hàng cũ giữ lại là số dư sau hai khoản giảm, không cộng lại cùng dòng trước xử lý. Phần điều chỉnh mới có **1.253 hóa đơn dự thảo cho 654 gốc năm 2024**, thêm 8 hóa đơn thuộc 4 gốc so với báo cáo cũ; đây là phần năm 2024 trong bộ đã bổ sung13 HĐ trên8chuỗi ở cả2024/2025. Cả 4 cặp **3/508, 498/2140, 574/2189, 478/2578** đều đã được tính giảm về0 theo file, giảm số dự tính cũ **230.812.119đ** tiền hàng, **18.464.969đ** VAT và **249.277.088đ** thanh toán. Không giữ bốn cặp này trong doanh thu cũ và không trừ thêm lần nữa. Cũ giữ hiện chỉ còn116gốc thuộc danh sách công ty **622.963.131đ**. Giữ nguyên6chuỗi thay thế theo kế hoạch KT xử lý riêng371.344.277đ.

Chênh tiền hàng còn **2.368.960đ = 3.888.889đ HĐ111 liên quan đơn ghi nhận2023 − 1.519.929đ chênh ròng16nhóm đã có BHH**. Cần KT đối chiếu kỳ của HĐ111 và giá/toàn chuỗi của16nhóm. 15nhóm khác chỉ lệch thuế/thanh toán1đ. Chưa chỉnh AMIS hoặc hóa đơn để tự cân khoản này.

Đã cập nhật cùng Google Sheet `IMPORT MEINV 2024`:

- `Ra soat chuoi 2024`, A1:R793: báo cáo tóm tắt A1:C14, giữ riêng hai dòng **Điều chỉnh về0** và **Thay thế bằng0**; 778 dòng rà ẩn trong nhóm hàng15–793. B5 -46.407.850.480, B7 622.963.131, B10 49.930.540.760, B12 2.368.960, B13 0 (bốn cặp cũ), B14 2.368.960.
- `Doi chieu lech 2024`, A1:R132: tóm tắt chỉ còn hai nguyên nhân chênh; bốn cặp đổi thành **Đã điều chỉnh về0**, 118 dòng chi tiết vẫn giữ để kiểm chứng, thu gọn14–132. Cập nhật nguồn và STT/dòng import vừa pull cho4cặp; bỏ các ghi chú cũ yêu cầu xác minh478 hoặc chốt giữ3cặp.

Đã kiểm độc lập tập1.445HĐ và từng gốc, đầy đủ liên kết thay thế, số dư của từng4cặp bằng(0,0,0); 778chuỗi không thiếu liên kết. Đọc lại cả hai tab sau ghi: tổng tiền giữ/giảm khớp, không lỗi công thức, VAT và thanh toán chênh khớp, đã xem hiển thị Google Sheets. Giữ bố cục ba cột và chi tiết thu gọn theo yêu cầu; không thêm file Excel hay tab, không sửa ba tab import, không sửa các workbook vừa pull. Căn cứ: forecast/revenue-chain-audit/revenue-gap JSON hiện hành và `post-pull-google-verification.json`; các builder và kiểm chứng đã sửa theo quyết định19/10 để chạy lại không quay về số cũ. Chưa commit/push đợt tính lại này.

---

# Đã pull bản máy kia — 28/10/2026

Đã fast-forward từ fb6e1ff lên 9cb2cca và ghép đủ ghi chú chưa commit của máy này. Ghi chú mới trên máy kia cùng latest_override_20261019 xác nhận đã áp dụng 13 hóa đơn bổ sung từ 19/10; bao gồm đưa các cặp 3/508, 498/2140, 574/2189, 478/2578 về 0 trong bộ import hiện tại. Người dùng nhớ đúng về 478/2578: MST cá nhân, đã xử lý, không còn giữ theo bộ lọc 10 số. Các ghi chú/báo cáo của máy này trước khi pull đang dựa trên workbook cũ và giữ bốn cặp này, cần tính lại trước khi dùng. Chưa sửa dữ liệu import để giải quyết xung đột; giữ workbook mới từ remote và toàn bộ thay đổi local không trùng đường dẫn. Bản MD có cả lịch sử hai phía; các khẳng định giữ bốn cặp trong phần lịch sử không còn là trạng thái mới nhất. Chưa cập nhật lại hai tab báo cáo Google Sheets sau pull, chưa phát hành hóa đơn.

---

# Bàn giao đã hợp nhất công việc hai máy — 28/10/2026

Đã lưu công việc trên máy `D:\UVG` (nền `0801258`) vào commit **`f923a60dfabcfb13a56ef6e30eaf94d5e6df4786`**, gồm 49 file code, Markdown, workbook và dữ liệu kế toán. Người dùng đã đồng ý đưa cả workbook và dữ liệu kế toán lên repo GitHub công khai `vansang2409/UVG`. Chưa nhập phần mềm, ký hoặc phát hành hóa đơn. Các mục bên dưới lưu lịch sử; trạng thái “chưa commit/push” trong mục cũ là trạng thái tại thời điểm thực hiện công việc đó.

**Đồng bộ Git:** hai lần push đầu bị từ chối `non-fast-forward` vì GitHub đã có 3 commit ngày 26–27/10: `ec7e77f`, `30561f2`, `fb6e1ff` (đầu remote lúc pull: `fb6e1ffcd5efb4e8d4ec2b60e348665998d926d7`). Sau đó người dùng cho phép **“vậy pull đi”**; đã pull bằng merge, giữ commit của cả hai phía. Chỉ có xung đột văn bản ở `.gitignore` và file bàn giao này: đã kết hợp quy tắc ignore và giữ đầy đủ lịch sử hai phía, đặt ghi chú mới trước ghi chú cũ. Code và workbook ghép tự động; không tái tính hay sửa số liệu kế toán để giải quyết khác biệt nghiệp vụ. Bước cuối là push thường lên `origin/master` và kiểm tra HEAD local/remote khớp nhau; không reset/clean/force push. Kết quả push xem thông báo cuối của phiên bàn giao và Git, không suy ra đã đồng bộ chỉ vì có commit cục bộ.

## Đã làm và các bộ file đã lưu

- Đã áp dụng phương án 19/10: bổ sung **13 hóa đơn / 206 dòng / 8 chuỗi**, STT 2033–2045; bộ Đã điều chỉnh đạt **2.045 hóa đơn / 41.996 dòng / 1.015 chuỗi**. Cộng 8.410 hóa đơn Chưa điều chỉnh thành **10.455 hóa đơn / 55.564 dòng**; Ngoại lệ còn **236 chuỗi / 322 hóa đơn**. Đây là bộ điều chỉnh cục bộ, không phải bộ bán gốc AMIS 2024 mới trên GitHub.
- File chính: `outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx`, `outputs/HOA_DON_DA_DIEU_CHINH_IMPORT_2_SHEET.xlsx`, `outputs/NGOAI_LE_F0_2024_T6_2025.xlsx`, `outputs/doi-chieu-85-20261019/UVG_72_HOA_DON_KHONG_CAN_DIEU_CHINH_20261019.xlsx`. Hai bản Chưa điều chỉnh ở `outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx` và `outputs/HOA_DON_CHUA_DIEU_CHINH_IMPORT_2_SHEET.xlsx` giữ nguyên.
- Đã hợp nhất thêm bộ AMIS từ GitHub: `outputs/amis-2024-theo-btl-bhh-20261026/IMPORT_HOA_DON_MOI_AMIS_2024.xlsx`, `outputs/amis-2024-theo-btl-bhh-20261026/KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`, `outputs/btl-bhh-20261026/LIEN_KET_BTL_BHH_2024.xlsx`, cùng nguồn, script và JSON kiểm chứng. Ghi chú ngày 27/10 bên dưới mô tả bộ bán gốc **34.739 hóa đơn / 44.189 dòng**, import trả lại **161 phiếu / 216 dòng**, 2 BTL chờ KT. Các tham chiếu hóa đơn gốc/ngày điều chỉnh của phần trả lại còn trống; chưa phải dữ liệu đã nhập/phát hành. Ghi chú đối soát HD 2025 ở `GHI_CHU_DOI_SOAT_HD_2025_20261027.md` cũng được giữ.
- Quyết định máy đọc được: `scripts/data/check-review/kt-workbook-20261011/accountant-decisions.json`, mục `latest_override_20261019`. Kế hoạch, dữ liệu nguồn trích, bản trước sửa, script và kết quả kiểm tại `scripts/data/apply-13-20261019/`; căn cứ danh sách 72 tại `scripts/data/decision-85-20261019/`. Báo cáo hiện tại là `apply-13-20261019/final-verification.json`; báo cáo `decision-85-20261019/verification.json` là thời điểm trước khi cập nhật tổng hợp của bảng 72, không dùng hash cũ của báo cáo đó để kiểm bản hiện tại.
- Đã gộp đối soát Shopee/TikTok tháng 7–12/2025 tại `outputs/merge_20261012/ĐỐI SOÁT ĐÃ GỘP.xlsx`: **301 tệp / 602 sheet nguồn, 26.550 dòng / 50 cột / 1 sheet kết quả**. Giữ `rows.jsonl`, `merge_manifest.json`, dữ liệu kiểm nguồn và script cùng thư mục để tiếp tục trên máy khác. `validation.json` ghi PASS, không lệch giá trị ô/tổng tiền; các lỗi nguồn 1.197 `#N/A` và 6 `#REF!` được giữ dưới dạng chữ. SHA-256 workbook khớp `final_delivery.json`: `9385a2b97b37f721763be1195ce8893ded8525542f6b29ad25ceef9a90763eef`. Đường dẫn Downloads trong `final_delivery.json` là nơi giao bản cũ; bản bàn giao qua Git nằm trong `outputs/merge_20261012/`.
- Mẫu đối soát công nợ sàn: `outputs/doi_soat_san_20261019/Mau_doi_soat_cong_no_san_theo_thang.xlsx`. Bản này được chép nguyên byte từ Downloads trong đợt bàn giao, SHA-256 `e58390dc4db69400005c3d36de0fd255d4760b9d9f7a755d69d5e27767a666de`. Script dựng/kiểm Excel và báo cáo có sẵn tại `outputs/doi_soat_san_20261019/support/`; báo cáo cũ ghi 17 kiểm tra logic và 8 kiểm tra Excel. Đợt bàn giao chỉ kiểm file/hash, không tuyên bố đã chạy lại các kiểm tra Excel đó.
- Đã rà hash ngày 28/10: cả 4 workbook của đợt áp dụng 13 hóa đơn khớp `final-verification.json`; 12 nguồn và 2 workbook Chưa điều chỉnh khớp hash trong `plan.json`. Không dựng lại hoặc thay đổi số liệu workbook trong đợt bàn giao. Giữ thêm script kiểm lịch sử `scripts/layout-tools/verify-reports.py`; script này dựa trên cache/baseline cũ, không dùng nó để kết luận bộ 19/10 hiện tại sai.

## Cặp 478/2578: có bằng chứng nhưng đang mâu thuẫn giữa hai máy

**Bằng chứng cục bộ ngày 19/10:** `decision-85-20261019/data.json` ghi `1C24TUV/00000478` và `1C25TUV/00002578` thuộc người mua Đào Duy Tùng, mã nguồn **8823260350**, lý do “MST cá nhân; bổ sung điều chỉnh trái dấu”. `accountant-decisions.json/latest_override_20261019` đánh dấu đã áp dụng, thay phân loại ngoại lệ theo mã 10 số của chuỗi này. `apply-13-20261019/plan.json` có STT **2042** giảm gốc 478 (SL âm, đơn giá có sẵn dương), STT **2043** đảo riêng khoản tăng 2578 (chỉ tiền, để trống SL/đơn giá). Báo cáo `final-verification.json` ghi nguồn + các khoản mới của chuỗi 478 đạt **(0, 0, 0)** tiền hàng/thuế/thanh toán, và hash workbook hiện tại khớp báo cáo. Như vậy **bản cục bộ đã xử lý cả cặp**, không còn là phương án chưa áp dụng.

**Bằng chứng từ phiên bản GitHub `fb6e1ff`:** mục “Làm rõ lý do giữ 10 chuỗi trong dự tính” ngày 27/10 vẫn mô tả 478/2578 là ngoại lệ do `tax_shaped` nhận mã 10 số, **chưa xác minh loại người mua** và không phải cặp được KT chỉ định giữ riêng như 3/508, 498/2140, 574/2189. Ghi chú này khác quyết định đã áp dụng cục bộ ngày 19/10. Đã giữ nguyên cả hai nguồn bằng chứng khi merge Git; chưa tự chọn một phương án để ghi đè phương án kia. Cần đối chiếu hai lịch sử và chốt phạm vi trước khi tiếp tục nghiệp vụ hoặc hợp nhất số liệu workbook/dự toán. Không suy ra loại người mua chỉ từ độ dài mã số.

## Việc còn lại khi tiếp tục

1. Khi chuyển máy, kiểm tra nhánh `origin/master` chứa cả commit cục bộ `f923a60` và đầu remote cũ `fb6e1ff`; đọc phần bàn giao 28/10 và các mục ngày 27/10 ngay dưới trước khi chạy script. Xung đột Git đã được xử lý bằng cách giữ cả hai phía; các khác biệt nghiệp vụ ở mục tiếp theo vẫn cần đối chiếu riêng.
2. Đối chiếu riêng quyết định 478/2578 nêu trên, đồng thời rà các chuỗi 3/508, 498/2140, 574/2189 mà bản 19/10 đã xử lý nhưng ghi chú remote ngày 27/10 vẫn gọi là giữ nguyên. Các tổng 10.455 của máy này và 34.739 hóa đơn bán gốc AMIS trên remote thuộc hai bộ khác nhau, không cộng/trừ lẫn hoặc coi là cùng một file hiện hành.
3. Với bộ cục bộ, tiếp tục lưu ý 2 dòng gốc 498 (import 41865 và 41915) có thành tiền nguồn khác SL × đơn giá làm tròn 1 đồng; xem `rounding-review.json`, không tự tính lại thành tiền. Giữ các hóa đơn chỉ đổi thông tin 6108/4203/4204/9186, không tự lập đảo tiền; không chạy lại builder/prepare/merge/promote lịch sử chỉ để kiểm.
4. Nguồn gộp 301 tệp gốc được script tham chiếu ở `C:\Users\thuong ngo\Downloads\ĐỐI SOÁT CẦN GỘP`, ngoài repo; snapshot hàng dữ liệu `rows.jsonl` và workbook kết quả được bàn giao, còn muốn kiểm lại trực tiếp từng tệp gốc cần có lại thư mục nguồn. Một số script đối soát còn đường dẫn tuyệt đối Windows, cần sửa đường dẫn khi dùng trên máy khác. Mẫu đối soát sàn là file mẫu; nhập nguồn thực tế và kiểm lại trước khi dùng.
5. Không đưa cache review `.codex_review/`, `tmp/`, `node_modules/`, `__pycache__/`, ảnh preview, workbook chụp mẫu và log xuất tạm vào commit. Giữ dữ liệu kế toán, JSON căn cứ, workbook trước sửa cần đối chiếu, code và Markdown. Không nhập, ký hoặc phát hành hóa đơn trong phiên bàn giao này.

---

# Tách hai dòng đưa hóa đơn cũ về 0 — 28/10/2026

Theo yêu cầu người dùng, `Ra soat chuoi 2024` đã tách hàng5 **Điều chỉnh về 0: -46.177.038.361đ** và hàng6 **Thay thế bằng 0: -371.344.277đ**, thay cho một dòng gộp. Cũ giữ tại B7 vẫn853.775.250đ, sau dự tính B10 vẫn50.161.352.879đ, chênh B12 vẫn233.181.079đ. Chèn đúng một hàng, công thức/nhóm chi tiết dịch theo: header hàng15, dữ liệu16–793, nhóm thu gọn15–793. Builder đã cập nhật mốc mới và hai loại riêng. Đã đọc lại các số chính và xem hiển thị. Không thay đổi import hay phát hành hóa đơn.

---

# Làm gọn cả tab đối chiếu chênh — 28/10/2026

Theo phản hồi tiếp của người dùng, đã sửa `Doi chieu lech 2024` A1:C13 còn ba cột: khoản cần kiểm tra, tiền chênh, việc cần làm. Tách rõ 6 khoản: 498/2140 215.558.148đ; 3/508 10.978.045đ; 574/2189 1.702.778đ; 478/2578 2.573.148đ; HĐ111 3.888.889đ; 16 nhóm khác -1.519.929đ; tổng **233.181.079đ**. Đưa HĐ498 lên đầu để ưu tiên. Thu gọn hàng14–132, giữ nguyên118 dòng chi tiết và công thức. Không thêm file/tab.

Vì đổi vị trí bảng tóm tắt, đã cập nhật công thức phụ thuộc tại `Ra soat chuoi 2024!B12` thành `=SUM('Doi chieu lech 2024'!B4:B7)`; B12 vẫn230.812.119đ, B13 vẫn2.368.960đ. Số dự tính và số import không đổi. Cả hai builder đã đổi cùng nội dung mới; đã kiểm số và xem hiển thị Google Sheets.

---

# Trình bày lại báo cáo cho dễ đọc — 28/10/2026

Người dùng phản hồi bảng rà chuỗi quá rối. Đã sửa ngay tab `Ra soat chuoi 2024`, A1:C13 thành báo cáo ba cột: bước xử lý, số tiền, giải thích. Gộp hai phần giảm cũ thành **46.548.382.638đ**; giải thích bên cạnh gồm 46.177.038.361đ theo import và 371.344.277đ KT làm thủ công. Hiển thị rõ cũ giữ 853.775.250đ + bán mới 49.533.443.555đ − trả hàng 225.865.926đ = sau dự tính **50.161.352.879đ**, so sổ AMIS **49.928.171.800đ**, chênh **233.181.079đ**. Hai dòng cuối tách 230.812.119đ từ 4 cặp cũ giữ và 2.368.960đ khác kỳ/tiền hàng.

Đã bỏ đoạn lịch sử nối chuỗi dài khỏi đầu báo cáo. Nhóm và thu gọn hàng 14–792 (header + 778 dòng rà), dữ liệu và công thức chi tiết vẫn nguyên, mở lại bằng dấu + bên trái khi cần. Chỉ cố định 3 hàng đầu. Không thêm tab/file, không sửa tab/import hay số nguồn. Đã xem hiển thị ở Google Sheets và kiểm số chính. Builder `scripts/build_chain_review_requests.py` đã đổi cùng nội dung để chạy lại không đưa phần giải thích rối trở lại.

---

# Truy nguyên chênh 233.181.079đ — 28/10/2026

Đã đối chiếu đủ 120 chuỗi cũ giữ lại với 115 BHH AMIS bị loại theo danh sách. Quy liên kết thay thế 60/61 về cùng gốc 60 để không tách sai một khoản cộng và một khoản trừ. Danh sách công ty thành 114 nhóm liên kết, cộng thêm 4 chuỗi ngoài danh sách: 118 dòng đối chiếu. Toàn bộ chênh doanh thu đã phân rã được, **chưa thay đổi import hay quyết định giữ nguyên**.

| Nguyên nhân theo phạm vi so sánh | Tiền hàng chênh (HĐ cũ giữ − AMIS đã loại) |
|---|---:|
| Ba cặp KT chốt giữ: 3/508, 498/2140, 574/2189 | +228.238.971 |
| Cặp 478/2578 được bộ lọc giữ do ô MST/CCCD có 10 chữ số | +2.573.148 |
| HĐ 1C24TUV/111: danh sách ghi đơn BHH2023005090 thuộc 2023 | +3.888.889 |
| 16 nhóm đã có BHH ghép, chênh tiền hàng cộng ròng | -1.519.929 |
| **Tổng** | **+233.181.079** |

Chi tiết 4 chuỗi chưa có trong danh sách BHH AMIS đã loại:

| Gốc 1C24TUV | Điều chỉnh 1C25TUV | Tiền hàng còn lại | Căn cứ giữ |
|---|---|---:|---|
| 3 | 508 | 10.978.045 | KT chốt giữ nguyên CC01 |
| 498 | 2140 | 215.558.148 | KT chốt giữ nguyên CC02 |
| 574 | 2189 | 1.702.778 | KT chốt giữ nguyên CC03 |
| 478 | 2578 | 2.573.148 | Bộ lọc tax_shaped nhận 8823260350; chưa xác minh loại người mua |

Nguồn quyết định 3 cặp: `scripts/data/check-review/kt-workbook-20261011/accountant-decisions.json`, điểm 6; workbook `NGOAI_LE_F0_2024_T6_2025.xlsx`, Ngoai le dòng 241–243. Cặp 478 tại dòng 78, giữ vì quy tắc mã 10 số, **không phải** một cặp được KT chỉ định riêng như CC01–CC03. Các chuỗi này chưa có BHH đối chiếu trong danh sách loại. Vì vậy chưa chứng minh từng khoản trùng với bán mới; không tự ghép theo tên/số tiền, không tự đưa về 0. Hướng xử lý: giữ quyết định đang có với CC01–CC03, xin bảng mã BHH/đơn gốc để đối chiếu phần xuất mới và xác định đơn cần loại; riêng 478 cần KT xác minh định danh và BHH trước khi chốt giữ/xử lý.

HĐ 111 / AKIA **3.888.889đ**: nguồn `data/HD Xuất CTY 2024.xlsx`, HD cty 2024 dòng 101, ghi rõ đơn BHH2023005090, note "Đơn hàng ghi nhận 2023, xuất hóa đơn 24 => sai thời điểm". Sổ AMIS năm 2024 không có mã 2023 này. Đây là khác kỳ của hai tập so sánh cần KT kiểm tra, không tự cộng vào AMIS 2024 hoặc tự giảm hóa đơn.

16 nhóm có BHH và lệch tiền hàng ròng **-1.519.929đ** (AMIS lớn hơn ròng). Trong đó đáng chú ý **64/BHHT2024002867 lệch +482.407đ**: toàn chuỗi meInvoice có điều chỉnh 2609 **-482.407đ** rồi 9375 **+482.407đ**, nên còn 2.407.407đ. Bảng HD cty chỉ trình bày tới 2609 và ghi 1.925.000đ; không dùng số bảng này thay thế phép cộng đủ chuỗi. Các nhóm còn lại cần đối chiếu giá/chiết khấu và giá trị chứng từ, không suy nguyên nhân phí sàn khi chưa có bằng chứng. Ngoài ra 15 nhóm chỉ lệch thuế/thanh toán 1đ, không làm thay đổi chênh doanh thu chưa VAT.

Đã bổ sung một tab trong cùng Google Sheet: [Doi chieu lech 2024](https://docs.google.com/spreadsheets/d/1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE/edit#gid=700000011), A1:R132: tổng nguyên nhân, 118 dòng đối chiếu, từng BHH, tiền hàng/VAT/thanh toán, nguồn và hướng xử lý. Kiểm đọc lại công thức: tiền cũ 853.775.250đ; AMIS loại 620.594.171đ; chênh 233.181.079đ; chênh VAT 18.654.487đ; chênh thanh toán 251.835.566đ; kiểm phân rã bằng 0. Đã xem hiển thị Google Sheets, không lỗi công thức. Không thêm file Excel.

Mã/chứng cứ: `scripts/audit_revenue_gap_2024.py`, `scripts/prepare_revenue_gap_review.py`, `scripts/data/revenue-gap-20261028/prepared.json`, `review-sheet.json`, `google-verification.json`. Hash các nguồn trong manifest của bộ AMIS vẫn khớp. Hai workbook import vẫn giữ nguyên hash; chưa nhập/ký/phát hành hay commit/push. Số dự tính **50.161.352.879đ** chưa thay đổi sau lần phân tích này.

---

# Chốt rà liên kết và báo cáo dự tính 2024 — 28/10/2026

**Đã giải quyết 6 liên kết còn thiếu bằng file người dùng vừa xuất.** Nội dung này thay thế trạng thái chưa chốt ở phần lịch sử bên dưới. Nguồn đã lưu nguyên bản tại `data/Du lieu Meinvoice/Bang_ke_hoa_don_thay_the_xuat_theo_HD_bi_thay_the_1793158686512.xls`; SHA256 `4c015c740185be2c856838c47667d42a02c6bd16c40ed6584010efb961b05b4a`, khớp file Downloads. Dù đuôi .xls, nội dung là OOXML; đọc theo nội dung, không đổi/sửa nguồn.

| Gốc 1C24TUV | Thay thế trung gian 1C24TUV | HĐ cuối 1C25TUV được tính | Doanh thu chưa VAT của HĐ cuối |
|---|---|---|---:|
| 80 | 188 | 2601 | 36.869.444 |
| 87 | 189 | 2610 | 17.660.647 |
| 116 | 190 | 2602 | 16.788.889 |
| 117 | 191 | 2603 | 28.913.888 |
| 118 | 192 | 2604 | 44.300.922 |
| 184 | 216 | 2617 | 226.810.487 |
| Tổng 6 chuỗi | | | 371.344.277 |

Các chặng gốc→trung gian nằm tại dòng 9–14, chặng trung gian→cuối tại dòng 15–20 của nguồn mới. Đây chính là 6 chuỗi đã nằm trong kế hoạch KT đưa về 0, **không phải 6 khoản doanh thu cộng thêm**. Bỏ giá trị cả gốc và trung gian, chỉ cộng HĐ cuối; không cộng 403.872.599đ của 6 gốc đã bị thay thế.

Kết quả sau quy về gốc thật: **778 chuỗi** (thay cho 784 nhóm khi còn thiếu cạnh), gồm 658 chuỗi điều chỉnh, 7 chuỗi thay thế, 113 chuỗi không phát sinh liên kết. Trong đó 777 chuỗi có hóa đơn đóng góp, 1 gốc hủy; **1.445 hóa đơn được cộng**, không cộng trùng, không còn chuỗi thiếu liên kết trong dữ liệu hiện có. Chuỗi công ty Phú Minh 1C24TUV/60→61→1C25TUV/2605 được tính riêng 2.536.111đ và giữ lại; không thuộc 6 chuỗi thủ công.

Phạm vi số liệu: **gốc năm 2024 và các hóa đơn điều chỉnh/thay thế liên quan lập năm sau**, không phải chỉ các hóa đơn phát hành trong năm dương lịch 2024. Trước xử lý theo nguồn; sau xử lý là dự tính khi hoàn tất các bộ import và phần KT thủ công, chưa chứng minh đã nhập/ký/phát hành thực tế.

| Khoản | Doanh thu chưa VAT | Thuế GTGT | Tổng thanh toán |
|---|---:|---:|---:|
| Trước xử lý | 47.402.157.888 | 3.792.174.626 | 51.194.328.614 |
| Điều chỉnh cũ theo bộ import | -46.177.038.361 | -3.694.165.058 | -49.871.199.519 |
| KT đưa 6 chuỗi thay thế về 0 | -371.344.277 | -29.707.543 | -401.051.820 |
| Cũ giữ lại sau cả hai phần trên | 853.775.250 | 68.302.025 | 922.077.275 |
| Bán mới AMIS | 49.533.443.555 | 3.962.675.787 | 53.496.119.342 |
| Điều chỉnh trả hàng | -225.865.926 | -18.069.275 | -243.935.201 |
| **Sau hoàn tất phương án (dự tính)** | **50.161.352.879** | **4.012.908.537** | **54.174.261.416** |
| Cao hơn AMIS | 233.181.079 | 18.654.487 | 251.835.566 |

Lưu ý hàng "Cũ giữ lại" là số dư sau điều chỉnh, không cộng thêm cùng cả hàng "Trước xử lý". Số trước có chênh tiền hàng + VAT - thanh toán +3.900đ từ nguồn, giữ nguyên; số sau khớp tiền hàng + VAT = thanh toán. Chênh AMIS 233.181.079đ vẫn cần đối chiếu nhóm giữ/loại (853.775.250đ cũ giữ lại trừ 620.594.171đ AMIS đã loại theo danh sách), không gọi là đã cân AMIS.

Đã cập nhật đúng tab [Ra soat chuoi 2024](https://docs.google.com/spreadsheets/d/1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE/edit#gid=700000010), A1:R792: đủ 778 chuỗi, quan hệ và chứng cứ mới, bỏ cảnh báo thiếu, xóa 6 dòng thừa cũ A793:R798. Kiểm đọc lại số liệu/công thức: tổng G 47.402.157.888đ, J -46.177.038.361đ, K -371.344.277đ, L 853.775.250đ; không lỗi công thức; 6 chuỗi đều về 0 theo kế hoạch. Không thay đổi 3 tab import và không sửa các workbook import; không tạo thêm Excel bàn giao.

Đã sửa việc quy gốc của 6 chuỗi trong `scripts/forecast_revenue_2024.py` và `scripts/audit_revenue_chains_2024.py`, để báo cáo theo tháng/ngày gốc không dùng nhầm ngày hóa đơn trung gian. Đối chiếu độc lập tập 1.445 hóa đơn và tổng tiền từng gốc giữa hai cách tính; kiểm nguồn mới đủ cả hai chặng cho 6 chuỗi. Chứng cứ: `scripts/data/revenue-chain-audit-20261028/closure-verification.json` (passed), `summary.json` (provisional=false, issues=[]), báo cáo forecast hiện hành đã tính lại. Còn việc KT thực hiện đưa 6 chuỗi về 0, map số hóa đơn thật cho phần trả hàng và đối chiếu chênh AMIS; không còn yêu cầu lấy bổ sung 6 liên kết thay thế. Chưa commit/push.

---

# Lịch sử rà độc lập theo chuỗi — 28/10/2026: trạng thái thiếu liên kết đã được giải quyết ở trên

**Nội dung này thay thế kết luận về tính đầy đủ của báo cáo ngay bên dưới.** Đã rà 784 gốc năm 2024 bằng cạnh điều chỉnh/thay thế riêng: điều chỉnh cộng dồn, thay thế chỉ cộng nhánh cuối còn hiệu lực. 1.445 hóa đơn thuộc phần đã nối được vẫn cho doanh thu **47.402.157.888đ**, nhưng đây chưa phải tổng đầy đủ đã xác minh. Lần tính trước loại gốc có trạng thái bị thay thế mà chưa kiểm đủ hóa đơn kế tiếp của 6 gốc sau.

| Gốc 1C24TUV thiếu liên kết thay thế | Ngày gốc | Tiền hàng gốc chỉ để tra cứu |
|---|---|---:|
| 00000080 | 01/03/2024 | 36.649.565 |
| 00000087 | 01/03/2024 | 17.256.573 |
| 00000116 | 20/03/2024 | 31.965.946 |
| 00000117 | 20/03/2024 | 49.111.340 |
| 00000118 | 20/03/2024 | 57.649.733 |
| 00000184 | 26/04/2024 | 211.239.442 |
| Tổng giá trị gốc, KHÔNG cộng vào doanh thu | | 403.872.599 |

Chưa có căn cứ lấy 403.872.599đ làm tiền hóa đơn thay thế. Cần tìm đủ số/ký hiệu, trạng thái, giá trị và quan hệ tới hóa đơn cuối của từng gốc, bao gồm các điều chỉnh liên quan nếu có. Đây khác với 6 chuỗi đã xác định trong kế hoạch KT đưa về 0 (gốc 188, 189, 190, 191, 192, 216).

Đã bổ sung tab [Ra soat chuoi 2024](https://docs.google.com/spreadsheets/d/1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE/edit#gid=700000010), đủ 784 dòng: hóa đơn được cộng/bị loại, cạnh liên kết, nguồn và dòng, tiền từng chuỗi, kế hoạch điều chỉnh. 6 chuỗi thiếu được tô vàng và để trống doanh thu, không ghi bằng 0. Kiểm đọc lại đủ 784 dòng, 6 dòng chờ, không lỗi công thức; tổng cột G 47.402.157.888đ, J -46.177.038.361đ, K -371.344.277đ, L 853.775.250đ. Không thay đổi ba tab import hay các workbook import; không thêm file Excel.

**Báo cáo tạm của phần đã nối được:** trước 47.402.157.888đ; điều chỉnh cũ -46.177.038.361đ; KT dự kiến đưa về 0 -371.344.277đ; cũ giữ lại 853.775.250đ; bán mới +49.533.443.555đ; trả lại -225.865.926đ; sau dự tính 50.161.352.879đ. Chênh AMIS 233.181.079đ cũng là tạm tính theo phạm vi này. Chưa chốt tổng trước/sau toàn bộ khi 6 gốc trên còn thiếu liên kết.

Đã đọc meInvoice đúng UVGREEN/MST 0316520076 trên Chrome Ke Toan (XYZ), nhưng công cụ không chụp/nhận thao tác bộ lọc được dù người dùng đưa cửa sổ ra trước. Chưa lấy được 6 quan hệ từ màn hình; không sửa chứng từ, không phát hành. Hướng tiếp theo: lấy báo cáo quan hệ thay thế có 6 gốc năm 2024 hoặc đọc trực tiếp khi Chrome đó kết nối được công cụ. Đã quét nguồn cục bộ, 472 XML và báo cáo thay thế 2025 trên Drive; chưa có 6 quan hệ này. Mã và chứng cứ: scripts/audit_revenue_chains_2024.py, scripts/data/revenue-chain-audit-20261028/{prepared,summary,review-sheet}.json. Chưa commit/push.

---

# Báo cáo dự tính doanh thu 2024: trước và sau xử lý — 28/10/2026 (bản trước khi rà đủ liên kết)

Phạm vi: **các chuỗi hóa đơn gốc năm 2024, cộng cả hóa đơn điều chỉnh/thay thế liên quan lập năm sau**. Số trước lấy theo dữ liệu meInvoice đã xuất đang có; số sau là dự tính từ các bộ import hiện tại và kế hoạch KT đưa 6 chuỗi thay thế về 0. Chưa xác minh thực tế đã import/ký/phát hành hoặc KT đã làm xong 6 chuỗi. Đơn vị: đồng, doanh thu chưa VAT.

| Chỉ tiêu | Doanh thu chưa VAT |
|---|---:|
| Trước xử lý | 47.402.157.888 |
| Giảm theo bộ điều chỉnh hóa đơn cũ | -46.177.038.361 |
| Cộng import bán gốc AMIS 2024 | 49.533.443.555 |
| Trừ import trả lại: 163 phiếu / 219 dòng | -225.865.926 |
| Sau các bộ import, trước phần KT làm thủ công | 50.532.697.156 |
| KT đưa 6 chuỗi thay thế về 0 theo kế hoạch | -371.344.277 |
| **Sau hoàn tất toàn bộ phương án** | **50.161.352.879** |
| **Tăng so với trước xử lý** | **2.759.194.991** |

Doanh thu thuần AMIS **49.928.171.800đ**; số sau toàn bộ phương án còn cao hơn **233.181.079đ**. Phép đối chiếu: phần hóa đơn cũ giữ lại sau thủ công **853.775.250đ**, trong khi phần bán AMIS bị loại theo danh sách **620.594.171đ**; chênh đúng **233.181.079đ**. Cần tiếp tục đối chiếu phần giữ/loại này, không gọi số dự tính là đã cân khớp AMIS. Không sửa các workbook import khi lập báo cáo.

Sau toàn bộ phương án: VAT **4.012.908.537đ**, thanh toán **54.174.261.416đ**, khớp doanh thu + VAT. Số trước theo nguồn: VAT **3.792.174.626đ**, thanh toán **51.194.328.614đ**; nguồn có chênh số học tiền hàng + VAT - thanh toán **+3.900đ** như đã ghi trong lịch sử, giữ nguyên nguồn, không tự cân.

Căn cứ mới tính lại: `scripts/forecast_revenue_2024.py`, `scripts/data/revenue-forecast-20261028/prepared.json` và `summary.json`. Đã đọc lại các nguồn và import hiện hành; hash nguồn nguyên, 650 gốc trong diện điều chỉnh cũ về 0, không hồ sơ thiếu tiền/đa gốc; kiểm độc lập phép cộng trước/sau và chênh AMIS. Báo cáo này thay số dự toán 27/10 vì đã tách bán/trả và bổ sung hai BTL theo Giá ghi trả. Chưa commit/push đợt cập nhật này.

---

# Bàn giao cuối ngày 28/10/2026 — trạng thái hiện hành

**Đã hoàn tất hôm nay:**

- Người dùng chốt BTL00023 và BTL00024 **theo đúng Giá ghi trả AMIS**, không quy đổi về giá bán phân bổ, không thêm khoản giảm giá/chiết khấu để cân tiền. Đã đưa cả hai vào import điều chỉnh: BTL00023 giảm **2.679.000đ**; BTL00024 giảm **219.000đ**.
- `Import dieu chinh` trên Google Sheets và workbook kiểm tra hiện đủ **163 phiếu / 219 dòng**; tiền hàng **-225.865.926đ**, VAT **-18.069.275đ**, tổng thanh toán **-243.935.201đ**. Giữ mẫu A:X, Y **Mã đơn hàng**, Z **Số chứng từ bán gốc BHH**; số phiếu BTL ở Lý do điều chỉnh. Tổng tiền chỉ ghi ở dòng đầu mỗi phiếu.
- Đã bỏ tab **Cho KT tra lai** khỏi Google Sheets và Excel vì không còn dòng chờ chốt giá. Builder/converter không tạo lại tab trống.
- Kiểm tra đủ 219 dòng trả theo nguồn, tổng tiền khớp; 216 dòng cũ trên Google giữ nguyên khi thêm 3 dòng của hai phiếu. Xóa tab trống không thay đổi các phần workbook khác. File bán gốc và công thức IFERROR cột S trên Google giữ nguyên.

**File đang dùng:** `outputs/amis-2024-theo-btl-bhh-20261026/IMPORT_HOA_DON_MOI_AMIS_2024.xlsx` và `KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`. Không sinh thêm bản Excel bàn giao. Google: [IMPORT MEINV 2024](https://docs.google.com/spreadsheets/d/1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE/edit), còn ba tab **Hóa đơn GTGT**, **Mã đơn hàng**, **Import dieu chinh**.

**Cần làm tiếp:** lập hóa đơn bán gốc trước; lấy ký hiệu/số/ngày hóa đơn thực tế rồi map vào phần trả lại, bổ sung ngày hóa đơn điều chỉnh và mã CQT nếu nghiệp vụ yêu cầu, kiểm lại trước khi nhập. Các trường tham chiếu này hiện vẫn trống; **chưa import, ký hoặc phát hành**. Hai BTL trên đã chốt, không còn là hồ sơ chờ quyết định giá. 7 mục loại trừ chưa rõ và phần KT xử lý hóa đơn thay thế thủ công vẫn cần theo dõi; dự toán doanh thu cũ chưa cập nhật theo luồng bán/trả tách riêng.

**Trạng thái Git:** commit gần nhất đã push là **fb6e1ff** (bộ chốt ngày 27/10). Các thay đổi ngày **28/10** và ghi chú này **chưa commit/push**. Căn cứ kiểm chứng trong `scripts/data/amis-2024-linked-20261026/`: `user-confirmed-return-prices.json`, `google-confirmed-return-prices-verification.json`, `return-adjustment-verification.json`, `verification.json`. Các mục bên dưới giữ làm lịch sử; ưu tiên trạng thái bàn giao này.

---

# Đã bỏ tab Cho KT tra lai — 28/10/2026

Theo yêu cầu người dùng, đã xóa tab **Cho KT tra lai** (đã hết dữ liệu) khỏi Google Sheets và workbook kiểm tra. Google còn `Hóa đơn GTGT`, `Mã đơn hàng`, `Import dieu chinh`. Excel giữ đủ **163 phiếu / 219 dòng điều chỉnh**, các phần workbook khác kiểm nguyên byte; file bán gốc giữ nguyên. Builder/converter đã sửa để không tạo lại tab chờ khi không có dòng chờ. Những mục ghi tab chờ bên dưới là lịch sử. Chưa commit/push thay đổi ngày 28/10.

---

# Đã chốt BTL00023/BTL00024 theo Giá ghi trả — 28/10/2026

Người dùng xác nhận: **“trả theo Giá ghi trả nha”**. Đã đưa BTL00023 và BTL00024 vào `Import dieu chinh` của Google Sheets và workbook kiểm tra, giữ đúng SL, đơn giá trả, tiền hàng, VAT và tổng trả nguồn AMIS. Không tính lại theo giá phân bổ lúc bán, không tạo dòng giảm giá bổ sung và không đổi phần bán gốc.

BTL00023: KA150SFULL SL -1, giá 1.733.199đ, thành tiền -1.733.199đ; KA150SFILTERFULL SL -1, giá **747.356đ**, thành tiền -747.356đ; tổng tiền hàng **-2.480.555đ**, VAT **-198.445đ**, thanh toán **-2.679.000đ**. BTL00024: KM100FULL SL -1, giá **202.778đ**, tiền hàng **-202.778đ**, VAT **-16.222đ**, thanh toán **-219.000đ**.

Import điều chỉnh hiện đủ **163 phiếu / 219 dòng**, tổng tiền hàng **-225.865.926đ**, VAT **-18.069.275đ**, thanh toán **-243.935.201đ**. Sheet `Cho KT tra lai` đã hết dòng chờ chốt giá. Ngày điều chỉnh và tham chiếu HĐ gốc vẫn trống, chờ người dùng map sau; chưa import/ký/phát hành. Cột Y mã đơn, Z BHH giữ nguyên. Google chỉ chèn 3 dòng, kiểm 216 dòng cũ giữ nguyên dữ liệu; không đổi sheet bán gốc hoặc công thức S.

Căn cứ: `user-confirmed-return-prices.json`, `return-adjustment-verification.json`, `google-confirmed-return-prices-verification.json`, `verification.json` trong `scripts/data/amis-2024-linked-20261026/`. Hai khoản chênh 100.000đ/32.659đ là chênh so với giá phân bổ, đã có quyết định dùng giá trả nguồn; **không còn là hai hồ sơ chưa chốt giá**. Những mục phía dưới ghi chờ KT là lịch sử trước quyết định này. Chưa commit/push đợt cập nhật 28/10.

---

# Google Sheets: cột S không còn lỗi tra cứu #N/A — 27/10/2026

Theo yêu cầu người dùng, chỉ sửa công thức cột **S — Mã đơn Hàng** của sheet **Hóa đơn GTGT**, Google `IMPORT MEINV 2024`, ID `1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE`, gid **555142675**. Phạm vi **S10:S44198** (44.189 dòng). Mẫu: `=IFERROR(INDEX('Mã đơn hàng'!B:B;MATCH(T10;'Mã đơn hàng'!A:A;0));"")`. Mã chưa tìm được hiển thị trống, không tự đặt mã thay thế.

Đã bọc **43.967** công thức INDEX/MATCH; **222** công thức vốn có IFERROR giữ nguyên. Đọc lại toàn bộ 44.189 công thức và kết quả: đúng từng dòng, không còn lỗi #N/A/#REF!/#VALUE!/#ERROR!. Chỉ cập nhật `userEnteredValue` ở S, giữ tiêu đề, định dạng, các cột khác, các sheet khác và file Excel. Căn cứ `scripts/data/amis-2024-linked-20261026/google-s-iferror-verification.json`.

**Trạng thái cần tiếp tục:** bán gốc 34.739 HĐ / 44.189 dòng; import trả lại 161 phiếu / 216 dòng theo mẫu, Y mã đơn/Z BHH; 2 BTL lệch giá vẫn tách chờ KT. Người dùng sẽ map số/ký hiệu/ngày HĐ gốc và ngày điều chỉnh sau khi lập hóa đơn bán gốc. Chưa import/ký/phát hành. Các mục bên dưới là lịch sử quyết định, ưu tiên mục mới nhất.

Để dựng lại trên máy khác, snapshot đầu vào trung gian được lưu nén tại `scripts/data/new-amis-2024-20261026/prepared.json.gz` (giữ bản JSON lớn cục bộ); script rebuild tự đọc bản nén khi không có JSON. Thứ tự: `rebuild_import_2024_linked_returns.py` → `build_import_2024_linked_returns.mjs --replace` → `export_import_2024_linked_returns.py --replace` (gồm chuyển mẫu điều chỉnh trả lại) → `verify_import_2024_linked_returns.py`. Các chỉnh sửa Google Sheets đã kiểm lưu riêng, không được coi file Excel là bản sao chứa công thức tra cứu của Google.

---

# Sheet trả lại đã đổi thành Import dieu chinh theo mẫu — 27/10/2026

Theo yêu cầu mới, đã thay `Tra lai cho dieu chinh` bằng **Import dieu chinh** trong Google `IMPORT MEINV 2024` (giữ gid **700000001**) và workbook kiểm tra hiện tại. **A:X = đúng 24 tiêu đề từ `data/Du lieu Meinvoice/Mau HD Dieu chinh.xls`; Y = Mã đơn hàng; Z = Số chứng từ bán gốc BHH**. BTL ghi trong Lý do điều chỉnh. Mã đơn hàng ghép từ tab Mã đơn hàng; Google dùng công thức theo Z, Excel giữ chuỗi nguồn đầy đủ, không đổi thành số. Không lấy số BHH làm số hóa đơn pháp lý.

**161 phiếu / 216 dòng**, tiền hàng **-223.182.593đ**, VAT **-17.854.608đ**, thanh toán **-241.037.201đ**. SL âm, đơn giá dương, thành tiền âm; tổng tiền hàng/VAT/thanh toán chỉ ở dòng đầu mỗi phiếu. Giữ nguyên số nguồn. **Ngày điều chỉnh, ký hiệu/số/ngày HĐ gốc và mã CQT vẫn trống** để người dùng map sau; chưa sẵn sàng nhập khi chưa bổ sung tham chiếu.

**BTL00023 và BTL00024** (3 dòng, tổng trả nguồn 2.898.000đ) giữ riêng trong sheet **Cho KT tra lai**, không đưa vào import điều chỉnh vì khoản lệch 100.000đ/32.659đ chưa chốt bản chất; chưa tự chiết khấu hoặc đổi giá. Tổng 216 dòng import + 3 dòng chờ = 219 dòng BTL nguồn.

File import bán gốc giữ nguyên hash **c6ea974aeabc67e9912d6359448f853f4257fe79a808e3f16cf60d82f6f321ee**. Workbook kiểm tra cập nhật tại chỗ; các phần không liên quan giữ nguyên byte, không thêm Excel bàn giao. Căn cứ `return-adjustment-import.json`, `return-adjustment-verification.json`, `return-order-map.json` và `verification.json` trong `scripts/data/amis-2024-linked-20261026/`. Script `prepare_return_adjustment_import.py` / `build_return_adjustment_import.mjs` đã nối vào exporter để lần dựng lại tiếp tục giữ phương án này. Chưa import/ký/phát hành.

---

# Phương án hiện hành: tách bán gốc và trả lại — 27/10/2026

Người dùng làm việc với kế toán và xác nhận: **lập hóa đơn bán gốc trước, lấy ký hiệu/số/ngày hóa đơn thực tế rồi lập import điều chỉnh trả lại**. Quyết định này thay thế phương án bù trừ BTL vào import mới phía dưới. Không áp dụng minh họa giảm giá 100.000đ / 32.659đ; hai khoản này vẫn chờ kế toán chốt ở phần trả.

Đã cập nhật tại chỗ đúng hai file chính tại `outputs/amis-2024-theo-btl-bhh-20261026/`: import bán gốc **34.739 hóa đơn / 44.189 dòng**, tiền hàng **49.533.443.555đ**, VAT **3.962.675.787đ**, thanh toán **53.496.119.342đ**. Giữ nguyên từng dòng bán AMIS, phục hồi 159 BHH trả hết, hai BHH từng bù trừ một phần và đưa hai đơn BHHS2024000629/BHHS2024005761 vào bán gốc nguyên số. **115 chứng từ loại trừ vẫn loại**; 7 mục loại trừ chưa rõ giữ nguyên. A:Q đúng mẫu, R khuyến mại, S số chứng từ; không thêm bản Excel bàn giao.

Trong workbook kiểm tra có tab **Tra lai cho dieu chinh**, gồm **163 BTL / 219 dòng**, tiền hàng trả **225.865.926đ**, VAT **18.069.275đ**, tổng trả **243.935.201đ**. Đây là bảng chuẩn bị từ nguồn, **chưa phải file import điều chỉnh sẵn dùng**: ký hiệu/số/ngày hóa đơn gốc và ngày điều chỉnh đều để trống. Hai BTL00023/BTL00024 giữ nguyên giá/tiền nguồn và ghi rõ chênh lệch; không tự phân bổ lại hoặc tự giảm thuế. Không sửa sổ AMIS hoặc workbook liên kết.

Đã cập nhật Google Sheets `IMPORT MEINV 2024`, ID `1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE`: giữ cột S **Mã đơn Hàng**, cột T **Số chứng từ** và tab nguồn Mã đơn hàng; khôi phục 218 dòng, tổng 44.189 dòng / 34.739 nhóm liền nhau, tiền hàng/VAT khớp Excel. Các dòng khôi phục đều tra được mã đơn hàng. Thêm tab `Tra lai cho dieu chinh` đúng 219 dòng nguồn, chờ hóa đơn gốc. Không ghi đè các dòng bán không liên quan; STT cũ của Google giữ nguyên nên có khoảng trống 30731 và nhóm mới nối cuối, khác thứ tự đánh lại liên tục của Excel nhưng không trùng nhóm.

Kiểm độc lập đã qua: toàn bộ dòng import giữ đúng nguồn, hash nguồn nguyên, tổng bán/thuế/thanh toán khớp; 219 dòng trả đúng nguồn và bốn trường hóa đơn/ngày điều chỉnh trống; không lỗi công thức. Căn cứ `scripts/data/amis-2024-linked-20261026/verification.json` và `google-split-verification.json`. **Các dự toán doanh thu trước đây chưa cập nhật theo luồng tách bán/trả; không dùng tổng import bán gốc như doanh thu sau điều chỉnh.** Chưa import, ký, phát hành hoặc commit/push.

---

# Đã bổ sung BTL00145 → BHHT2024024972 theo người dùng — 27/10/2026

Người dùng xác nhận **BTL00145 trả cho BHHT2024024972**. Đối chiếu đúng **SP000042 / Bình nước giữ nhiệt thông minh UVGREEN UK100 Màu trắng / Cái / SL 1 / đơn giá và tiền hàng 972.222 / VAT 77.778 / thanh toán 1.050.000đ**. Chứng từ bán ngày 28/11/2024, trả ngày 03/12/2024, khớp toàn bộ: **trả hết**. Đã loại BHH này khỏi import mới, đưa sang nhóm `Đã trả hết`, bỏ BTL00145 khỏi `Can kiem tra`.

Đã cập nhật **tại chỗ** file liên kết `outputs/btl-bhh-20261026/LIEN_KET_BTL_BHH_2024.xlsx` và đúng 2 file chính tại `outputs/amis-2024-theo-btl-bhh-20261026/`. Hiện **34.578 hóa đơn / 43.971 dòng**, tiền hàng **49.304.852.397**, VAT **3.944.388.494**, thanh toán **53.249.240.891đ**. Tổng trả hết **159 BHH / 211 dòng**; **163 phiếu BTL / 219 dòng đều đã có BHH xác định** (218 dòng từ AMIS + 1 dòng người dùng xác nhận), không còn BTL thiếu liên kết. Còn **2 chứng từ / 4 dòng cần kiểm tra**: BHHS2024000629/BTL00023 và BHHS2024005761/BTL00024; 7 mục loại trừ chưa rõ giữ nguyên. Giữ đủ A:Q mẫu, R khuyến mại, S số chứng từ.

**Nguồn liên kết phải phân biệt:** giữ nguyên `verified-live-reads.json` và sổ AMIS gốc, không sửa chứng từ trên AMIS. Bổ sung tại `scripts/data/btl-bhh-live-20261026/user-confirmed-links.json`; cột trạng thái của dòng BTL00145 và sheet Thong tin ghi rõ do người dùng xác nhận 27/10/2026, snapshot AMIS lúc trích vẫn trống. Không được mô tả BHH này là đã đọc trực tiếp từ AMIS. Giữ hash snapshot file liên kết cũ tại `original_snapshot_sha256` trong báo cáo kiểm mới; bộ import dùng hash workbook liên kết cập nhật đã kiểm thay hash a59c... cũ. Đủ **11 xác nhận** trong dữ liệu chuẩn bị và hướng dẫn của file kiểm tra.

Kiểm hoàn tất độc lập cả file liên kết và 2 file import/kiểm tra: 219 dòng trả khớp đủ trường nguồn; BHH mới khớp trả hết; mọi dòng import khớp nguồn/bù trừ; đủ 35.017 chứng từ / 44.578 dòng phân nhóm, không thiếu/trùng; 159 BHH trả hết không có trong import; source hashes không đổi; không lỗi ô/công thức; hai file cuối giữ đúng hash bộ đã kiểm. Báo cáo `scripts/data/btl-bhh-live-20261026/verification.json` và `scripts/data/amis-2024-linked-20261026/verification.json`. Đã dọn thư mục xuất tạm, không tạo bản Excel bàn giao khác. Chưa nhập/ký/phát hành hoặc commit/push.

Dự toán doanh thu cũng cập nhật giảm phần import mới đúng **972.222đ doanh thu / 77.778đ VAT / 1.050.000đ thanh toán**: theo file import hiện tại, tổng **50.529.971.924 / 4.042.398.062 / 54.572.369.986đ**. Nếu KT hoàn tất kế hoạch đưa 6 chuỗi thay thế về 0 thủ công, còn **50.158.627.647 / 4.012.690.519 / 54.171.318.166đ**, chênh AMIS **230.455.847 / 18.436.469 / 248.892.316đ**. Phần KT thủ công vẫn là kế hoạch chưa xác minh đã thực hiện. Cập nhật căn cứ tại `scripts/data/revenue-forecast-20261027/`, không đổi doanh thu cũ hoặc tự xử lý 4 chuỗi giữ nguyên.

---

# Kiểm tra hóa đơn công ty trong nhóm thay thế — 27/10/2026

Đã kiểm bảng tổng, bảng liên kết thay thế và `Chi tiet nguon` của workbook ngoại lệ. **Có một chuỗi thay thế công ty thuộc gốc 2024:** `1C24TUV/00000060 → 1C24TUV/00000061 → 1C25TUV/00002605`, người mua **CÔNG TY CỔ PHẦN XÂY DỰNG VÀ NỘI THẤT PHÚ MINH**, MST **0106357063**. 60/61 đều đã bị thay thế; 2605 là hóa đơn thay thế còn hiệu lực, tiền hàng **2.536.111**, VAT **202.889**, thanh toán **2.739.000đ**. Chuỗi này đã thuộc ngoại lệ công ty và được giữ trong phần doanh thu cũ; **không nằm trong 6 chuỗi KT dự kiến đưa về 0**, không trừ thêm khi cập nhật dự toán.

Sáu chuỗi kế hoạch thủ công (188/2601,189/2610,190/2602,191/2603,192/2604,216/2617) đều có tên người mua khách lẻ không lấy hóa đơn, MST trống ở cả gốc và thay thế trong bảng tổng; không xuất hiện trong tab `HD cty 2024`. Đây là kết quả theo nguồn hiện có, không xác minh pháp nhân bên ngoài. Giữ nguyên số liệu dự toán sau thủ công **50.159.599.869đ doanh thu**, không sửa import.

---

# Kế hoạch KT đưa 6 chuỗi thay thế về 0 thủ công — 27/10/2026

Người dùng xác nhận: **“HĐ thay thế kt sẽ làm thủ công đưa về 0”**. Áp dụng nhóm 6 chuỗi 188/2601, 189/2610, 190/2602, 191/2603, 192/2604, 216/2617 đã liệt kê; đây là **kế hoạch xử lý thủ công**, chưa xác minh KT thực hiện hoặc phát hành. Không bổ sung nhóm này vào import điều chỉnh, không sửa workbook/import hiện tại.

Dự tính **sau khi hoàn tất cả bộ import và phần KT xử lý thủ công**: trừ thêm doanh thu **371.344.277đ**, VAT **29.707.543đ**, thanh toán **401.051.820đ**. Hóa đơn cũ còn **853.775.250đ doanh thu**, **68.302.025đ VAT**, **922.077.275đ thanh toán**. Cộng import mới: doanh thu **50.159.599.869đ**, VAT **4.012.768.297đ**, thanh toán **54.172.368.166đ**. Chênh với AMIS giảm còn **231.428.069đ doanh thu**, **18.514.247đ VAT**, **249.942.316đ thanh toán**; doanh thu chênh gồm **230.812.119đ** của 3 cặp chốt giữ + cặp 478 ngoại lệ theo mã 10 số, và **615.950đ** chênh còn lại. Giữ kết quả chỉ theo file import là kịch bản cũ, không ghi đè hoặc gọi số mới là đã thực hiện.

Căn cứ lưu tại `including_manual_replacement_zero` trong `scripts/data/revenue-forecast-20261027/prepared.json` và `summary.json`, có đủ 6 khóa gốc, số giảm và trạng thái kế hoạch. Chưa nhập/ký/phát hành hoặc commit/push.

---

# Làm rõ lý do giữ 10 chuỗi trong dự tính — 27/10/2026

Đối chiếu lại nguồn quyết định: nhóm 4 chuỗi được gọi chung là giữ nguyên trước đó **không cùng một lý do**. **3/508 (CC01), 498/2140 (CC02), 574/2189 (CC03)**: phản hồi KT ghi “Giữ nguyên, không xử lý điều chỉnh hóa đơn này”; người dùng xác nhận ngày 11/10 giữ cả cặp ngoài đợt xử lý, áp dụng vào workbook ngày 12/10. Không có ghi chú nguyên nhân nghiệp vụ chi tiết cho từng cặp trong phản hồi này; không suy ra tất cả là hóa đơn công ty/hộ kinh doanh chỉ từ quyết định giữ.

Riêng **478/2578**: người mua **Đào Duy Tùng**, ô MST/CCCD nguồn **8823260350**. Chuỗi được đưa vào `Ngoai le` bởi quy tắc bộ lọc `tax_shaped` nhận mã 10 số, không phải CC01–CC06. Chưa xác minh bản chất người mua là công ty/hộ kinh doanh hay cá nhân. Không gọi đây là một cặp đã được kế toán chỉ định giữ riêng như ba cặp trên.

**6 chuỗi thay thế 188/2601, 189/2610, 190/2602, 191/2603, 192/2604, 216/2617**: ngoài phạm vi workflow điều chỉnh đã dựng; ghi chú cũ nêu rõ nghiệp vụ thay thế chưa mở xử lý. Đây là còn nguyên vì chưa được xử lý trong bộ import, không phải xác nhận cuối rằng không cần xử lý. Dự tính chỉ cộng hóa đơn thay thế còn hiệu lực, không cộng gốc đã bị thay thế. Số dự tính không đổi; chưa sửa import hoặc quyết định giữ của chuỗi nào.

Căn cứ: `scripts/data/check-review/kt-workbook-20261011/accountant-decisions.json` điểm 6, `accountant-reply-review.json` CC01–CC03; `Chi tiet nguon` trong workbook ngoại lệ; quy tắc `tax_shaped` tại `scripts/filter_invoice_exceptions.py`; mục lịch sử workflow thay thế phía dưới.

---

# Dự tính doanh thu 2024 sau import điều chỉnh và import mới — 27/10/2026

Đã tính **chỉ đọc**, không sửa workbook/source/import, theo ba file import hiện hành (hai bộ `HOA_DON_*_IMPORT_2_SHEET.xlsx` và bộ mới AMIS). Phạm vi chính là **hóa đơn gốc có ngày năm 2024**, cộng đủ các hóa đơn liên quan 2025/2026 theo ký hiệu + số HĐ; loại hóa đơn hủy và hóa đơn bị thay thế, giữ hóa đơn thay thế còn hiệu lực. Đọc cả liên kết 2026 xuất 09/10 và 22/10; không thấy giá trị 2026 hiện hữu thuộc gốc 2024. Giữ liên kết đã xác minh tại `Chi tiet nguon` của workbook, gồm quan hệ lịch sử bị thiếu trong bảng liên kết xuất mới; không tính trùng 6 chuỗi có mặt ở cả bảng điều chỉnh và ngoại lệ.

**Kết quả dự tính theo gốc 2024:**

| Hạng mục | Doanh thu trước VAT | VAT | Tổng thanh toán |
|---|---:|---:|---:|
| Hóa đơn cũ hiện có, kể cả điều chỉnh năm sau của gốc 2024 | 47.402.157.888 | 3.792.174.626 | 51.194.328.614 |
| Phần import điều chỉnh cho gốc 2024 | -46.177.038.361 | -3.694.165.058 | -49.871.199.519 |
| Còn lại từ hóa đơn cũ sau xử lý | 1.225.119.527 | 98.009.568 | 1.323.129.095 |
| Import mới AMIS 2024 | 49.305.824.619 | 3.944.466.272 | 53.250.290.891 |
| **Dự tính tổng** | **50.530.944.146** | **4.042.475.840** | **54.573.419.986** |

Có 1.245 hóa đơn điều chỉnh dự thảo cho 650 gốc năm 2024; mọi gốc trong diện xử lý đều cộng về 0. Import mới 34.579 HĐ. Phần cũ còn giá trị gồm: 116 gốc thuộc danh sách `HD cty 2024` còn doanh thu **622.963.131đ**, 4 chuỗi đã chốt giữ nguyên (3,478,498,574) **230.812.119đ**, 6 chuỗi thay thế ngoài bộ import điều chỉnh (188,189,190,191,192,216) **371.344.277đ**. Không tự bỏ thêm các chuỗi giữ nguyên hoặc kết luận trùng với đơn AMIS.

**Lưu ý phạm vi ngày:** Các điều chỉnh dự thảo đang mang ngày 08/10/2026 (1.202 HĐ), 09/10/2026 (26), 12/10/2026 (17). Số 50.530.944.146đ là theo chuỗi gốc 2024, không phải phép cộng chỉ các hóa đơn có ngày lập năm 2024. Nếu chỉ cộng hóa đơn còn hiệu lực mang ngày 2024 trong bộ nguồn hiện có rồi cộng import mới mang ngày 2024, dự tính **92.308.124.327đ doanh thu**, **7.384.650.247đ VAT**, **99.692.774.574đ thanh toán**; các điều chỉnh dự thảo ngày 2026 sẽ nằm ngoài khoảng lọc này. Đây là dự tính từ file, chưa xác minh báo cáo sau nhập/phát hành trên phần mềm.

**Chưa khớp hoàn toàn với AMIS:** Doanh thu thuần AMIS (doanh số trừ hàng trả) **49.928.171.800đ**, VAT **3.994.254.050đ**, thanh toán **53.922.425.850đ**. Dự tính theo chuỗi cao hơn lần lượt **602.772.346đ**, **48.221.790đ**, **650.994.136đ**. Phần lớn đến từ 4 chuỗi giữ nguyên + 6 chuỗi thay thế ngoài danh sách công ty (doanh thu **602.156.396đ**); còn khác biệt tiền cũ giữ lại với AMIS loại trừ và 3 chứng từ chưa đưa vào import. Không gọi đây là doanh thu đã chốt/đã phát hành, chưa sửa phương án vì chỉ được yêu cầu dự tính báo cáo.

Nguồn tiền trống được giữ nguyên; dùng đúng quy ước 0 đã xác nhận cho 574/671/744, trường hợp còn lại bổ sung từ giá trị có thật ở nguồn chi tiết, không đoán mọi ô trống = 0. Các số liệu nguồn cũ có tổng sai số học tiền hàng + thuế - thanh toán **+3.900đ**, phần đảo theo số đã phát hành là **-3.900đ**, bù trừ về 0 theo phương án đã chốt; không tự cân số nguồn. Tổng sau điều chỉnh và tổng dự tính khớp tiền hàng + VAT = thanh toán. Đã kiểm tổng 12 tháng, không gốc xử lý còn dư, không hồ sơ thiếu tiền/đa gốc, hash nguồn không đổi. Phân tích: `scripts/forecast_revenue_2024.py`; căn cứ `scripts/data/revenue-forecast-20261027/prepared.json`, `summary.json`, `calendar-check.json`. Không tạo thêm Excel báo cáo.

---

# Đã điền số chứng từ AMIS vào cột S — 27/10/2026

Theo yêu cầu người dùng, cập nhật trực tiếp `outputs/amis-2024-theo-btl-bhh-20261026/IMPORT_HOA_DON_MOI_AMIS_2024.xlsx`: thêm **S9 = Số chứng từ**, điền mã chứng từ AMIS dạng chuỗi tại **S10:S43981**, lặp trên mọi dòng hàng cùng hóa đơn. Nguồn ghép là `Doi chieu import` của file kiểm tra: ghép đúng STT hóa đơn và số dòng Excel, không suy đoán theo khách hàng/tiền hàng. Đã kiểm đủ **43.972 dòng**, mỗi mã khớp bảng đối chiếu. Toàn bộ ô A:R và mọi phần khác của workbook giữ nguyên, ngoại trừ mở rộng phạm vi/định dạng cột S; file kiểm tra giữ nguyên hash.

Bố cục hiện hành: **A:Q** theo mẫu, **R – Hàng khuyến mại**, **S – Số chứng từ** theo yêu cầu. Vẫn **34.579 hóa đơn**, tổng thanh toán **53.250.290.891đ**, còn 3 chứng từ cần kiểm tra. Không tự ghép S vào Mã đơn hàng của meInvoice; đây là số chứng từ AMIS. Hàm cập nhật hẹp `add_document_numbers()` trong `scripts/match_import_2024_template.py` (`--add-document-number`); luồng xuất mới cũng tạo cột S để không mất khi tái xuất. Verifier đã kiểm thêm S phải khớp BHH trong bảng đối chiếu. Căn cứ `scripts/data/amis-2024-linked-20261026/document-column-verification.json`; cập nhật hash ở hai báo cáo kiểm chính. Không tạo file bàn giao mới, không nhập/ký/phát hành, chưa commit/push.

---

# Đã đưa hai đơn khuyến mại 0đ vào import — 27/10/2026

Người dùng xác nhận **BHHT2024013153** và **BHHT2024013026** là đơn khuyến mại; sau đó xác nhận riêng **dùng thuế suất 8%, giữ toàn bộ tiền 0đ**. Đã đưa cả hai vào import, mỗi đơn **KM300FULL / SL 1**, giá/tiền hàng/VAT/thanh toán 0đ. Giữ nguyên ngày chứng từ 14/08/2024 và 15/08/2024 và thông tin khách hàng nguồn; không suy ra thuế suất từ tiền 0, không mặc định cho chứng từ khác. Nguồn AMIS giữ nguyên cả ô Hàng khuyến mại đang trống; cờ khuyến mại là dữ liệu nhập bổ sung theo xác nhận người dùng.

**Hai file chính đã cập nhật tại chỗ**, không tạo bản khác. File import hiện **34.579 hóa đơn / 43.972 dòng**, tổng tiền hàng **49.305.824.619**, VAT **3.944.466.272**, thanh toán **53.250.290.891đ** (các khoản tiền không đổi). **Còn 3 chứng từ / 5 dòng nguồn cần kiểm tra**: `BHHS2024000629` / `BTL00023`, `BHHS2024005761` / `BTL00024`, `BTL00145`. 7 mục loại trừ chưa rõ vẫn giữ riêng.

**Bố cục import hiện tại:** A:Q giữ nguyên 17 cột mẫu; **R bổ sung Hàng khuyến mại** là cột nghiệp vụ cần nhập, điền **1 tại đúng hai dòng khuyến mại**, các dòng khác trống. Khi ghép cột trong meInvoice, ghép R vào đúng trường Hàng khuyến mại; không ghép vào mã đơn hàng hay trường khác. Các cột đối chiếu BHH/BTL vẫn nằm trong sheet `Doi chieu import` của workbook kiểm tra. Đã thông báo thay đổi cột này cho người dùng khi thực hiện. Căn cứ MISA: https://helpv4.meinvoice.vn/kb/nhap-khau-hoa-don-mot-thue-suat/ (dòng khuyến mại thêm cột, giá trị 1); hướng dẫn MTT: https://helpv4.meinvoice.vn/kb/nhap-khau-hoa-don-khoi-tao-tu-may-tinh-tien-vao-phan-mem/.

Kiểm hoàn tất đầy đủ nguồn/loại trừ/từng dòng/tổng tiền, không lỗi ô Excel/công thức; mọi số tiền giữ đúng nguồn/bù trừ thực tế. Đã xác minh riêng **BHHT2024013153 tại dòng Excel 23747 / STT HĐ 18359**, **BHHT2024013026 tại dòng Excel 23995 / STT HĐ 18533**, mỗi dòng có **K=8, L=0, O=1, P=0, Q=0, R=1**. Đủ **10 xác nhận** trong `Tong hop` của file kiểm tra và dữ liệu chuẩn bị (1 số HĐ nhập nhầm + 7 lệch làm tròn + 2 khuyến mại). Nguồn giữ nguyên hash, hash file chính khớp bộ đã kiểm; thư mục xuất tạm đã dọn. Báo cáo `scripts/data/amis-2024-linked-20261026/verification.json`, `template-format-verification.json`. Chưa nhập/ký/phát hành, chưa commit/push thay đổi này.

---

# Đã đưa 7 chứng từ lệch 2–3đ vào import theo xác nhận — 27/10/2026

Người dùng xác nhận **giữ nguyên tiền hàng, thuế GTGT và tổng thanh toán theo AMIS**, đưa vào import riêng 7 chứng từ: `BHHT2024003185`, `BHHT2024011030`, `BHHT2024015896`, `BHHT2024017438`, `BHHN2024000179`, `BHHT2024023472`, `BHHT2024025057`. Giữ nguyên cả SL/đơn giá nguồn và thành tiền AMIS, không tính lại thành tiền từ SL × đơn giá. Không mở rộng dung sai cho chứng từ khác. Mỗi chứng từ có ghi xác nhận và phép tính lệch chi tiết tại `Tong hop` của file kiểm tra và `confirmations` trong bản dữ liệu chuẩn bị.

Đã cập nhật tại chỗ đúng **2 file chính** ở `outputs/amis-2024-theo-btl-bhh-20261026/`, giữ nguyên mẫu 17 cột A–Q của import và sheet `Doi chieu import` riêng trong workbook kiểm tra. Hiện **34.577 hóa đơn / 43.970 dòng**, tiền hàng **49.305.824.619**, VAT **3.944.466.272**, tổng thanh toán **53.250.290.891đ**. Tăng đúng 7 hóa đơn / 10 dòng, tiền hàng **43.857.407**, VAT **3.508.593**, thanh toán **47.366.000đ** so với bản trước. Còn **5 chứng từ / 7 dòng nguồn cần kiểm tra**: `BHHS2024000629` / `BTL00023`, `BHHS2024005761` / `BTL00024` (giá trả khác bán), `BHHT2024013153`, `BHHT2024013026` (tiền 0), `BTL00145` (thiếu BHH). **7 mục loại trừ chưa rõ** giữ nguyên.

Kiểm hoàn tất: đọc độc lập sổ AMIS, đủ nguồn/phân nhóm, mọi dòng import khớp nguồn hoặc bù trừ liên kết thật; riêng 7 chứng từ mới tiền hàng/VAT/thanh toán nguyên nguồn và cộng khớp; tổng 7 chứng từ 47.366.000đ. Các kiểm tra khác vẫn đạt, nguồn giữ nguyên hash, K chứa 8 định dạng số theo mẫu; chưa nhập/ký/phát hành. Đã kiểm 8 xác nhận hiện diện trong workbook (1 xác nhận số hóa đơn nhập nhầm trước đây + 7 xác nhận mới). Báo cáo `scripts/data/amis-2024-linked-20261026/verification.json`. Bỏ thư mục xuất tạm sau khi kiểm và thay vào đúng file chính; không sinh phiên bản bàn giao khác. Người dùng đã đóng file Excel theo yêu cầu rồi mới cập nhật. Chưa commit/push thay đổi này.

---

# Đã ghi số lệch cụ thể trong lý do kiểm tra — 27/10/2026

Theo yêu cầu người dùng, cập nhật tại chỗ `KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`, không tạo bản mới. Cột **K – Lý do** tại `Can kiem tra` ghi rõ dòng AMIS, SKU, SL × đơn giá, thành tiền AMIS và chênh có dấu (AMIS - ROUND(SL × đơn giá, 0)). 7 chứng từ: `BHHT2024003185` -2đ, `BHHT2024011030` -3đ, `BHHT2024015896` -2đ, `BHHT2024017438` -3đ, `BHHN2024000179` -2đ, `BHHT2024023472` -2đ, `BHHT2024025057` -2đ. Các chứng từ này đều AMIS thấp hơn phép nhân, chưa đổi cách phân loại hoặc số tiền.

Hai trường hợp trả lại ghi rõ giá bán, giá trả, SL trả, tiền theo giá bán và tiền trả thực tế: `BHHS2024000629` / `BTL00023` / KA150SFILTERFULL trả cao hơn **92.592đ**; `BHHS2024005761` / `BTL00024` / KM100FULL trả cao hơn **30.240đ**. Đồng bộ lý do tương ứng ở `Bu tru tra lai` và `Lien ket BTL-BHH`. Không tự bù trừ/đổi giá hai trường hợp này. Các ô quyết định và mọi dữ liệu khác giữ nguyên; vẫn 12 chứng từ cần kiểm tra, 7 mục loại trừ chưa rõ. File import không thay đổi.

Script sửa hẹp: `scripts/annotate_import_2024_differences.py`; hàm giải thích dùng lại tại `scripts/rebuild_import_2024_linked_returns.py`. Đã kiểm từng ô trên các sheet chỉnh sửa, chỉ các ô lý do được thay; mọi ZIP entry khác gồm toàn bộ sheet `Doi chieu import` giữ nguyên byte. Căn cứ `scripts/data/amis-2024-linked-20261026/differences-verification.json`. Không tự commit/push.

---

# File import đã chuyển đúng bố cục mẫu MTT — 27/10/2026

Theo yêu cầu người dùng, **cập nhật tại chỗ** hai workbook chính trong `outputs/amis-2024-theo-btl-bhh-20261026/`. File `IMPORT_HOA_DON_MOI_AMIS_2024.xlsx` hiện chỉ có sheet **Hóa đơn GTGT**, **17 cột A–Q**; giữ nguyên dòng hướng dẫn 1–7, dòng tiêu đề 9, font Times New Roman, màu chữ/đường viền, độ rộng cột từ mẫu `Mau_hoa_don_GTGT_MTT_1_thue_suat.xls`. Ngày thật giữ kiểu ngày Excel và hiển thị dd/mm/yyyy; tên hàng dài được ngắt dòng. Không còn sheet hướng dẫn riêng hay các cột phụ R:W trong file import.

Các cột phụ chuyển sang sheet **Doi chieu import** của `KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`: STT hóa đơn, dòng Excel import, BHH, dòng nguồn AMIS, mã hàng, BTL đã bù trừ và tổng tiền hàng/thanh toán. Giữ nguyên 7 sheet kiểm tra và các cột quyết định trước đó. **34.570 hóa đơn / 43.960 dòng**, tổng thanh toán **53.202.924.891đ**; dữ liệu A:Q không đổi so với bản đã kiểm hôm trước, 12 chứng từ và 7 mục loại trừ cần kiểm tra vẫn như cũ. Không tự ghép BHH vào Mã đơn hàng.

Script bảo toàn gói định dạng mẫu: `scripts/match_import_2024_template.py`, được gọi ở cuối `export_import_2024_linked_returns.py`. Khi dựng lại cả bộ, dựng review trước, xuất import sau để giữ sheet đối chiếu. Kiểm định dạng/dữ liệu không đổi: `scripts/data/amis-2024-linked-20261026/template-format-verification.json`; verifier hiện ghép dữ liệu đối chiếu từ sheet riêng khi đọc bản import 17 cột. Không sinh bản Excel bàn giao mới, không sửa nguồn hoặc thao tác AMIS/meInvoice. Chưa commit/push thay đổi định dạng này.

---

# Chỉ dùng một bộ 2 file, cập nhật tại chỗ — 26/10/2026

Đã gom bộ mới nhất về **`outputs/amis-2024-theo-btl-bhh-20261026/`**, chỉ còn `IMPORT_HOA_DON_MOI_AMIS_2024.xlsx` và `KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`. Đã bỏ các bản Excel cũ trùng tên và hai thư mục `outputs/amis-2024-moi-20261026/`, `cap-nhat-xac-nhan-ke-toan/` sau khi kiểm SHA256 xác nhận không có sửa đổi của người dùng. Hai file chính giữ nguyên hash của bộ cập nhật đã kiểm, dữ liệu không đổi: **34.570 hóa đơn / 43.960 dòng**, thanh toán **53.202.924.891đ**, **12 chứng từ cần kiểm tra**, **7 mục loại trừ chưa rõ**. File liên kết BTL–BHH nguồn và các bộ điều chỉnh khác giữ nguyên.

**Yêu cầu người dùng:** từ nay cập nhật trực tiếp hai file này, không tạo thêm bản/thư mục phiên bản. Nếu Excel đang mở làm khóa file, báo người dùng đóng file rồi mới cập nhật. Không tự sinh bản thay thế khác. Builder dùng thư mục chính mặc định và `--replace` khi tái xuất đã được yêu cầu. Các phần dưới là lịch sử, đường dẫn cũ có thể đã được dọn.

---

# Cập nhật theo xác nhận kế toán: BHHT2024004546 chưa có số hóa đơn — 26/10/2026

Người dùng xác nhận: kế toán nhập nhầm `BHHT2024004546` vào ô Số hóa đơn; thực tế chưa có số hóa đơn. **Đã bỏ riêng lý do giữ kiểm tra này, đưa BHHT2024004546 vào import**. Không áp dụng xác nhận cho chứng từ khác, không sửa ô nguồn AMIS.

Bộ cập nhật dùng tiếp ở **`outputs/amis-2024-theo-btl-bhh-20261026/cap-nhat-xac-nhan-ke-toan/`**, hai file `IMPORT_HOA_DON_MOI_AMIS_2024.xlsx` và `KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`. Lưu chung tại thư mục cập nhật vì workbook kiểm tra bản trước đang mở trong Excel. **Bộ này thay bộ ngay ở thư mục cha.** Hiện **34.570 hóa đơn / 43.960 dòng**, tiền hàng **49.261.967.212**, VAT **3.940.957.679**, tổng thanh toán **53.202.924.891 đồng**. Tăng đúng 1 hóa đơn / 1 dòng, tiền hàng 1.659.555, VAT 132.764, thanh toán 1.792.319 đồng. Còn **12 chứng từ / 17 dòng nguồn cần kiểm tra**, 7 mục loại trừ chưa rõ vẫn giữ nguyên. Các quy tắc ngày chứng từ, TM/CK, liên kết BTL–BHH thực tế, 158 đơn trả hết, 115 đơn loại trừ và 2 đơn bù trừ một phần giữ nguyên.

Xác nhận lưu tại `confirmations` trong `scripts/data/amis-2024-linked-20261026/prepared.json` và trong hướng dẫn/tổng hợp của hai workbook; verifier đã điều chỉnh kiểm đúng việc thêm duy nhất chứng từ này. Bản Excel trước xác nhận được giữ tại `scripts/data/amis-2024-linked-20261026/before-confirmation-BHHT2024004546.zip`. Không sửa nguồn, không nhập/ký/phát hành, không commit/push.

---

# Bộ import mới 2024 đã làm lại theo BTL → BHH thực tế — 26/10/2026

Theo yêu cầu mới nhất, bộ dùng để kiểm tra/nhập là **`outputs/amis-2024-theo-btl-bhh-20261026/`**. Bộ `outputs/amis-2024-moi-20261026/` bên dưới là lịch sử, **đã được thay thế, không dùng nhập**. Đã bỏ hoàn toàn việc giữ đơn ngoài import theo ứng viên trùng khách hàng/mã hàng/giá/ngày. Chỉ dùng BHH gốc thực tế ở `outputs/btl-bhh-20261026/LIEN_KET_BTL_BHH_2024.xlsx`.

- `IMPORT_HOA_DON_MOI_AMIS_2024.xlsx`: **34.569 hóa đơn / 43.959 dòng**, tiền hàng sau bù trừ **49.260.307.657**, VAT **3.940.824.915**, tổng thanh toán **53.201.132.572 đồng**. Giữ 17 cột A:Q của mẫu MTT người dùng đưa, tiêu đề dòng 9. R:W là cột đối chiếu BHH/tổng tiền/dòng nguồn/SKU/BTL. Ngày theo **Ngày chứng từ AMIS**, toàn bộ **TM/CK** theo xác nhận. STT lặp cho các dòng cùng hóa đơn; thông tin đầu hóa đơn và tổng VAT chỉ ghi dòng đầu. Email/điện thoại/CCCD trống theo nguồn. Thuế suất 8% được suy ra duy nhất từ tiền hàng/VAT với dung sai làm tròn 1đ/dòng; không tự cân tiền.
- `KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`: **13 chứng từ / 18 dòng nguồn cần kiểm tra**; **115 BHH / 170 dòng loại theo danh sách**; **158 BHH / 210 dòng đã trả hết**; đủ **219 dòng liên kết BTL–BHH**; bảng bù trừ từng dòng; **7 mục loại trừ chưa rõ** vẫn giữ để đối chiếu. Các sheet: `Tong hop`, `Can kiem tra`, `Dong nguon can kiem tra`, `Da loai va tra het`, `Bu tru tra lai`, `Lien ket BTL-BHH`, `Loai chua ro`.

**Cách bù trừ:** 158 BHH khớp hết SL, tiền hàng và VAT theo từng SKU/ĐVT với phiếu trả lại thì không lập hóa đơn mới. **2 BHH trả một phần đã trừ chính xác:** `BHHT2024006103` / `BTL00070` và `BHHT2024024887` / `BTL00180`. Ưu tiên khớp nguyên dòng bằng SKU/ĐVT/SL/tiền/VAT; chỉ trừ một phần khi xác định duy nhất dòng và phần còn lại khớp SL × giá nguồn, dung sai làm tròn tối đa 1đ. Đơn giá nguồn giữ nguyên; không suy đoán đơn giá mới.

**Các trường hợp tách riêng:** `BHHS2024000629` / `BTL00023` có giá trả SKU KA150SFILTERFULL 747.356đ, khác giá bán 654.764đ; `BHHS2024005761` / `BTL00024` có tiền trả KM100FULL 202.778đ vượt tiền bán 172.538đ. Hai BHH này giữ nguyên nguồn, chưa bù trừ, chưa vào import. `BTL00145` thiếu BHH thực tế nên giữ riêng, không loại/giữ lại bất kỳ đơn bán nào theo suy đoán. 10 BHH khác cần kiểm tra: 7 lệch SL × đơn giá hơn 1đ (`BHHT2024003185`, `BHHT2024011030`, `BHHT2024015896`, `BHHT2024017438`, `BHHN2024000179`, `BHHT2024023472`, `BHHT2024025057`), 2 chứng từ tiền 0 (`BHHT2024013153`, `BHHT2024013026`), 1 có số hóa đơn nguồn (`BHHT2024004546`).

**Kiểm hoàn tất:** đọc lại sổ thật và đối chiếu đủ 44.578 dòng với bản đọc trung gian; đối chiếu từng dòng import đúng ngày/BHH/SKU/giá/dòng nguồn; tổng SL/tiền hàng/VAT từng hóa đơn đúng bán gốc trừ trả thực tế; đủ 35.017 chứng từ phân nhóm không trùng/thiếu; tổng các nhóm nguồn khớp dòng tổng AMIS; không lỗi công thức/ô Excel, import chỉ có giá trị; SHA256 cả sổ AMIS, danh sách loại trừ, liên kết và mẫu giữ nguyên. Tổng ở bảng phân nhóm là **số nguồn trước bù trừ**, tổng import **sau bù trừ** hiển thị riêng bên dưới. Báo cáo `scripts/data/amis-2024-linked-20261026/verification.json`. Builder/phân tích: `scripts/rebuild_import_2024_linked_returns.py`, `scripts/build_import_2024_linked_returns.mjs`, `scripts/export_import_2024_linked_returns.py`, `scripts/verify_import_2024_linked_returns.py`. File kiểm tra do Artifact Tool dựng/render/xuất; import lớn dùng bundled XlsxWriter đóng gói dự phòng do Artifact Tool export lớn không hoàn tất trong phiên trước, đã kiểm độc lập lại bằng openpyxl chỉ đọc.

**Chưa nhập, ký, phát hành; không sửa chứng từ AMIS; không sửa nguồn và các bộ điều chỉnh cũ; không commit/push.**

---

# Đã lấy đủ liên kết BTL → BHH trực tiếp từ AMIS năm 2024 — 26/10/2026

Theo yêu cầu chỉ đọc của người dùng, đã truy cập AMIS Chrome Sang, xác nhận công ty **CÔNG TY CỔ PHẦN CÔNG NGHỆ UVGREEN**, dữ liệu **KETOAN_2020_PhucHoi_190120262046**. Lọc ngày hạch toán **01/01/2024–31/12/2024**, đọc đủ **163 phiếu / 219 dòng trả lại** tại cột **Số CT bán hàng** của bảng Chi tiết. **218 dòng có BHH, 162 BHH gốc khác nhau**. Chỉ **BTL00145 / SP000042 / dòng 1** trống BHH trên AMIS, đã đọc lại xác nhận và giữ nguyên trống. Phiếu này tiền hàng **972.222**, VAT **77.778**, tổng **1.050.000**.

File hoàn tất: `outputs/btl-bhh-20261026/LIEN_KET_BTL_BHH_2024.xlsx`, gồm `Chi tiet BTL-BHH`, `Phieu BTL`, `Thong tin`, `Chua co BHH`. Có BTL, BHH gốc, ngày, số dòng, mã/tên hàng, ĐVT, SL trả, đơn giá, tiền hàng/VAT/tổng tiền, khách hàng và số dòng trong sổ AMIS. Tổng hàng trả **225.865.926**, VAT **18.069.275**, tổng thanh toán **243.935.201 đồng**. Đủ 163 phiếu, số dòng và bộ mã hàng/SL/đơn giá/tiền/VAT của mọi phiếu đều khớp sổ AMIS đã xuất; không có hồ sơ khác sổ. Số tiền trong file liên kết là giá trị trả lại dương theo giao diện BTL; sổ bán hàng thể hiện VAT/thanh toán trả lại âm.

Căn cứ và kiểm: `scripts/data/btl-bhh-live-20261026/verified-live-reads.json`, `list-pages.json`, `prepared.json`, `verification.json`; ảnh xác nhận BTL00145 và BTL00071 cùng thư mục. **Không dùng `live-reads.json` ban đầu**: bản đọc nhanh bị trễ chi tiết khi chuyển phiếu; bản `verified-live-reads.json` đã đọc lại đủ, kiểm đúng phiếu được chọn, thay đổi nội dung sau chuyển phiếu, đủ số dòng, tổng chi tiết khớp tổng phiếu và ổn định qua hai lần đọc. Có cơ chế chuyển sang phiếu đối chứng khác tiền rồi đọc lại nếu nội dung không đổi. Script kiểm nguồn `scripts/prepare_btl_bhh_live_20261026.py`; builder Artifact Tool `scripts/build_btl_bhh_live_20261026.mjs`. File xuất đã đọc kiểm độc lập đủ 219 dòng, kiểm tổng từng phiếu và cache công thức; không có lỗi ô Excel, hash nguồn sổ giữ nguyên. Đã xem trước các sheet và vùng tiền.

**Không sửa chứng từ, không ghi/bỏ ghi sổ/xóa, không nhập/ký/phát hành, không commit/push.** Chỉ thay bộ lọc, số dòng/trang và chọn hàng để đọc. Chưa tái xuất bộ import hóa đơn mới theo các liên kết thật này; hai file import/review cũ vẫn chưa chốt do quy tắc dò rộng đã bị người dùng bác bỏ. Khi làm tiếp, dùng liên kết thật ở file này thay toàn bộ nhãn “Có thể liên quan trả lại”, không coi 15.879 ứng viên cũ là căn cứ loại đơn.

---

# Cần lấy liên kết BTL → BHH thực tế; bản import 26/10 chưa chốt — cập nhật 26/10/2026

Người dùng đã bác bỏ cách giữ ngoài import các đơn bán chỉ do trùng mã khách theo sàn/mã hàng/đơn giá/ngày với phiếu trả lại. **15.879 đơn gắn nhãn “Có thể liên quan trả lại” không được coi là đã trả hoặc căn cứ loại/giữ ngoài import.** File import 18.850 HĐ ở mục dưới thiếu phạm vi do điều kiện dò quá rộng; kiểm tra kỹ thuật/tổng tiền đã đạt không chứng minh phương án lọc này đúng nghiệp vụ. Chưa tái xuất hoặc thay thế hai workbook sau khi người dùng bác bỏ quy tắc. Không dùng bản đó để nhập cho tới khi hoàn tất sửa phạm vi.

Người dùng yêu cầu tìm và lấy **Số CT bán hàng** gốc BHH trên từng dòng chi tiết của từng phiếu BTL, không xuất lại Sổ chi tiết bán hàng. Ảnh người dùng xác nhận trực tiếp **BTL00071 → BHHS2024008957**, hàng **KK100FULL**, SL **1**, tiền hàng **962.037**, VAT **76.963**, tổng **1.039.000**. Trường này có trong bảng Chi tiết dưới danh sách phiếu và trong chứng từ; không có trong các cột danh sách BTL người dùng đã kiểm tra.

Tài liệu MISA xác nhận Alt+Y xuất chi tiết chứng từ nhưng **chưa xác nhận file chứa trường Số CT bán hàng**. Không tiếp tục trả lời rằng xuất Sổ chi tiết bán hàng là đủ. Hướng xác minh/lấy dữ liệu thực tế là đọc bảng chi tiết từng BTL qua phiên AMIS đang đăng nhập, ghi đủ số phiếu, thứ tự dòng, mã hàng, SL, BHH gốc, tiền hàng/VAT/thanh toán và nguồn; giữ trống nếu AMIS không có liên kết. Đối chiếu số phiếu/dòng với 163 BTL/219 dòng của sổ, không tự ghép theo tiền/tên.

Đã mở AMIS trong Chrome hồ sơ Sang qua công cụ trình duyệt; chuyển tới màn hình đăng nhập MISA. Cần người dùng đăng nhập đúng dữ liệu UVG để tiếp tục đọc trực tiếp. Không yêu cầu gửi mật khẩu/OTP/cookie/token. Chưa sửa chứng từ, chưa nhập/ký/phát hành, chưa commit/push.

---

# Đã dựng import hóa đơn mới từ sổ AMIS 2024 sau loại trừ — 26/10/2026

Người dùng yêu cầu tạo hóa đơn mới từ sổ AMIS vừa xuất, tách trường hợp chưa rõ để kiểm tra. Đã xác nhận **ngày lập lấy theo AMIS** (dùng cột Ngày chứng từ; giữ ngày hạch toán/ngày hóa đơn nguồn riêng ở bảng đối chiếu) và **TM/CK cho toàn bộ**. Mẫu người dùng gửi: `data/Du lieu Meinvoice/Mau_hoa_don_GTGT_MTT_1_thue_suat.xls`, thực chất OOXML, sheet `Hóa đơn GTGT`, 17 cột A:Q, tiêu đề dòng 9; đây là mẫu GTGT máy tính tiền, một thuế suất.

Nguồn mới giữ nguyên: `data/Du lieu Amis/So_chi_tiet_ban_hang_2024.xlsx` và `data/HD Xuất CTY 2024.xlsx` (Drive ID `1E6-s4ivoAuMBYMub5SzlEtpNWATNj2Tu`). Lọc BHH ở tab `HD cty 2024` và mã tách/gộp tại Hướng xử lý; **không loại toàn bộ tab F0/F1/F2 tham khảo**. Tab loại có 118 hóa đơn, nhưng sổ mới có 115 BHH được loại; mã 2023 không thuộc sổ 2024, HĐ 60/63 thiếu BHH. Các mã tách 0113/.1, 0116/.1/.2 và gộp 0079 về 0078 được xét. Sheet `Loai chua ro` giữ 7 mục gồm mã cũ/2023, hai HĐ thiếu BHH và 3 HĐ chỉ có trong bảng ghi nhận (585,586,785).

- `outputs/amis-2024-moi-20261026/IMPORT_HOA_DON_MOI_AMIS_2024.xlsx`: **18.850 hóa đơn / 24.677 dòng**, tiền hàng **27.244.438.581**, VAT **2.179.555.670**, thanh toán **29.423.994.251 đồng**. Giữ đúng 17 cột mẫu; R:V bổ sung BHH, tổng tiền hàng/thanh toán, dòng nguồn và mã hàng. Ngày và thông tin đầu hóa đơn/VAT chỉ ở dòng đầu, STT lặp cho các mặt hàng cùng hóa đơn. Email/điện thoại/CCCD trống do nguồn không có. Thuế suất 8% được suy ra duy nhất từ tiền hàng/VAT của từng dòng, sai số làm tròn tối đa 1đ; không đổi số tiền nguồn.
- `outputs/amis-2024-moi-20261026/KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx`: **16.052 chứng từ / 19.731 dòng cần kiểm tra**, **115 chứng từ / 170 dòng đã loại**, bảng đối chiếu import và phiếu trả lại. Cột vàng để người dùng bổ sung quyết định/BHH đúng. Có **163 phiếu trả lại / 219 dòng** không ghi chứng từ bán gốc; dò cùng mã KH/mã hàng/đơn giá ±1đ, bán trước hoặc cùng ngày trả và đủ SL tạo ra **15.879 đơn bán có thể liên quan**, đều giữ ngoài import. Đây là ứng viên thận trọng do khách hàng sàn dùng chung mã/tên và đơn giá; **không kết luận 15.879 đơn đã trả** và không tự chọn để bù trừ. Những trường hợp khác gồm tiền âm/0, lệch SL×đơn giá và thông tin số HĐ có sẵn.

Ba nhóm cộng đúng **35.017 chứng từ / 44.578 dòng**, doanh số **50.154.037.726**, trả lại **225.865.926**, VAT **3.994.254.050**, thanh toán **53.922.425.850 đồng**. Kiểm độc lập từng dòng import và đủ 42 cột nguồn của nhóm kiểm tra, phân nhóm không trùng/thiếu dòng, header đúng mẫu, không có BHH loại/ứng viên trả lại trong import, không có lỗi ô Excel hoặc công thức trong import, hash hai nguồn giữ nguyên. Báo cáo `scripts/data/new-amis-2024-20261026/verification.json`; chuẩn bị `scripts/prepare_new_amis_2024_20261026.py`, dựng/preview Artifact Tool `scripts/build_new_amis_2024_20261026.mjs`, đóng gói dự phòng `scripts/export_new_amis_2024_20261026.py`, kiểm `scripts/verify_new_amis_2024_20261026.py`. Artifact Tool dựng/render được nhưng xuất lớn không hoàn tất sau chờ kéo dài, bộ nhớ tăng ~3.7GB; đã dừng riêng hai tiến trình tác vụ, xuất cuối bằng bundled XlsxWriter rồi đọc độc lập bằng openpyxl chỉ đọc. Không sửa các bộ import điều chỉnh cũ.

**Chưa nhập, ký, phát hành; chưa commit/push.** Đây là bản AMIS sau lọc theo danh sách người dùng, chưa chứng minh mọi hóa đơn lịch sử trên meInvoice đã được xử lý hoặc được phép phát hành lại. Khi tiếp tục, dùng hai file trên; không tự nới nhóm trả lại hay đưa hồ sơ chưa rõ vào import.

---

# Đã tạo bảng F0/F1/F2 cho gốc 01/07–31/12/2025 — 22/10/2026

Rà riêng536 dòngFghi lệchtên vớiNB22/10: đã đọc lại trực tiếp37.087MEINV B:N layout33cột và26.550NB C:T; bỏ điều kiện tên chỉ cho phép so sánh tài chính, cùngsốHĐ/cả3 khoản trêncùng1dòngNB sau làmtròn nguyênđồng. **24 khớp cả3tiền chính xác;3 lệch tốiđa1đ;0 chỉlệch2đ;509 lệchquá2đ**. Nhóm27khớp/gầnkhớp có25mãNBduynhất và2nhiềumã. Chỉ phân tích, CHƯAđiền mã hay sửa cộtnguyênnhân/sốtiền vì tênvẫnkhác. Ví dụ1C25MUV/00001579 tênMEINVcôngtyGiaKhánhTâm, NBnhutrang833, cả3tiền509.204/40.736/549.940khớp. Chi tiết `scripts/data/nb-fill-20261022/name-mismatch-money-audit.json` (đủ536dòng,cácứngviênNBvàchênhtiền). NB khôngcókýhiệu, khớpsố+tiềnkhôngtự chứngminh cùngHĐthật.

**Layout Google mới nhất sau thêm cột nguyên nhân22/10:**33cột A:AG. D=mãđơnF0; E=Mãđơntrùng; **F=Nguyên nhân chưa ghép mã đơn**, tôvàng/ghi lýdo1.210dòng (659khôngsốNB,536lệchtên,10tiềnlệchquá2đ,5nhiềumã). Đã đọc lại trướcghi xácnhận1.210D vẫntrống/khóa+tên+tiềnkhớpaudit; chèn1cộtF, ghichỉcộtF vàcopyđịnhdạngtừcộtghichú; giữmọi dữliệu cũ. Kiểm toànbộ37.087dòng B:N sauchèn: tấtcả12giátrịcũ B:M khớp saudịch, đúng1.210nhãn/1.210ôvàng, bộlọc/filterview mởrộngAG, headerđúng. Nguồn/plan/verif/ảnh: `scripts/data/nb-fill-20261022/unmatched-cause-marker-before.json`, `unmatched-cause-marker-verification.json`, `unmatched-cause-marker-after.png`. Kýhiệu/sốF0 nay I/J, ngàyK, tiềnhàng/thuế/thanh toánF0 L/M/N; ba tiềnhiệntại AC/AD/AE, tìnhtrạngchuỗiAF, ghichúAG. Đọc lạiheader trước tácvụsau vì cácmốc cộtởcácđoạnlịchsử dướiđãdịch. Chỉ sửaGoogle; ExcelcụcbộvẫnchưađồngbộcácmãNB/cộtE/F.

Phân loại1.210 ômãđơn còntrống sau2đợtNB (đọc lại trực tiếp22/10 cả37.087 dòngMEINV B:M theo layoutmới và26.550 dòngNB C:T):659 khôngcùngsốHĐtrongNB;536 cósốHĐnhưng tênngườimua khôngkhớp;10 cùngsố+tênnhưng bộtiền lệchquá2đ;5 khớpđủra nhiềumãNB. Tổng659+536+10+5=1.210, khôngcó trườnghợpmới đủđiềukiện chờđiền. NBkhôngcókýhiệu, nênnhóm536 chưađượcgọi là ghi sai tên hay cùngHĐthật. Ví dụ2853(1C25MUV/00002853): meInvoice679.000/NB1.358.000;2858 meInvoice1.039.000/NB2.078.000. Chỉphân tích, khôngsửaGoogle. Chi tiết `scripts/data/nb-fill-20261022/unmatched-reasons-live.json`.

**Cột đánh dấu trùng mới nhất nằm ở E (ngay sau D)** theo yêu cầu22/10: tabGoogle `MEINV F0, F1, F2, ...` hiện32cột A:AF; D vẫn `Mã đơn hàng F0`, E mới `Mã đơn trùng`, giá trị `Trùng mã (2 HĐ)`/3/4 ở229dòng thuộc112mã, ôkhácđểtrống; conditionalformat Ekháctrống tôvàng. Đã mở rộngbasicFilter vàfilterview sangAF, giữthôngtin lọc. Banđầu thêmAF rồi dùngmoveDimension chuyểnAFsangE theo yêu cầu tiếp theo; **các cột E:AE cũ dịch phải1 cột**, cần đọc lạiheader trước mọi tác vụ sau. Ví dụ kýhiệu/sốF0 nay H/I, ngàyJ, tiềnhàng/thuế/thanh toánF0 nay K/L/M; tìnhtrạngchuỗiAE, ghichúAF. Đã kiểm đủ37.087 dòng: mãD và cácgiátrị cũ giữnguyên, ngày/3khoản tiền F0 vàblankthanhtoán khôngđổi; đúng229marker/vàng, headerđúngthứtự, ruleđãđổi=$E2<>"". Báo cáo `scripts/data/nb-fill-20261022/duplicate-marker-before.json`, `duplicate-marker-verification.json`. Chỉ đánh dấu trùngmãghi chú, không kếtluận xuấttrùngHĐ. Googlelàbảnhiệnhành, Excelcụcbộchưađượcthêmcột này.

Rà trùng sau bổ sungNB22/10 (đọc trực tiếp D:H đủ37.087 dòng):112 mãđơn xuất hiện ở nhiềuF0, gồm229 hóađơn gốc/117 lượt lặp thêm, tốiđa4HĐ/mã. Không có trùng khóa kýhiệu+số hóađơn. Đây là trùng mãghi chú, chưa kết luận xuấttrùng; không sửa/tô thêmGoogle trong lượt chỉhỏi này. Chi tiết `scripts/data/nb-fill-20261022/duplicate-orders-after-fill.json`. Ví dụ580771131904198000 gắn1C25MUV17754/17811/17994/18305;581284749358761903 gắn23087–23090;HD010043 gắn28831/28832/34385. Cần đối chiếu nguồn nghiệp vụ trước quyết định xửlý.

**Đã xử lý61 trườnghợp lệch1–2đ theo yêu cầu mới nhất22/10:** người dùng cho phép bổ sung mãNB, chỉ ôtrống/tôvàng/giữtiền. Đã đọc lại61 dòng đích B:AE và61 dòng nguồnNB, xác nhậnkhóa+tên+3khoản tiền/mãkhôngđổi, ôD vẫn thậttrống vàkhôngvalidation/côngthức; ghi61 ôD bằngchuỗi và tô #FFF2CC trong1batch, khôngghi tiền. Đọc lại đủ61 xácnhận mã/màuđúng, tấtcảuserEnteredValue/effectiveValue của29ôcòn lại trênmỗi dòng giữnguyên, khôngbỏqua trườnghợp nào. Báo cáo `scripts/data/nb-fill-20261022/rounding-write-before.json`, `rounding-write-verification.json`; ảnhGoogle thực tế `rounding-write-after.png`. Tổng bổsungNB qua2đợt18.698 mã; số còntrống suy ra1.210. Quyết định này thay thế ghi chú CHƯAđiền61 ở dưới. Chỉ GoogleSheetđược sửa; Excelcụcbộvẫnchưacậpnhật.

Rà lệch1–2đ sau bổ sung mãNB (chỉ đọc,22/10): trong1.271 ôD còntrống có61 trườnghợp cùngsốHĐ+tên, cả3 khoản tiền sau làm tròn nguyênđồng chênh tốiđa2đ;60 trườnghợp tốiđa1đ và1 trườnghợp2đ (`1C25MUV/00005487`, dòng5487, NB6617, mã2508088CFM4RR0; tiền hàng meInvoice1.257.409/NB1.257.407, thuế100.593 bằngnhau, thanh toán1.358.002/NB1.358.000). Cả61 có1mãNB duy nhất; đã đọc lại trực tiếp61 dòng đích và61 dòngNB xác nhận ôD vẫntrống/mãvàtiền khôngđổi. **CHƯA điền thêm61 mã hoặc sửa số tiền**; người dùng hiện chỉ hỏi có lệchkhông, chưa yêu cầu mở dung sai. Chi tiết `scripts/data/nb-fill-20261022/rounding-audit.json`, `rounding-live-verification.json`, script `scripts/audit_nb_rounding_20261022.py`.

**Bản làm việc mới nhất là Google Sheets** `HD 2025`, tab `MEINV F0, F1, F2, ...` (gid2070647506), URL `https://docs.google.com/spreadsheets/d/1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw/edit?gid=2070647506#gid=2070647506`. Theo yêu cầu trực tiếp22/10 đã chỉ điền18.637 ôD đang thật sự trống và tô vàng #FFF2CC, từ sheet `Đối soát NB`. Khớp sốHĐ + tênNFC/hoa-thường/khoảngtrắng + cả3 khoản tiền làm tròn vềđồng rồi bằngnhau chính xác; chỉ1mãNB duy nhất và1HĐđích duy nhất. Không dùng ngày làm điều kiện theo yêu cầu lượt này. Giữ nguyên17.179 mã cũ/côngthức và sốtiền. Trước mỗi đợt ghi đã kiểm lại ôtrống/khóa/tiền; không có thayđổi đồng thời. Đã đọc lại toàn bộ37.087 dòng B:L:18.637 mã và18.637 màu đúng, không đổi ôngoàiphạmvi. Sau bổsung có35.816 mã; còn1.271 trống gồm1.266 chưa khớp đủ và5 khớp nhiều mãNB. Không tự chọn5mã này. Nguồn/plan/nhật ký/verification/ảnh saughi ở `scripts/data/nb-fill-20261022/`. Chỉ ghiGoogleSheet, **fileExcel cục bộ chưa cập nhật mãNB**; không upload lạiExcel cũ đèGoogle đã chỉnh.

Rà mã đơn theo thư mục người dùng chỉ định `data/Du lieu Meinvoice/T7_T12_2025_HD_da_su_dung` (22/10): đọc8xlsx,37.768 khóa riêng biệt; không đọc chồng filezip. File F đã có cộtD `Mã đơn hàng F0` và cộtL sheetnguồn `Mã đơn hàng`. Kiểm toàn bộ37.087 dòng gốc trong workbook:17.179 có mã,19.908 trống; 8file mới không bổ sung mã khác/không trống so với dữ liệu đang có, không có xung đột ngày/tiền/mã. 251gốc không có trong bộ8file; không tự bỏ hoặc đoánmã. Workbook KHÔNG bị sửa/không tạo cột trùng. Audit `scripts/audit_h2_order_codes_20261022.py`, `scripts/data/h2-order-codes-20261022/audit.json`. Mã đơn chỉ là ghi chú nguồn, không dùng để nối chuỗi hoặc xác nhận chắc chắn đơn thực.

File hiện hành: `outputs/HOA_DON_GOC_T7_T12_2025_F0_F1_F2.xlsx`, 2 sheet `F0, F1, F2, ...` và `Chi tiet nguon`. 37.087 gốc / 37.758 hóa đơn nguồn; 662 chuỗi có liên quan, 671 hóa đơn liên quan. Giữ nguyên436 thanh toán trống ở các khối F và sheet nguồn; tô vàng và ghi F tại cộtF `Thanh toán trống tại F`, thuộc436 chuỗi. Tổng thanh toán hiện tại của436 chuỗi này cũng để trống. Tiền hiện tại cộng điều chỉnh, còn thay thế thì dùng hóa đơn thay thế sau cùng, không cộng lại hóa đơn đã bị thay thế. Không phải file import và chưa thêm kế hoạch điều chỉnh mới.

Đã kiểm toàn bộ dòng/các khối F, tiền nguồn và cache công thức với JSON ghép; giữ số0 đầu, ngày kiểuExcel, bộ lọc, frozen panes và conditional formatting. Báo cáo `scripts/data/h2-chain-readiness-20261022/workbook-verification.json`, SHA256 `e01960a0f2fb784b1f76bab5ce9dce82807d2c9ee41442fcd438cc5547152dca`. Đã xem ảnh main/source/blank trước xuất. Artifact Tool tạo/tính/render được nhưng xuất workbook lớn không hoàn tất sau chờ kéo dài; đã dừng đúng tiến trình xuất riêng, ghi cuối bằng bundled XlsxWriter với cùng bố cục/công thức/cache và kiểm độc lập toàn bộ bằng openpyxl chỉ đọc. Tái tạo: `scripts/audit_h2_chain_readiness_20261022.py` -> `scripts/build_h2_workbook_20261022.py` -> `scripts/verify_h2_workbook_20261022.py`; builder không tự ghi đè file đã có. Chưa commit/push. Các đoạn dưới mô tả lịch sử chuẩn bị và được cập nhật bởi phần này.

---

# Nguồn dựng F0/F1/F2 cho gốc 01/07–31/12/2025 — 22/10/2026

**Cập nhật sau ảnh UI:** đã nối đủ `1C25MUV/00031054` (19/12/2025) -> `1C25MUV/00034893` (31/12/2025, BBTT00109) -> `1C26MUV/00002639` (13/01/2026). Nguồn ảnh lưu `data/Bo sung/2026-10-22/`; báo cáo thay thế2026 dòng14 và bảng tổng2026 phần1 dòng2644 xác nhận2639 là Hóa đơn thay thế, HĐ hợp lệ, tiền hàng0/thuế0/thanh toán trống. Không tự điền0 thanh toán. Tái ghép còn37.087 gốc,662 chuỗi có liên quan,671 hóa đơn liên quan,37.758 hóa đơn trong chuỗi; không còn thiếu cạnh theo trạng thái. Đủ dựng bảng F theo nguồn hiện có, vẫn giữ436 thanh toán trống. Các nhận xét thiếu31054 và số37.088 gốc/670 liên quan trong đoạn cũ dưới đây đã được thay thế bởi cập nhật này. Chưa tạo workbook mới/chưa sửa import.

Đã lưu nguyên bản hai báo cáo liên kết 2026 mới vào `data/Du lieu Meinvoice/Xuat_lai_2026-10-22/`, kiểm SHA-256 khớp Downloads. Đã ghép với bảng tổng và liên kết 2025: 37.088 gốc, 662 gốc có liên quan, 670 hóa đơn liên quan (610 năm 2025, 60 năm 2026). Chưa chốt đủ: `1C25MUV/00031054` ngày19/12/2025 ghi đã bị thay thế nhưng chưa có liên kết tới hóa đơn thay thế; cần mở lịch sử trên meInvoice. 436 hóa đơn có tiền hàng/thuế0 nhưng thanh toán trống, giữ nguyên nguồn. Excel H2 có38.019 hóa đơn, nhiều hơn snapshot Google251 gốc, không tự bỏ phần này. Báo cáo `scripts/data/h2-chain-readiness-20261022/KET_QUA_KIEM_TRA.md`, `audit.json`, script `scripts/audit_h2_chain_readiness_20261022.py`. Chưa tạo workbook F mới hoặc sửa import. Đủ để dựng bản có đánh dấu thiếu, chưa xác nhận đầy đủ mọi F/tổng thanh toán.

Chốt diễn giải mới theo người dùng22/10: chỉ thay dòng diễn giải không phải hàng hóa trong cột Tên hàng hóa/dịch vụ; luôn tham chiếu hóa đơn gốc F0, kể cả dòng đảo khoản cũ. Đã cung cấp prompt cho máy khác, CHƯA thực hiện sửa file ở lượt này. Giữ nguyên tên hàng hóa/khuyến mãi, số liệu và cột Lý do điều chỉnh.

---

# Đã rà mã đơn ghi chú theo yêu cầu người dùng — 21/10/2026

Người dùng đính chính: mã đơn trên meInvoice và NB do nội bộ note, có thể sai. Đã đọc toàn bộ Google Sheets:65.495AMIS/37.768meInvoice/26.550NB/65.436ánhxạMISA vàcác tabreview. Không dùng mãnote độc lập; kiểm số/ngàyHĐ,tên,bộ tiềnNB,rồi AMISthuế/thanh toán/chứngtừ,kiểmtrùng. Chuẩn hóa ngày chữ05-Jul-2025, không tự đoán38dòngNBghi2709/2025và577ngàytrống. 23.938khớp trường đốisoát3nguồn;276HĐmới cùngmãAMIS nhưng lệchthuế/tổng;629mãcónhiềuHĐ, trongđó48nhómđãcó96dòngKTnote đượcgiữnguyên. 1.619cùngsốNBchưakhớpngày,114lệchtên,12lệchtiềnsau làmtròn,5nhiềumãNB. Không gọi mọi khácbiệt là mãsai vàkhông tựloại HĐđiềuchỉnh/thaythế. Ví dụ11305dòngmeInvoice1237,tổng1.379.000,mã250707F76E0BV1trỏAMISBHHS2025010663dòng34508tổng2.479.000;NB377cũng ghi mãnày, nênhai mãnote khớpnhau chưa chứng minhđúngliênkết.

Nhóm trướcT7/sauT7: **116ứngviên khớp trường3nguồn +1chưatìmAMIS (1C25TUV/00010101,NB2395)**,tổng117. Số sơbộ114thành116sau đọcđúngngàytextNB231–232cho11052/11053; không sửaSheet. Cả116làHĐmới,AMISvàngàynhậnNBtrước01/07,ngàyHĐtừ01/07; mãkhônglặpởHĐkháctrongbảng07–12. 35ngườimua chetên mộtphần; chỉgọiứngviênkhớpđốisoát,chưachốtđúngđơnthực hoặcđủmọiđơn2024. Chưađưa thêm vàoimport. Chi tiết `scripts/data/link-audit-20261021/KET_QUA_KIEM_TRA.md`, `summary.json`, `audit.json`cục bộ; script `scripts/audit_order_links_20261021.py`. Snapshotnguồn Google local trongsource/,khôngtựđưalênGit; táikiểm máy khác cầnđọclạitheomanifest. KhôngsửaGoogleSheets/cácfileimport/ký/pháthành.

---

# Nguồn Google Sheets để rà xuất bù sau 01/07/2025 — 21/10/2026

Người dùng cung cấp file `HD 2025`: https://docs.google.com/spreadsheets/d/1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw/edit . Đã đọc trực tiếp qua Google Drive/Sheets, chỉ xem, không sửa file. Theo người dùng: `AMIS` là dữ liệu xuất AMIS; `Meinvoice 07 -> 12 / 2025` là bảng kê hóa đơn01/07–31/12/2025; `Đối soát NB` là dữ liệu note nội bộ. SheetID tương ứng1264954379/1180894322/13815067. Có thêm `MISA_HD_GOC`(1882710702) ánh xạ số chứng từ/mã đơn và các tab kiểm tra trùng/lệch.

Khác hai bản AMIS cũ trong Git, AMIS Google hiện có ngày hạch toán(A), số chứng từ(C), thuế(G), thanh toán(H), mã đơn(I), số HĐ meInvoice(J), thanh toán meInvoice(K). MeInvoice có ký hiệu(B), số(C), ngày(D), mã đơn(F), người mua(L), tiền hàng/thuế/thanh toán(P:R), trạng thái(U), số chứng từ AMIS(X), cờ bổ sungNB(Y), ghi chú(Z). NB có ngàyHĐ(A), sốHĐ(C), người mua(H), mãđơn(I), ngày đặt/nhận/thanh toán(J:L), tiền hàng/thuế/tổng(R:T); NB không có cột ký hiệu HĐ. Các cột liên kết AMIS/meInvoice có công thức MATCH/INDEX và một số #N/A; không coi liên kết trả về dòng đầu là xác nhận duy nhất nếu chưa kiểm trùng. Không ghép chỉ bằng số HĐ, mã đơn, tên hoặc số tiền.

Ví dụ đã đọc khớp3nguồn: `1C25TUV/00010299`, ngày01/07/2025, người muaelectronmastervvvl, tiền hàng2.573.148,thuế205.852,tổng2.779.000. NB dòng2: đặt26/06/2025 13:50, nhận26/06/2025 16:06. AMIS dòng36420: ngày hạch toán26/06/2025, chứng từBHHS2025010042, thuế205.852,tổng2.779.000, sốHĐ00010299. MeInvoice dòng231: cùng số/ngày/tên/bộ tiền và sốchứngtừX=BHHS2025010042. Đây là ví dụ phù hợp nhóm đơn trước tháng7/hóađơn sau đó; chưa rà toàn bộ để chốt tổng số. Dùng ngày nhận hàng làm nhóm riêng với ngày đặt/ngàyhạchtoán, không tự đồng nhất. Cần đánh dấu đơn đặt trước nhưng nhận sau01/07riêng; xét hóađơngộp,hoàn/trả,đã hủy/điềuchỉnh và quy tắcngoạilệ trước khi mở phạm viimport. Hai fileimport hiện tại chưa thay đổi.

---

# Đã bổ sung 13 hóa đơn vào bộ import theo phương án 19/10/2026

Người dùng đã yêu cầu thực hiện: “Ok, bổ sung 13 hóa đơn vào bộ import theo phương án mới đi, giữ đơn giá dương và kiểm tra tổng các hóa đơn liên quan về 0 nha.” Đã cập nhật các file chính, thay thế trạng thái chưa bổ sung của mục lịch sử bên dưới.

- Bộ Đã điều chỉnh: **2.045 hóa đơn /41.996 dòng /1.015 chuỗi**. Bổ sung13 hóa đơn /206 dòng, STT**2033–2045**, ngày dự thảo19/10/2026. Bản đầy đủ và `HOA_DON_DA_DIEU_CHINH_IMPORT_2_SHEET.xlsx` có XML cả2sheet chung và styles giống hệt nhau. Giữ nguyên2.032hóađơn đã chuẩn bị trước.
- Cả bộ: **10.455 hóa đơn /55.564 dòng** =8.410chưađiềuchỉnh +2.045đãđiềuchỉnh. Hai workbook Chưa điều chỉnh giữ nguyên byte.
- Đã chuyển8chuỗi ra khỏi Ngoại lệ, còn **236 chuỗi /322 hóa đơn**. Hai gốc478 và40 được thêm vào bảng Đã điều chỉnh; 4hóađơn nguồn liên quan được bổ sung nguyên trạng vào Chi tiet nguon, thành2.117hóađơn nguồn. MST8823260350 và8743683287 giữ nguyên giá trị nguồn, được xác nhận là cá nhân trong phạm vi hai chuỗi này.
- Danh sách72hóađơn khôngcầnđiềuchỉnh vẫn giữ nguyên72dòng và lýdo; sheet tổnghợp đã ghi “đã bổ sung13” và tổng hiện tại10.455.

Đã kiểm độc lập từng dòng mới, số/ngày thamchiếu, 8chữsố, trường thanh toán TM/CK, giá có sẵn dương và số lượng gốc giảm đúng. **Cả8chuỗi về(0,0,0)** theo tiền hàng/thuế/thanh toán, cả phép cộng nguồn +13HĐmới và công thức dự kiến lưu trong file. Mọi trường nguồn ngoài phạm vi, công thức, styles, bảng lọc, frozen panes và tài nguyên gốc được giữ nguyên; hơn1,8triệuô ngoài phạm vi đã đối chiếu trên hai bản Đã điều chỉnh. 0lỗi côngthức; đã xem trước/sau các vùng thay đổi. Các file chính có SHA-256 khớp bản đã kiểm độc lập sau khi thay file.

Các điểm cần giữ đúng: 5HĐđiềuchỉnh cũ508/2140/2189/2578/662 chỉ có khoản tiền, nên phần đảo không tự thêm số lượng/đơn giá. Gốc574 chỉ giảm số lượng, ô đơn giá/thànhtiền dòng vẫn trống; thanh toán nguồn trống giữ nguyên, tính dự thảo0 theo chốt cũ. Giữ6108/4203/4204 chỉđổi thôngtin ngườimua và9186 chỉđổi tênhàng; không lập đảo tiền các HĐnày. Giá âm theo các quyết định cũ trong2.032HĐtrước không thuộc phạm vi sửa13HĐlần này.

Có **2dòng của gốc498** giữ thành tiền đã phát hành khác SL×đơn giá làmtròn1đồng: STT2035, dòng import**41865 và41915** (dòng chi tiết58 và108 của gốc). Đã giữ đơn giá dương và thành tiền nguồn đảo dấu để tổngchuỗi khớp0; cần đối chiếu2dòng này khi xem bản nháp, tránh phần mềm tự tính lại từ đơn giá làmtròn. Căn cứ: `scripts/data/apply-13-20261019/rounding-review.json`.

Script, kế hoạch và bản trước sửa: `scripts/data/apply-13-20261019/`. Báo cáo hiện tại: `final-verification.json`. Kiểm hiện tại bằng `python scripts/data/apply-13-20261019/verify.py --final`; chỉ đọc Excel và ghi JSON. Không chạy lại prepare/build/merge/promote chỉ để kiểm tra. Chưa nhập phần mềm, ký/phát hành hoặc commit/push đợt này. Nguồn gốc không thay đổi.

---

# Đã tính lại nhóm 85: giữ 72, bổ sung 13 — 19/10/2026

Người dùng thay quyết định giữ 6 cặp riêng và xác nhận 4 hóa đơn 1C24TUV/00000478, 1C25TUV/00002578, 1C25TUV/00000040, 1C25TUV/00000662 dùng MST cá nhân, cần xử lý. Sau đó làm rõ: điều chỉnh giảm hóa đơn gốc bằng giảm số lượng, giữ đơn giá dương; hóa đơn mới phải trái dấu với hóa đơn được điều chỉnh và kiểm tổng các hóa đơn liên quan bằng 0. Không lập đảo tiền cho hóa đơn chỉ sửa thông tin mà không đổi tiền hàng, tiền thuế; chỉ xử lý gốc tương ứng.

Trong 85 hóa đơn: giữ **72 = 62 thuộc 31 chuỗi đã khớp + 6 thuộc 3 chuỗi cùng sản phẩm đã khớp + 4 chỉ đổi thông tin**. Ba HĐ 00006108/00004203/00004204 chỉ đổi thông tin người mua; 00009186 chỉ đổi tên hàng. Cả4 có tiền hàng/thuế/thanh toán XML=0; không gọi 9186 là đổi người mua. Hóa đơn gốc 9072 đã có trong import hiện tại; ba gốc1348/2377/2378 cần bổ sung. Chín HĐ còn lại thuộc6cặp cũ +4HĐ MSTcá nhân = **13 HĐ bổ sung trên8chuỗi**. Tổng theo phương án mới **10.455 =10.442+13 =10.527−72**.

Theo yêu cầu hiện tại đã xuất riêng `outputs/doi-chieu-85-20261019/UVG_72_HOA_DON_KHONG_CAN_DIEU_CHINH_20261019.xlsx`: 2sheet tổng hợp và danh sách72, có ký hiệu/số/ngày/gốc/người mua/lý do ngắn. Đã kiểm độc lập toàn bộ72 dòng, số0đầu, ngày, công thức, tổng phân nhóm; đọc lại nguồn để kiểm34chuỗi giữ nguyên đều cộng tiền hàng/thuế/thanh toán bằng0; kiểm4XML chỉđổi thôngtin và xem trước cả2sheet. 0lỗi côngthức. Kết quả: `scripts/data/decision-85-20261019/verification.json`; JSONcùngthưmục có đầyđủ13khóa cần bổ sung.

**Chưa cập nhật bộ import chính/bản2sheet/Ngoại lệ theo quyết định mới; các file import hiện vẫn có10.442HĐ.** Chưa áp dụng lại quy tắc dấu/đơn giá mới vào các khoản đã chuẩn bị trước. Không chạy builder lịch sử hoặc nhầm phương án10.455 với dữ liệu import đã cập nhật. Nguồn và5workbook hiện tại đã kiểm hash giữ nguyên. Không nhập/ký/phát hành; chưa commit/push đợt tính lại này. Danh sách85 và các quyết định giữ6cặp bên dưới là lịch sử trước thay đổi này.

---

# Danh sách 85 hóa đơn giữ ngoài import — 19/10/2026

Đã đối chiếu file kế toán `KE HOACH XUAT LAI HD 2024-T6.2025.xls`, sheet `BK HĐ 24-25 (BSUNG) (2)`, và xuất `outputs/doi-chieu-85-20261019/DANH_SACH_85_HOA_DON_KHONG_DUA_VAO_IMPORT.xlsx`. File gồm `Tong hop` và `85 hoa don`; mỗi hóa đơn có nhóm, lý do giữ ngoài import, gốc của chuỗi, ngày, người mua, thanh toán, vị trí trong sheet KT và căn cứ phân nhóm. Đã kiểm độc lập đủ 85 khóa riêng biệt, toàn bộ 1.020 ô dữ liệu khớp bản trích, tổng nhóm 62/12/4/6/1 và không có ô lỗi Excel; đã xem trước bố cục. Không sửa file KT hoặc các workbook import.

Tập KT đang lọc cột AL trống có 10.207 dòng; cộng AM được 11.274 lượt. Có 747 lượt trùng (739 F1 + 8 F2), còn 10.527 hóa đơn riêng biệt. KT đang trừ 739 + 9 thành 10.526 tại `Tong hop`/`Tổng hợp` B49: dư một lượt trừ so với tập đang lọc. Toàn bộ 10.442 hóa đơn thuộc bộ xử lý hiện tại có trong tập KT; 85 khóa còn lại gồm 62 hóa đơn của 31 chuỗi đã khớp, 12 hóa đơn của 6 cặp giữ riêng, 4 hóa đơn thuộc 2 chuỗi ngoại lệ công ty, 6 hóa đơn của 3 chuỗi cùng sản phẩm giữ nguyên và điều chỉnh tên 1C25TUV/00009186.

Thông tin mới thay thế phần chưa có định danh ở mục lịch sử bên dưới: KT ghi “HĐ không hợp lệ, hủy nội bộ” cho 1C25TUV/00001943 và 1C25TUV/00002060 tại dòng 2734 và 2851. Cả hai vẫn có AL trống và thuộc bộ xử lý hiện tại; không thuộc danh sách 85. Chưa có quyết định thay đổi import theo ghi chú này. Không tự loại chúng hoặc tự diễn giải các khoản 4 + 2 điều chỉnh tên/địa chỉ trong bảng tổng KT.

---

# Rà chênh số lượng theo ảnh kế toán — 19/10/2026

Theo yêu cầu đã xuất danh sách287hóađơn tại `outputs/DANH_SACH_287_HD_DIEU_CHINH_GOC_2022_2023.xlsx`, sheet `287 HD ngoai pham vi`: 37điều chỉnh cho gốc2022 và250cho gốc2023; có số/ngày hóa đơn điều chỉnh và gốc, tên người mua, file/dòng bảng tổng và liên kết để đối chiếu. Kiểm đủ287khóa riêng biệt, số8chữsố/ngày/gốc khớpJSONnguồn, xem ảnh đã giữ số0đầu. Không sửa import. Builder `scripts/layout-tools/list-287-20261019.mjs`.

KT gửi ảnh tính10.855−2hủy nội bộ+321liên quan−10hủy−6đổi tên/địa chỉ−325côngty−33thaythế=10.800, soimport10.442lệch358. Đã kiểm nguồn theo ký hiệu/số, không sửa import:10.855trongkỳ+321liênquan=11.176; trừ287điều chỉnh cho gốc2022–2023,10hủy,338ngoạilệ,30thaythế ngoài ngoạilệ,62hóađơn của31chuỗiđãkhớp,6hóađơn của3chuỗicùngsảnphẩm,1đổitên9186=10.442. Các nhóm loại trừ không trùng nhau. 338ngoại lệ=325côngty/MSTtrongkỳ+1liênquansaukỳ(1C25TUV/00012764ngày15/07/2025)+12hóađơn của6cặp giữriêng. 33thaythế có3đã thuộc ngoạilệ(1C24TUV60,61;1C25TUV2605), nên chỉ trừ thêm30. Chưa có danh sách2hủy nội bộ và6đổi tên/địa chỉ KT đang trừ; cần định danh để khớp từng khoản, không tự trừ thêm hoặc tuyên bố đã đối chiếu chính xác358theo danh sáchKT. Xem `scripts/data/count-difference-20261019/KET_QUA_DOI_CHIEU.md`, JSONchi tiết cùng thư mục và `scripts/audit_count_difference_20261019.py`. Chưa nhập/ký/phát hành.

---

# Bản gọn 2 sheet để nhập và đối chiếu — 14/10/2026

Rà số học hai bản gọn: 8.410 gốc chưa điều chỉnh và 1.007 chuỗi đã điều chỉnh đều dự kiến về 0 theo ba khoản sau cộng import; gốc6/16 có thanh toán nguồn trống, giữ giả định dự thảo0 đã chốt. Sáu chuỗi CC01–CC06 được quyết định **giữ cả cặp ngoài đợt xử lý**, không có trong import và vẫn còn số dư; gồm3/508. Không nhầm đây là hồ sơ đang chờ kế toán.

Nguồn mới người dùng đặt tại `data/Du lieu Amis/Bang_ke_chi_tiet_hoa_don_da_su_dung_2024.xls` thực tế là bảng kê chi tiết hóa đơn đã sử dụng (586 hóa đơn1C24TUV/23.428 dòng,09/01–31/12/2024), chưa phải Sổ chi tiết bán hàng AMIS. Chưa dùng để thay nguồn AMIS hoặc tái tính import. Hai Sổ chi tiết bán hàng AMIS cũ đang bị xóa cục bộ và file mẫu điều chỉnh đang có sửa cục bộ ngoài phạm vi cập nhật lần này; không đưa các xóa/sửa này vào commit bàn giao.

Đã tạo thêm bản cho **bộ chưa điều chỉnh**: `outputs/HOA_DON_CHUA_DIEU_CHINH_IMPORT_2_SHEET.xlsx`, chỉ gồm `Import dieu chinh` và `F0` (đổi tên từ `Chua dieu chinh`). Đủ 8.410 hóa đơn / 13.568 dòng dữ liệu import; giữ cả các dòng trống cuối sheet như nguồn. Kiểm độc lập 460.488 ô, giá trị/định dạng/công thức giữ nguyên, bản gốc không thay đổi. Báo cáo `scripts/data/import-2-sheet-untreated-20261014/verification.json`; chạy script tách với `--untreated`. SHA-256: `d31f5f1c7f6e48b4c345bff14472053d1c6e5489977d91fbb35d10b0f1d7c96c`.

Theo yêu cầu tạo file chỉ gồm 2 sheet, đã tách từ bộ **đã điều chỉnh** hiện tại ra `outputs/HOA_DON_DA_DIEU_CHINH_IMPORT_2_SHEET.xlsx`: (1) `Import dieu chinh`, (2) `F0, F1, F2, ...` (sheet `Dieu chinh` đổi tên). Chứa 2.032 hóa đơn dự kiến / 41.790 dòng import; không gộp 8.410 hóa đơn của bộ chưa điều chỉnh. Bản gốc không thay đổi. Giữ nguyên XML hai sheet, dữ liệu, công thức, cache, định dạng, bảng và frozen panes; bỏ các sheet/tài nguyên không dùng, cập nhật tên và vị trí sheet/bộ lọc. Kiểm độc lập 1.060.871 ô đạt; không có công thức tham chiếu sheet đã bỏ. Báo cáo `scripts/data/import-2-sheet-20261014/verification.json`; script `scripts/layout-tools/extract-two-import-sheets-20261014.py`. SHA-256 bản gọn: `f5661f3f3717ee2a9c78446f185b231a6b0372f13174d4bda4a83c775c5f7ea9`. Chưa nhập/ký/phát hành.

---

# Cập nhật diễn giải import theo câu người dùng chốt — 14/10/2026

Đã sửa **chỉ cột N — Lý do điều chỉnh** trong sheet `Import dieu chinh` của hai workbook chính. Mẫu: **“Điều chỉnh tăng/giảm cho hóa đơn ký hiệu …, số …, ngày … do xuất sai nội dung”**. Ký hiệu/số/ngày lấy đúng J–L cùng dòng (tham chiếu F0); chiều tăng/giảm theo tổng tiền, các hồ sơ chỉ giảm số lượng dùng giảm. Không đổi cột O tên hàng/dịch vụ, số tiền, ngày, tham chiếu hoặc công thức. Số hóa đơn điều chỉnh cũ vẫn truy tại `Dieu chinh`/`Chi tiet nguon`; diễn giải mới không còn ghi số khoản cũ.

- Chưa điều chỉnh: 8.410 hóa đơn / 13.568 dòng, đều giảm. SHA-256 mới: `8a7cdf7b45f533403461076df350e27e27bb685f35ab24cdb97e0da82a09f658`.
- Đã điều chỉnh: 2.032 hóa đơn / 41.790 dòng, 425 tăng và 1.607 giảm. SHA-256 mới: `f0056ce91b0a351aa25ea42c7d792b047ad6ce284778f0e073b86a98e79e9f37`.

Hai file đã có thay đổi cục bộ trước thao tác này; đã giữ bản trước sửa tại `scripts/data/import-reasons-20261014/before/`, không dùng baseline 12/10 để ghi đè. Công cụ import toàn workbook chạy chậm và được dừng; nội dung các ô thay đổi được dựng/xuất bằng artifact-tool riêng rồi ghép vào package gốc. Chỉ sheet import và sharedStrings được xử lý; các ZIP entry khác giữ nguyên. Kiểm độc lập bằng openpyxl: mọi ô ngoài phạm vi, công thức, định dạng, thứ tự sheet, merged cells, bảng và frozen panes giữ nguyên. Báo cáo `verified-0.json` và `verified-1.json` trong thư mục trên đều passed. Chưa nhập/ký/phát hành.

Kiểm trạng thái mới: `python scripts/layout-tools/verify-import-reasons-20261014.py 0 --final` và tương tự `1 --final`. Script kiểm 12/10 chỉ dùng cho baseline nghiệp vụ trước đổi diễn giải, không dùng hash cũ để kết luận file hiện tại sai. Không chạy lại builder cũ.

---

# Đã áp dụng phương án đã chốt vào cả 3 workbook chính — 12/10/2026

Người dùng đã cho phép thực hiện: **“ok làm đi”** sau khi được thông báo bước tiếp theo là cập nhật bộ import, sau đó yêu cầu kiểm tra kỹ lại rồi commit/push. Đã hoàn tất cập nhật cả 3 file tại `outputs/`, có bản sao nguyên trạng trước thay đổi tại `scripts/data/apply-confirmed-20261012/before/`. Mục này thay thế trạng thái “chưa cập nhật workbook import” của các mục lịch sử bên dưới. Chưa nhập phần mềm, ký hoặc phát hành.

- `HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx`: bổ sung 6 hóa đơn / 25 dòng (STT8405–8410), thành **8.410 hóa đơn / 13.568 dòng**. Đã cập nhật ghi chú 6 hồ sơ ở bảng chính và sheet `Kiem tra SL 0`; dữ liệu 27 dòng lịch sử của sheet kiểm tra vẫn giữ nguyên. Bản import dùng chi tiết xuất lại: dòng diễn giải giữ SL/giá/tiền trống, bỏ đúng 2 dòng không có giá trị của2796/2797.
- `HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx`: bổ sung **42 hóa đơn / 522 dòng** (STT1991–2032), thành **2.032 hóa đơn / 1.007 chuỗi / 41.790 dòng**. Đã cập nhật F4–F7 và công thức dự kiến của 21 chuỗi xử lý; xóa phương án phát sinh của 9 hồ sơ giữ nguyên. `Cho KT xac nhan`: cả19 đã chốt,13 chuỗi có import và6 giữ cả cặp. `42 - Can check 11`:8 chuỗi làm lại và3 giữ nguyên. Giữ nguyên toàn bộ1.990hóađơn đã chuẩn bị trước.
- `NGOAI_LE_F0_2024_T6_2025.xlsx`: thêm6cặp/12hóađơn vào cuối `Ngoai le` và `Chi tiet nguon`, thành **244 chuỗi /338 hóa đơn liên quan**. Bảng `Dieu chinh` vẫn lưu6chuỗi này như nguồn lịch sử, nhưng đã xóa F4–F7 và không có trong import. Không được cộng đôi các chuỗi khi kiểm phạm vi.

Tổng cộng thêm **48 hóa đơn /547 dòng trên27gốc**:21đảo điều chỉnh cũ và27giảm gốc. Phần mới dùng ngày dự thảo **12/10/2026**; giữ ngày08/10 và09/10 của dữ liệu đã chuẩn bị trước. Tên khách hàng bằng người mua; hình thức thanh toán lấy đúng cộtJ của bảng chi tiết xuất mới (`TM/CK`), email trống vì nguồn không có. Không dùng cộtI(CCCD) hoặc cộtN của bảng tổng làm hình thức thanh toán.

Đã thực hiện đúng6điểm và xác nhận dấu của5cặp giá âm. Gốc555 dòng6 chỉ đảo thành tiền−46.296, không thêm SL/giá. Gốc6/16 giảm số lượng, không thêm đơn giá/thành tiền dòng; tổng dự thảo0 nhưng ô thanh toán nguồn trống vẫn trống. Đảo655/657 theoXML. Giữ9186, chỉ giảm9072. Cả3giá trị tiền của8khoản lệch được đảo đúng số đã phát hành, không tự cân lại; các tổng đảo được tô vàng ở import.

Kiểm độc lập toàn bộ3bản đã xuất bằng công cụ đọc khác: giữ mọi giá trị/công thức ngoài phạm vi được phép, số/ngày tham chiếu, đủ8chữsố, bộ lọc bảng, tên/order sheet và frozen panes;21phép đảo đạt;48hóađơn mới không trùng; không có hóa đơn thuộc diện giữ nguyên trong import;0lỗi công thức. Rà toàn bộ **9.695gốc =9.417gốc đã chuẩn bị import +247gốc giữ nguyên(244ngoại lệ+3cùng sản phẩm) +31gốc đã khớp**, không còn gốc chưa phân nhóm. Đã xem trước/sau các vùng thay đổi. Nguồn xuất/XML/phản hồi kế toán không thay đổi.

Các file chính được thay từ bản staged đã đạt kiểm tra; SHA-256 sau thay giống hoàn toàn bản đã kiểm. Báo cáo: `scripts/data/apply-confirmed-20261012/final-verification.json`; kế hoạch, nguồn dòng và script kiểm cùng thư mục. Dùng báo cáo này cho trạng thái hiện tại; các audit/build cũ có số cố định1.990 hoặc8.404 là lịch sử trước bổ sung, không dùng để ghi đè bản mới.

Kiểm lại trên máy khác bằng `python scripts/data/apply-confirmed-20261012/verify.py --final`. Script tự tìm thư mục repo, kiểm hash nguồn/phương án/bản trước cập nhật, đọc cả 3 workbook hiện tại và ghi lại báo cáo JSON; không ghi Excel. Các bản `before/` được bàn giao cùng Git để kiểm độc lập. Ảnh xem trước, cờ chờ duyệt và bản xuất `staged/` chỉ giữ cục bộ.

SHA-256 hiện tại:

- Chưa điều chỉnh: `296c47cb3868a5ca4a308dc211dafc67782528b643910e1bd263acded6a34e4c`.
- Đã điều chỉnh: `9d041000c31027161714b66059cb9420a7fa02b3ed6d6c8b13a32ddb92cd2377`.
- Ngoại lệ: `e71cd26bd35e833b1caa587612ad3ccb1b3e6d8c1e699018539d809acc84eaa8`.

Việc còn lại thuộc bước nhập nháp trên phần mềm: **kiểm tra khả năng nhập8khoản đảo lệch tổng tiền**. Mẫu24cột không có trường tính chất dòng, nên kiểm tra/chọn tính chất diễn giải cho dòng chú thích nếu meInvoice không tự nhận. Kiểm tra ngày dự thảo, tham chiếu và nội dung trên bản nháp rồi mới ký/phát hành. Không yêu cầu kế toán chốt lại6điểm đã trả lời; chưa thử import thực tế trên meInvoice.

---
# Đã chốt toàn bộ 6 điểm và hoàn tất file phương án — cập nhật 12/10/2026

Mục này thay thế các trạng thái chờ làm rõ ở các mục lịch sử bên dưới. Người dùng đã xác nhận đủ 6 điểm và xác nhận thêm dấu của 5 cặp giá âm: **“Đúng, áp dụng cách hiểu này”**. Không hỏi lại các lựa chọn đã chốt.

Nguồn phản hồi mới nhất có SHA-256 `c55b6ba35463c8431a590f840ad95bddbf0f1b247e066e940b8ddb1f59f677e8`, được lưu nguyên bản tại `data/Bo sung/Ke toan/2026-10-11/c55b6ba35463/UVG_CAC_VAN_DE_CAN_KE_TOAN_XAC_NHAN_20261011.xlsx` kèm `PROVENANCE.json`.

Quyết định đã chốt:

- LT01–LT08: đảo đúng số đã phát hành trên XML, từng chỉ tiêu tiền hàng/thuế/thanh toán. Giữ độ lệch số học ngược dấu của nguồn; không tự cân lại một chỉ tiêu.
- Các cặp 1094/13038, 3113/13032, 4081/13033, 5555/13034, 5918/13035: giảm gốc với số lượng −1, giá dương; đảo điều chỉnh cũ với số lượng −1, giữ giá âm để thành tiền dương. Thuế/thanh toán đảo theo số đã phát hành. Dùng cặp số trong các ô định danh và XML; các đoạn mô tả phản hồi chép lặp gốc 3113 được giữ nguyên ở cột gốc, có quyết định cuối đúng ở cột mới.
- Gốc 555: chỉ đảo số tiền ở thành tiền, không tự thêm số lượng/đơn giá vốn không có trong nguồn.
- Gốc 2797: đúng sản phẩm Vợt bắt muỗi tự động 2 trong 1 thu hút các loại côn trùng KM003.
- Gốc 6 và 16: chỉ điều chỉnh giảm số lượng; không thêm giá hoặc chuyển ô tiền nguồn trống thành 0. Đảo hai điều chỉnh cũ 655/657 theo phương án đã bổ sung.
- CC01–CC06: giữ nguyên cả cặp ngoài đợt xử lý, gồm cả 3/508; quyết định này thay thế phần chuẩn bị xử lý 3/508 trước đây. Các hóa đơn xuất cho công ty/hộ kinh doanh, hoặc đã điều chỉnh tên người mua thành công ty/hộ kinh doanh, phải giữ nguyên toàn bộ hóa đơn trong chuỗi liên quan.

File hoàn chỉnh: `outputs/kt-review-20261011-01a12992/UVG_PHUONG_AN_DA_CHOT_20261011.xlsx`. Có 10 sheet, giữ nguyên phản hồi gốc và 67 hóa đơn nguồn cùng số/ngày/giá trị trống. Tổng 36 hồ sơ = 27 hồ sơ xử lý + 9 hồ sơ giữ nguyên; 48 hóa đơn dự kiến phát sinh = 21 đảo điều chỉnh cũ + 27 giảm gốc. Tất cả 6 điểm đã được ghi vào cột quyết định cuối. Đây là file phương án chuẩn bị, chưa phải chứng từ đã nhập/phát hành.

Đã rà quy tắc giữ nguyên trên dữ liệu hiện có gồm 48.874 hóa đơn, 3.780 liên kết và XML; truyền điều kiện loại trừ qua toàn chuỗi, gồm cả nội dung đổi tên người mua. Giữ 238 ngoại lệ trước và thêm 6 cặp được yêu cầu, thành 244 chuỗi/338 hóa đơn liên quan. Không có hóa đơn nào trong 10.394 hóa đơn dự kiến đang có bị ảnh hưởng bởi phần loại trừ mới. Ba hồ sơ cùng sản phẩm HH02/HH03/HH05 cũng được giữ nguyên riêng trong file rà soát.

Đã kiểm độc lập toàn bộ số liệu, 21 phép đảo, công thức tổng hợp, 10 bảng lọc, ngày và số hóa đơn; đã xem trước cả 10 sheet. Kết quả đạt, không có lỗi công thức. SHA-256 đầu ra: `40a601cbed6a1752ca1082094a9451bddfb9219765abf28401c6c40e70e0b1e9`. Báo cáo dưới `scripts/data/check-review/kt-workbook-20261011/`: `accountant-decisions.json`, `confirmed-decisions-audit.json`, `confirmed-plan-data.json`, `confirmed-output-verification.json`.

Ba workbook import chính và dữ liệu nguồn không thay đổi. Chưa cập nhật vật lý sheet chờ cũ bằng file phương án mới; không được nhầm việc đã chốt với việc đã nhập chứng từ. Chưa import, ký, phát hành, commit hoặc push.

---
# Đã chốt điểm 1: đảo số đã phát hành — 11/10/2026

Người dùng xác nhận trực tiếp: **“1. đảo số đã phát hành”**. Áp dụng cho 8 hồ sơ LT01–LT08: dùng đúng tiền hàng, thuế, thanh toán trên XML hóa đơn điều chỉnh đã phát hành, đảo dấu từng chỉ tiêu; không dùng bộ số kế toán vừa tính lại làm số gốc để đảo. Điểm 1/R01 đã được chốt, không hỏi lại lựa chọn này. Các điểm 2–6 trong câu trả lời chat vẫn chờ làm rõ.

Đã chuẩn bị riêng bộ số đảo ở `scripts/data/check-review/kt-workbook-20261011/confirmed-issued-reversals.json` và lưu quyết định tại `accountant-decisions.json`. Đọc lại trực tiếp 8 XML, kiểm số/ngày/thành tiền khớp nguồn; từng hóa đơn cũ + bộ số đảo cộng về (0,0,0). Ví dụ 1C25TUV/00000649 ngày05/05/2025: bộ đã phát hành (-381481,-30518,+411999), bộ đảo (+381481,+30518,-411999). Kiểm này chỉ áp dụng từng điều chỉnh cũ với phần đảo của chính nó, chưa có nghĩa toàn chuỗi gốc đã xử lý.

Do 8 XML nguồn có lệch hàng+thuế−thanh toán, bộ đảo giữ nguyên độ lớn và mang độ lệch ngược dấu. Không tự sửa một chỉ tiêu để làm bộ đảo cân bằng, không coi phép cộng về0 là chứng minh bản nhập/phát hành mới đã hợp lệ. Đây là bước chuẩn bị số theo quyết định đã chốt. Chưa sửa nguồn/ba workbook chính, chưa nhập, ký hoặc phát hành.

---
# Đã nhận file phản hồi kế toán — 11/10/2026

Người dùng gửi lại file tổng hợp từ `C:/Users/thuong ngo/OneDrive/Desktop/JOB UVG/JOB HOA DON/UVG_CAC_VAN_DE_CAN_KE_TOAN_XAC_NHAN_20261011.xlsx`, SHA-256 `275e9881d8f58eefc05e74ac1d6bf8ed2610b118c64abbeaf86a1c9b87af9f1d`. Đã lưu nguyên bản vào `data/Bo sung/Ke toan/2026-10-11/275e9881d8f5/` kèm provenance. **Cả 29 hồ sơ đang chờ đã có nguyên nhân, hướng xử lý và nội dung đúng; cả 7 hồ sơ tham khảo cũng có phản hồi mới.** Không được tiếp tục nói chưa có phản hồi kế toán hoặc yêu cầu trả lời lại toàn bộ. 21 dòng vẫn hiện “Cần bổ sung” vì riêng cột căn cứ còn trống; 8 dòng lệch tiền ghi căn cứ “Không có, tính lại cho đúng logic số liệu”. Đây là trạng thái đủ trường nhập, không phải kết luận nghiệp vụ.

Đã đọc/đối chiếu toàn bộ giá trị và công thức với bản gửi ban đầu: mã hồ sơ, ngày, nguồn tiền, công thức và liên kết giữ nguyên; ngoài các ô phản hồi và cột H mới ở `Da co can cu`, chỉ khác khoảng trắng cuối ở 8 ô ghi chú. Tám bộ số tiền kế toán tính lại đều cân bằng hàng + thuế = thanh toán. Chưa áp dụng phản hồi vào nguồn hoặc import. Hash ba workbook chính không đổi. Kết quả đầy đủ: `scripts/data/check-review/kt-workbook-20261011/accountant-reply-review.json`; script đọc: `inspect_accountant_reply.py`.

Các điểm cần làm rõ trước khi chuẩn bị dữ liệu xử lý theo phản hồi:

- LT01–LT08: “Xuất trái dấu ở tất cả các chỉ tiêu” chưa phân biệt đảo số đã phát hành hay số vừa tính lại. Đảo bộ số mới vẫn để dư so với XML cũ ở cả 8 trường hợp. Ví dụ 649 ngày 05/05/2025: số đã phát hành (-381481, -30518, +411999), bộ đúng KT ghi (-381481, -30518, -411999); đảo bộ đúng để dư thanh toán +823998. Cần rõ mục tiêu, hóa đơn tham chiếu và trình tự sửa/đảo, không tự chọn cách.
- DC03–DC06: cột I chép lại 13038 sai ở cả 4 dòng; 3 dòng cuối còn chép gốc3113. Cặp đúng trong chính file: 3113/13032, 4081/13033, 5555/13034, 5918/13035. DC02–DC06 còn cần tách dấu của phần đảo điều chỉnh cũ (đang âm tiền) khỏi phần giảm gốc: số lượng -1 × giá dương tạo tiền âm lần nữa, không hoàn lại tiền âm cũ.
- SL01: gốc1C24TUV/555 ngày31/10/2024 được yêu cầu giữ dòng diễn giải, âm số lượng, nhưng chỉ cho tiền hàng -46296. Chưa chốt số lượng hoặc cách xử lý tính chất dòng diễn giải; không tự suy ra -1.
- SL08: gốc1C25TUV/2797 ngày01/06/2025 có dòng KM003 trong nguồn, phản hồi ghi KM002. Cần xác nhận ý là bỏ đúng dòng2 KM003 của2797.
- TT01/TT02: gốc6 ngày07/01/2025 và16 ngày14/01/2025 đã cho hướng âm số lượng, nhưng chưa xác định giá trị thanh toán gốc trống hoặc cách xử lý điều chỉnh tăng655/657 liên quan. Không tự mở rộng quy ước0 của574.
- CC01–CC06: KT ghi giữ nguyên, không điều chỉnh; cần phân biệt giữ cả cặp ngoài đợt xử lý với giữ điều chỉnh cũ nhưng vẫn xử lý gốc. Riêng3/508 có khác so với quyết định chuẩn bị trước, chưa tự ghi đè quyết định cũ.

Các nội dung đã rõ được ghi nhận theo file: HH02/HH03/HH05 xác nhận cùng sản phẩm, không xử lý lại; HH01/HH04 đề nghị giảm đúng hàng gốc và tăng lại hàng trên điều chỉnh sai; SL03–SL06 giữ các dòng chú thích hàng tặng; SL02/SL07 đề nghị bỏ dòng0 xác định; DC01 đề nghị giảm gốc1061 và hoàn lại khoản trên12835; CC07 xử lý số lượng gốc9072, giữ9186. Đây là phản hồi phục vụ bước chuẩn bị tiếp, chưa phải chứng từ mới đã lập hoặc phát hành. Người dùng lượt này chỉ gửi file, không có yêu cầu sửa workbook/import hay phát hành.

---
# File mới giao kế toán xác định nguyên nhân và hướng xử lý — 11/10/2026

Theo yêu cầu người dùng, đã tạo riêng `outputs/kt-review-20261011-01a12992/UVG_CAC_VAN_DE_CAN_KE_TOAN_XAC_NHAN_20261011.xlsx`, không sửa ba workbook chính. File có 8 sheet: tổng hợp; 5 nhóm vấn đề cần phản hồi; hồ sơ đã có căn cứ; danh sách hóa đơn liên quan. Có **29 hồ sơ cần phản hồi, liên quan 53 hóa đơn**: 8 lệch tiền, 8 số lượng 0/trống, 2 thanh toán gốc trống, 5 khác hàng/mô tả, 6 cách ghi điều chỉnh. Nhóm số lượng bao gồm 6 hóa đơn chưa điều chỉnh được tìm thêm từ sheet `Kiem tra SL 0` và đối chiếu lại bảng xuất mới; không chỉ tổng hợp 23 hồ sơ đã nói trong chat.

Giữ riêng 7 hồ sơ đã có căn cứ/quyết định để không yêu cầu chốt lại, nhất là 3/508. Bảng nguồn gồm 67 hóa đơn phân biệt, có đủ ký hiệu, số, ngày, tổng tiền bảng kê/XML, nguồn file/dòng và ghi chú. Ô nguồn trống được giữ trống, tách khỏi XML=0. Các sheet cần phản hồi có câu hỏi cụ thể và ô vàng cho nguyên nhân, hướng xử lý, nội dung/số liệu đúng, căn cứ, người/ngày phản hồi. Trạng thái và tổng hợp tự tính; “Đủ trường phản hồi” không có nghĩa đã xử lý, nhập hoặc phát hành.

Đã kiểm tra bản XLSX xuất bằng công cụ đọc độc lập: 29+7 hồ sơ, 67 số hóa đơn dạng chuỗi đủ 8 chữ số, ngày, tiền và 4 khác biệt trống/0; bộ lọc bảng, cố định dòng/cột; các ô phản hồi ban đầu trống; không có lỗi công thức. Thử quy trình nhập phản hồi rồi khôi phục trống trước khi xuất. Hash cả ba workbook chính không đổi. Script tạo/kiểm và kết quả nằm ở `scripts/data/check-review/kt-workbook-20261011/`. Chưa có phản hồi mới của kế toán, chưa import/ký/phát hành, chưa commit/push.

---
# Đã nhận đủ XML điều chỉnh của 19 chuỗi chờ — 11/10/2026

Người dùng gửi tiếp `Hoa_don_dien_tu_11.10.2026_01.40.zip`, SHA-256 `abdf5a414af68fe942217b59580c869d7a451870280fb591544a4a6a71505682`, dung lượng 47.967.465 byte. ZIP có 213 XML + 213 PDF, đọc thành công toàn bộ và kiểm CRC ZIP đạt. Có đủ 16 số đã yêu cầu, cộng 3 điều chỉnh từ lô 13.25. Tính cả 508 đã có trước, hiện đủ **20/20 XML điều chỉnh thuộc 19 chuỗi chờ**. Không yêu cầu tải lại 16 hóa đơn hoặc toàn bộ 89 trang nữa. Đây là đủ XML điều chỉnh của nhóm này; không đồng nghĩa đủ XML gốc của toàn bộ dự án.

Đã giữ nguyên ZIP gốc và toàn bộ XML dưới `data/Bo sung/meInvoice/2026-10-11-pending-XML/`, cập nhật `PROVENANCE.json`. Các XML trùng với lô trước có cùng byte. Báo cáo nhận lô: `scripts/data/check-review/pending-xml-abdf5a414af6.json`. Trong 19 điều chỉnh có trong lô này, 15 khớp chính xác tổng tiền bảng xuất; 4 khác riêng thanh toán XML=0 / Excel trống. Ngày và tham chiếu cả 19 đều khớp. Không có sai khác số tiền khác ngoài trống/0.

Tám chênh tiền hàng + thuế − thanh toán đã được xác nhận **trong nội dung XML nhận được**, cùng giá trị với bảng xuất. Không phải chỉ do cách xuất Excel; chưa xác minh mật mã chữ ký hoặc trạng thái CQT. Cùng ký hiệu 1C25TUV, đơn vị đồng:

| Hóa đơn điều chỉnh | Chênh số học |
|---|---:|
| 00000649 | −823.998 |
| 00000561 | +136.000 |
| 00001819 | +2.000 |
| 00001167 | +1.000 |
| 00002145 | +900 |
| 00002619 | +1 |
| 00000626 | +1 |
| 00001821 | −1 |

Đã làm rõ các trường hợp sau bằng XML, không tự sửa ô nguồn hoặc workbook:

- 00004203 → gốc 1C25TUV/00002377; 00004204 → gốc 1C25TUV/00002378; 00006108 → gốc 1C25TUV/00001348: điều chỉnh thông tin người mua, không có danh sách dòng hàng, ba tổng tiền đều 0. Nội dung người mua được ghi trong `TTHDLQuan/GChu`; cần giữ nội dung này khi chuẩn bị xử lý tiếp, không coi thiếu dòng hàng là lỗi đọc XML.
- 00009186 → gốc 1C25TUV/00009072: điều chỉnh tên hàng, hai dòng TChat=4, ba tổng tiền đều 0; phần tổng hợp thuế ghi 8% với tiền hàng/thuế 0. Bốn hóa đơn trên có XML thanh toán=0 trong khi Excel trống; nguồn vẫn giữ riêng hai giá trị, chưa đưa vào import.
- 00002140 → gốc 1C24TUV/00000498: không có danh sách dòng hàng; phần `TToan/THTTLTSuat/LTSuat` ghi rõ 8%, tiền hàng 17.830.586, thuế 1.426.447, thanh toán 19.257.033. Đã có căn cứ XML cho thuế suất, không yêu cầu kế toán tìm lại thuế suất chỉ vì dòng chi tiết Excel không có.

Việc cần thông tin mới trong nhóm 19: kế toán chốt cách xử lý 8 hóa đơn lệch tiền; đối chiếu chi tiết gốc số lượng 0/trống của 1C24TUV/00000555 và 1C25TUV/00000113; làm rõ ô thanh toán gốc trống ở 1C25TUV/00000006, 1C25TUV/00000016. Tổng cộng 12 chuỗi. Gốc 1C24TUV/00000574 đã có quy ước riêng thanh toán gốc=0, không hỏi lại hoặc mở rộng quy ước.

Đính chính về 3/508 khi rà lại bàn giao: phần ngày 09/10 ghi người dùng đã chốt phương án như mẫu gốc00000003/điều chỉnh00000508 và đã áp dụng cho195chuỗi cùng dạng; phần bàn giao10/10 cũng ghi rõ phương án3/508 đã chốt. Không yêu cầu người dùng/KT chốt lại phương án này. Chuỗi3/508 vẫn nằm trong sheet chờ và chưa được thêm import; cần tiếp tục chuẩn bị dữ liệu theo quyết định đã có, tách việc chưa thực hiện khỏi việc chưa được chốt.

Bảy chuỗi còn lại của nhóm19 có căn cứ hoặc quyết định để chuẩn bị tiếp: 3/508; gốc574 theo quy ước riêng; gốc498/điều chỉnh2140 có thuế suất XML8%; bốn gốc1348,2377,2378,9072 với điều chỉnh thông tin tiền0 tương ứng6108,4203,4204,9186. Chưa đổi workbook hoặc tự suy ra đã được nhập/phát hành. Nhóm11chuỗi cầnchốt (2khác loại hàng,3khác mô tả,5đơn giá âm,1khoản tổng) vẫn chưa có phản hồi nghiệp vụ mới. Sheet chờ hiện vẫn19dòng; 12/7 là phân loại công việc sau khi nhậnXML, không phải số dòng import đã cập nhật.

Rà lại `audit_pending_review.py` sau bổ sung: đủ 20/20 XML điều chỉnh nhóm chờ và 11/11 nhóm cần chốt; bốn khác biệt XML=0 / Excel trống được liệt kê riêng, không ghi là khớp chính xác. Không có sai khác số tiền khác, sai ngày, sai tham chiếu hoặc xung đột nguồn mới; cả 30 gốc vẫn ngoài import. `passed=true` chỉ xác nhận kiểm tra dữ liệu đạt theo phạm vi này, không giải quyết thay các vấn đề nghiệp vụ. Hash cả ba workbook chính không đổi. Chưa import/ký/phát hành; các bổ sung cục bộ chưa commit/push. Các mục phía dưới là lịch sử theo thời điểm kiểm tra.

---
# Nhận lô XML tải lại 13.25 — 11/10/2026

Người dùng gửi `Hoa_don_dien_tu_11.10.2026_13.25.zip`, SHA-256 `3a74a6eb65a1cce744516bfc25eb9e0cc1e9f3cdcb682f9b339735b17ef0f974`, dung lượng18.070.627byte. ZIP có100XML +100PDF, ký hiệu1C25TUV, số00000543–00000642, ngày04–05/05/2025. Đã đọc toàn bộ100XML, không lỗi/trùng. Đây là lô100hóa đơn, chưa phủ toàn bộ khoảng thời gian người dùng đang tải lại.

Lô này có3/19 điều chỉnh cần tìm thêm trên máy hiện tại: 00000561→gốc00000056, 00000626→gốc00000102, 00000635→gốc00000113. Ngày, tổng tiền và tham chiếu cả3 khớp bảng xuất hiện tại. XML561 xác nhận chênh tiền hàng +thuế −thanh toán=+136.000đồng; XML626 xác nhận chênh+1đồng. Tổng XML635 khớp số học. Cả3 có dòngTChat1, số lượng/đơn giá/thành tiền dòng=0 trong khi tổng tiền hàng khác0; đã lưu chi tiết để tiếp tục kiểm tra nội dung, chưa tự sửa import.

Đã giữ nguyên toàn bộZIP gốc (có cảPDF) và trích nguyên byte cả100XML dưới `data/Bo sung/meInvoice/2026-10-11-pending-XML/`, kèm `PROVENANCE.json`. Thêm ngoại lệ `.gitignore` choZIP đúng thư mục này để không bỏ sót khi chuyển máy; chưa commit/push. Bằng chứng `scripts/data/check-review/pending-xml-3a74a6eb65a1.json`; script nhận lô mới `scripts/audit_pending_xml_archive.py <đường_dẫn_ZIP>`. Hash nguồn tải và các workbook không đổi. Chưa kiểm mật mã chữ ký/CQT trực tiếp.

Tính cả508 đã có từ trước, nhóm20điều chỉnh của19chuỗi chờ hiện cóXML4/20. Còn16điều chỉnh chưa cóXML trongrepo, cùng ký hiệu1C25TUV: 00000649,00000655,00000657,00001167,00001819,00001821,00002140,00002145,00002189,00002619,00002632,00002646,00004203,00004204,00006108,00009186. Tiếp tục nhận các trang/lô tải còn lại hoặc tải đúng16số này. Nhận định này chỉ nói độ phủ bản hiện tại, không khẳng định người dùng chưa cung cấp chúng trong haiZIP lịch sử1.436XML.

---
# Tiếp tục trên máy ở nhà — 11/10/2026

Theo yêu cầu người dùng, đã reset thay đổi cục bộ về bản vừa pull `af3386a`, khôi phục 16 file nguồn đang bị xóa trên máy này. Sau reset repo sạch; các thay đổi trong lượt rà dưới đây là kết quả mới. Người dùng xác nhận chưa có phản hồi kế toán mới, tiếp tục rà dữ liệu hiện tại.

Đã chạy trên máy này cả bốn kiểm tra chỉ đọc hiện hành: `verify_handoff.py`, `audit_supplement_57.py`, `audit_current_import_plan.py`, `audit_current_sources.py`; đều đạt. 18 nguồn nền khớp manifest, 10.394 hóa đơn dự kiến / 54.811 dòng import giữ đúng bản bàn giao, 986 chuỗi đã điều chỉnh cộng dồn về 0. Không sửa workbook hoặc nguồn. Kết quả audit57 được cập nhật danh mục hash có thêm XML733 và provenance hiện tại; kết quả nội dung vẫn 31 khớp chính xác, 26 chỉ khác trống/0.

Rà bổ sung bằng `scripts/audit_pending_review.py`: đối chiếu trực tiếp 61 hóa đơn thuộc 19 chuỗi chờ và 11 chuỗi cần chốt với bảng tổng, chi tiết, liên kết xuất mới và XML đang có trong repo. Kiểm lại ngày, số tiền, dòng gốc, số lượng 0/trống, số cộng dồn, tham chiếu, chi tiết/thuế suất XML và việc giữ ngoài import. Kết quả `scripts/data/check-review/pending-current-audit-20261011.json`: passed=true, không có sai khác mới, cả 30 F0 vẫn ngoài import; hash workbook và toàn bộ nguồn đã đọc không đổi. PASS ở đây là khớp dữ liệu hiện tại, không có nghĩa các vấn đề đang chờ đã được giải quyết.

Nhóm 19 chuỗi gồm 39 hóa đơn: 19 F0 + 20 điều chỉnh cũ. Trong phần XML đã trích và lưu vào repo, hiện chỉ tìm thấy gốc00000003 và điều chỉnh00000508 thuộc nhóm này; 19 điều chỉnh khác chưa có file XML rời trong repo (danh sách tại `pending_adjustment_xml_not_in_repo` của JSON). Đây là kết quả về độ phủ của bản chuyển máy, chưa xác định người dùng đã cung cấp hay chưa. Người dùng đã gửi hai ZIP tổng1.436XML ngày09/10; script `audit_downloaded_xml.py` chỉ lưu riêng237XML mục tiêu vào repo, không lưu toàn bộ1.436XML. Hai ZIP gốc không có trên máy hiện tại. Phải đối chiếu lại chúng trước khi yêu cầu người dùng xuất/gửi thêm. Câu trả lời trước trong chat gọi19XML là “còn thiếu” đã quá phạm vi bằng chứng và được đính chính tại đây.

Theo bảng tổng/chi tiết đã xuất, vẫn có 8 hóa đơn lệch tiền hàng + thuế so với thanh toán, 7 hóa đơn có ô tiền bảng tổng trống và 2 F0 có dòng số lượng 0/trống; các nhóm có thể giao nhau. Bảy ô tiền trống gồm gốc574 đã có quy ước riêng thanh toán=0 trong phép cộng dồn; vẫn giữ ô nguồn trống, không mở rộng quy ước. Chưa dùng các vấn đề của bảng xuất này để kết luận XML đã ký cũng lệch.

Ưu tiên tìm và đối chiếu XML của 8 hóa đơn lệch số học trong hai ZIP đã nhận trước khi đề nghị bổ sung. Chênh dưới đây là tiền hàng + thuế − thanh toán trong bảng tổng mới, đơn vị đồng; repo chưa có XML rời của các hóa đơn này để kết luận dữ liệu đã ký.

| Hóa đơn điều chỉnh (cùng ký hiệu 1C25TUV) | Chênh số học |
|---|---:|
| 00000649 | −823.998 |
| 00000561 | +136.000 |
| 00001819 | +2.000 |
| 00001167 | +1.000 |
| 00002145 | +900 |
| 00002619 | +1 |
| 00000626 | +1 |
| 00001821 | −1 |

Nhóm 11 chuỗi cần chốt: đủ 11 XML điều chỉnh, các trường đã rà khớp sheet và nguồn hiện tại, cộng dồn cả 11 về 0. Phân nhóm giữ nguyên: 2 khác loại hàng, 3 khác mô tả, 5 giảm bằng đơn giá âm, 1 điều chỉnh khoản tổng. Chưa có phản hồi nghiệp vụ mới nên chưa đưa các chuỗi này hoặc nhóm19 vào import. Sáu chuỗi thay thế chưa mở xử lý. Không đổi ngày dự thảo08/10 và09/10 thành ngày hiện tại. Chưa kiểm mật mã chữ ký/CQT trực tiếp, chưa import/ký/phát hành.

---
# Rà lại toàn bộ import hiện tại — 11/10/2026

Bộ bàn giao 11/10 gồm workbook chính đã sửa11đơn giá, nguồn bổ sungXLS/20XML nguyên byte, provenance/hash, các script audit và kết quảJSON, cùng hướng dẫn máy khác đã cập nhật. Tất cả nguồn cần cho audit hiện tại nằm trongrepo; không phụ thuộcDownloads máy cũ. Các câu “chưa commit/push” bên dưới là trạng thái lúc thực hiện trước bước bàn giaoGit; dùng git log và remote để xác nhận bản đã nhận.

Đã rà hai sheet import hiện hành: 8.404 hóa đơn / 13.543 dòng nhóm chưa điều chỉnh và 1.990 hóa đơn / 41.268 dòng / 986 F0 nhóm đã điều chỉnh; tổng 10.394 hóa đơn dự kiến / 54.811 dòng. Kiểm nguồn nền, bộ xuất mới, nguồn bổ sung, số tiền, ngày/tham chiếu, người mua/MST, dòng hàng và trùng STT; kiểm toàn bộ 41.268 dòng nhóm đã điều chỉnh với kế hoạch đã chốt (gồm 390 hóa đơn bổ sung và 11 đơn giá sửa theo XML). 986 chuỗi cộng dồn về 0; 19 chuỗi chờ vẫn ngoài import. Rà 237 XML điều chỉnh nguyên byte, hash/tổng/tham chiếu khớp kết quả đã lưu. Không xác minh mật mã chữ ký hoặc CQT trực tiếp.

Rà đầy đủ phát hiện thêm hai F0 00000608 và 00000733 trong phần 195 chuỗi bổ sung sau snapshot57. File xuất người dùng gửi ngày10/10 có đủ 4 dòng của608; ZIP9.32 đã có XML733, nay đã lưu bổ sung nguyên byte kèm provenance. Ngoại lệ609 cũng có 2 dòng trong file xuất này. Vì vậy bộ nguồn bổ sung hiện phủ60F0 (57 ban đầu +608 +733 +609 ngoại lệ), không còn thiếu chi tiết trong hai import. Các ghi chú thiếu57 phía dưới là lịch sử. Liên kết ngoại lệ60→61 chưa có ở bảng mới nhưng đã kiểm lại blob Git9c50b042, DS bị thay thế2024 dòng7; không coi là liên kết suy đoán.

Điểm cần kiểm khi xem trước meInvoice: 931 dòng/209 F0 nhóm chưa điều chỉnh và 1.254 dòng/305 F0 nhóm đã điều chỉnh có SL×đơn giá khác thành tiền dòng đang giữ từ nguồn; tổng2.185dòng/514F0. Chênh tối đa52đồng/dòng (nhóm đã điều chỉnh tối đa45); cả2.185chênh đều nằm trong0,5đồng×|SL|, phù hợp mất độ chính xác đơn giá khi xuấtExcel, chưa cóXML gốc toàn bộ514F0 để xác nhận đơn giá chính xác. Đây là chênh đã có trong nguồn, không phải import tự đổi thành tiền. Không lấy thành tiền/SL để tự suy ra đơn giá thay thế. Chưa sửa thêm workbook trong vòng rà này. Phải kiểm xem trước tùy chọn tự tính thành tiền/thuế; giữ số tiền nguồn và so tổng trước khi nhập cả lô. Các khác biệt trống/0 cũng giữ nguyên, không tự đồng nhất.

Bằng chứng: `scripts/data/check-review/full-current-source-audit-20261011.json`, `current-import-plan-audit-20261011.json`, `current-xml-237-audit-20261011.json`, và `scripts/data/untreated/final-audit.json` (rà lại hiện tại, passed=true). Script nguồn mới `scripts/audit_current_sources.py` và script rà kế hoạch `scripts/audit_current_import_plan.py` chỉ đọc workbook, ghi kết quảJSON riêng. Không chạy verifier lịch sử có copy2 vào file chính. Các nhóm nghiệp vụ chờ19, cầnchốt11 và thaythế6 chưa mở thêm; chưa import/ký/phát hành, chưa commit/push.

---
# Bổ sung đủ chi tiết 57 F0 — 10/10/2026

Đã hoàn tất và lưu workbook chính. Kết quả sau sửa: 31/57 F0 khớp chính xác chi tiết; 26/57 chỉ khác trống/0 đã ghi nhận, ngày và tổng tiền 57/57 khớp. `audit_supplement_57.py`: coverage_complete=true, content_checks_passed=true. `verify_handoff.py`: PASS, 18 hash nguồn nền không đổi, số lượng import giữ 8404/13543 và 1990/41268/986. Đã so sánh 638.284 ô có dữ liệu/công thức: chỉ 11 giá trị đơn giá và number format của đúng 11 ô thay đổi; font/màu/viền/căn lề/bảo vệ, công thức, cấu trúc, ô gộp, pane, validation và chiều cao dòng giữ nguyên. Render vùng sửa từ bản staged đã kiểm, hiển thị được phần 0,5 đồng. Chưa commit/push.

Công cụ artifact-tool đã sửa/render được nhưng xuất toàn workbook vượt heap Node 4 GB. Vì vậy bản cuối dùng fallback `scripts/stage_supplement_57_xml.py`: chỉ sửa 11 ô số và format trong sheet4.xml/styles.xml; tất cả ZIP entry khác giữ nguyên byte. `verify_supplement_57_edit.py` kiểm tra bản staged đầy đủ trước khi thay file chính. Không lưu workbook bằng openpyxl và không dựng lại sheet. Script chỉnh sửa có guard hash bản trước; không chạy lại sau khi bản chính đã cập nhật.

Đã nhận file chi tiết `Bang_ke_chi_tiet_hoa_don_da_su_dung_1791595024735.xls` (năm 2024) và ZIP `Hoa_don_dien_tu_10.10.2026_9.32.zip`. File chi tiết phủ 38/57 F0 đang dùng nguồn cũ (6 chưa điều chỉnh, 32 đã điều chỉnh); ZIP có đủ XML 19 F0 còn lại, ký hiệu 1C24TUV, số 00000724–00000732 và 00000734–00000743, ngày 25–27/12/2024. Đã lưu file chi tiết cùng 19 XML nguyên byte ở `data/Bo sung/meInvoice/2026-10-10-57-F0/`; `PROVENANCE.json` ghi hash ZIP gốc, entry XML và hash từng nguồn. Không cần bổ sung thêm nguồn để phủ nhóm 57 này. Manifest 18 nguồn nền giữ nguyên; đây là nguồn bổ sung riêng.

Đối chiếu ngày và tổng tiền cả 57 F0 khớp workbook. Có 23 F0 khớp chính xác chi tiết ban đầu, 26 F0 chỉ khác ô trống ở đơn giá/thành tiền trong file xuất mới với số 0 đang giữ từ nguồn cũ. Không tự thay ô trống bằng 0 hoặc bỏ dòng khuyến mãi. Có 8 F0 (00000724, 00000725, 00000728, 00000729, 00000731, 00000735, 00000737, 00000739) gồm 11 đơn giá trong import đã bị làm tròn lên 0,5 đồng so với XML; thành tiền dòng và tổng không lệch. Script `scripts/data/check-review/apply-supplement-57.mjs` sửa đúng 11 ô R của sheet Import dieu chinh theo đơn giá XML; không tái tạo sheet. Chỉ thay workbook chính sau khi `scripts/verify_supplement_57_edit.py` xác nhận toàn bộ giá trị/công thức/định dạng và cấu trúc khác được bảo toàn. Kết quả trước/sau và danh sách ô tại `scripts/data/check-review/supplement-57-*.json`; chạy lại `scripts/audit_supplement_57.py` và `scripts/verify_handoff.py` để kiểm tra bản hiện tại.

Đủ độ phủ không có nghĩa file xuất mới giữ số 0 giống nguồn cũ: 26 trường hợp trống/0 tiếp tục được ghi rõ trong kết quả. Không chạy builder cũ ghi đè các ô đơn giá chính xác. Chưa kiểm mật mã chữ ký hoặc CQT trực tiếp; chưa import/ký/phát hành. Nhóm 19 chuỗi ở Cho KT xac nhan là nhóm nghiệp vụ khác với 19 XML F0 bổ sung lần này, chưa thay đổi. Các việc còn lại về 19 chuỗi chờ, 11 chuỗi cần chốt, nhóm thay thế và xem trước meInvoice vẫn giữ theo bàn giao trước.

---
# Bộ bàn giao Git ngày 09/10/2026

Đã chạy verify_handoff.py trên bản cuối: 18 nguồn nền khớp manifest; import chưa điều chỉnh 8404 hóa đơn / 13543 dòng; import đã điều chỉnh 1990 hóa đơn / 41268 dòng / 986 chuỗi / 7019 dòng khuyến mãi. Workbook đã có hai sheet 42 - Da khop 31 và 42 - Can check 11. Nguồn xuất lại 10 file (gồm 3 ZIP), XML237 và bộ XML/ZIP3-508 được đưa cùng workbook, script, kết quả kiểm tra và tài liệu bàn giao. ZIP lớn tải ở Downloads không cần đưa lên Git vì đã lưu XML237 liên quan nguyên byte trong data.

Qua máy khác: git pull --ff-only origin master, đọc phần cập nhật đầu tài liệu này, chạy python scripts/verify_handoff.py. Không chạy các builder cũ để ghi đè import hoặc ghi chú. Việc còn lại: cập nhật nhóm 19 chuỗi đang giữ riêng (phương án3-508 đã được người dùng chốt), quyết định xử lý11chuỗi còn lại, bổ sung57F0nếu yêu cầu import chỉ từ nguồn mới, kiểm tra xem trước meInvoice trước khi nhập cả lô. Chưa import/ký/phát hành.

Các câu chưa commit/push ở các mục phía dưới ghi trạng thái lịch sử trước bộ bàn giao này; xem kết quả commit và xác nhận remote trong chat để xác định việc push hoàn tất.

---
# Tách 42 chuỗi vào hai sheet kiểm tra — 09/10/2026

Theo yêu cầu người dùng, đã thêm hai sheet trong workbook chính HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx: 42 - Da khop 31 (31 chuỗi, A4:AA34) và 42 - Can check 11 (11 chuỗi, A4:AA14). Mỗi dòng có định danh/ngày hóa đơn gốc và điều chỉnh, tên hàng, ĐVT, số lượng, đơn giá, tiền hàng/thuế/thanh toán hai phía, thuế suất, phân nhóm, cộng dồn W:Y bằng công thức, kết quả cần kiểm tra và tên XML. 31 chuỗi là khớp các chỉ tiêu đã rà, không phải xác nhận toàn bộ hồ sơ pháp lý. 11 chuỗi gồm2khác loại hàng,3khácmô tả,5đơn giá âm và1khoản tổng.

Đã render và kiểm tra đủ42khôngtrùng, giữ nguyên1.076.093ôcũ cảgiátrị/côngthức/địnhdạng, bảng/ôgộp/pane/datavalidation/chiềucaodòng. Không sửa AB hoặc import. Các công thức cộng dồn cả42chuỗiđều0. Bản chính chỉ lưu sau kiểm tra đạt. Bằng chứng scripts/data/check-review/split-42-verification.json; dữ liệu split-42-data.json, builder split-42.mjs, verifier verify-split-42.py. Không tạo thêm workbook báo cáo riêng. Chưa commit/push.

---
# Rà XML 42 chuỗi đã 0 còn lại — 09/10/2026

Đã đọc lại và kiểm hash XML42, đối chiếu nguồn chi tiết F0: cả42 có tổng dòng tiền hàng/thuế/thanh toán khớp tổng XML, tổng XML đảo đúng tiền F0; không cùng lỗi195/508. Phân nhóm:31 khớp tên hàng, ĐVT, đơn giá giữ nguyên, SL âm đảo đúng số lượng gốc, thuế8%, không có dòng tiền0 cần đánh dấu KM. 2khác loại hàng rõ: gốc3548→13072 (máy thành lõi lọcKA150S), gốc5905→13095 (lõi lọc thành máyKA600). 3khác mô tả:4026→13079 và7766→13099 thêm màn hìnhLCD;4953→13089 khác ProUpgrade/Pro-Upgrade. 5giảm bằng đơn giá âm nhưng SL+1:1094→13038,3113→13032,4081→13033,5555→13034,5918→13035. Đơn giá âm không tự nó là sai, cần phân biệt giảm giá với giảmSLtheo mục tiêu người dùng. 1giảm tổng:1061→12835, SL/đơn giá0 nhưng thành tiền dòng−499074, thuế−39926, thanh toán−539000, khớp tổng; khác508vì thành tiền dòng không0.

Kết quả scripts/data/check-review/remaining-42-xml-review.json. Chưa sửa workbook/AB/import. Không xác nhận toàn bộ hồ sơ pháp lý/chữ ký mật mã. Hướng dẫnMISA cho phép nghiệp vụ giảm đơn giá âm và điều chỉnh tổng nhưng phải đúng bản chất; không kết luận cứâmhoặcghi tổnglàsai. Đề xuất31 giữ nguyên về các chỉ tiêu đã kiểm,2khác loại hàng cần xửlý,9khácmô tả/giảmgiá/giảmtổng cần xácđịnh đúng nghiệp vụ trước khi thêmhóađơn mới.

---
# Bỏ ghi chú mới ở AB theo yêu cầu người dùng — 09/10/2026

Người dùng không cần note CHỐT LÀM LẠI. Đã bỏ phần ghi chú mới trên đúng 195 ô AB, giữ ghi chú cũ và phục hồi chiều cao dòng trước khi thêm note. F4/F5, ngày dự thảo và toàn bộ 390 hóa đơn / 491 dòng bổ sung giữ nguyên. Sheet import vẫn 1990 hóa đơn / 41268 dòng / 986 chuỗi; nhóm mới STT1601–1990. Không dùng nhãn CHỐT LÀM LẠI ở AB để lọc nữa. Mục bên dưới mô tả bước bổ sung trước khi người dùng yêu cầu bỏ note.

Đã so sánh 1.076.093 ô với bản trước: chỉ 195 giá trị AB thay đổi, toàn bộ giá trị/công thức/định dạng và cấu trúc khác giữ nguyên, trừ chiều cao đúng các dòng khôi phục. Bằng chứng scripts/data/check-review/remove-195-notes-verification.json. File chính đã lưu; chưa import hoặc phát hành.

---
# Đã bổ sung 195 chuỗi làm lại theo XML — 09/10/2026

Người dùng chốt áp dụng phương án như mẫu gốc 00000003 và điều chỉnh 00000508 cho 195 chuỗi cùng dạng: đảo khoản điều chỉnh giảm cũ bằng tổng tiền XML đổi dấu, sau đó giảm đầy đủ chi tiết F0. Đây là chốt chuẩn bị dữ liệu, chưa import, ký hoặc phát hành.

Đã đọc hai ZIP Hoa_don_dien_tu_09.10.2026_02.51.zip và 02.52.zip: 999 + 437 XML, số PDF tương ứng. Có đủ 237 điều chỉnh mục tiêu (19 + 218), không trùng, không lỗi đọc. Cả 237 khớp số tổng, tham chiếu gốc và ngày trong workbook hiện tại. Có 195 XML tiền hàng chi tiết 0 nhưng tổng tiền hàng khác 0; 42 trường hợp còn lại chưa thêm import. XML 00012885 xác nhận ngày 18/08/2025. Chưa kiểm tra mật mã chữ ký hoặc CQT trực tiếp. Các ghi chú cũ bên dưới về chờ XML237 đã được thay bằng kết quả này.

Workbook chính HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx:

- Sheet Dieu chinh: cập nhật AB đúng 195 dòng, bắt đầu bằng CHỐT LÀM LẠI 09/10/2026; giữ ghi chú trước đó dưới nhãn lịch sử. F4/F5 ghi phương án đã chốt và ngày dự thảo 09/10/2026. Số tiền dự kiến giữ đúng XML/gốc; BA:BC đều 0.
- Sheet Import dieu chinh: thêm STT 1601–1990, tức 390 hóa đơn / 491 dòng, vào A40787:X41277. Có 63 dòng khuyến mãi, chỉ đánh dấu 1 khi thành tiền gốc bằng 0. Hóa đơn đảo khoản có SL và đơn giá trống, thành tiền dòng bằng tiền hàng đảo từ tổng XML. Hóa đơn giảm gốc đủ dòng, SL đổi dấu, đơn giá giữ nguyên, thành tiền đổi dấu đúng nguồn. Các tổng chỉ ở dòng đầu mỗi hóa đơn. Tên khách hàng bằng Người mua hàng, giữ trống nếu nguồn trống.
- Toàn sheet hiện có 986 chuỗi / 1.990 hóa đơn / 41.268 dòng; ngày 08/10/2026 ở nhóm cũ giữ nguyên, nhóm mới 09/10/2026. Có thể sửa đồng nhất ngày các dòng cùng STT trước khi import.
- 19 chuỗi ở Cho KT xac nhan giữ nguyên, gồm chuỗi 00000003/00000508; lượt này chỉ mở đúng 195 chuỗi XML đã 0 được chốt. 42 chuỗi đã 0 còn lại vẫn ngoài import.

Nguồn XML đã lưu byte gốc trong data/Bo sung/meInvoice/2026-10-09-237-XML. Kết quả đọc XML tại scripts/data/check-review/zip-237-xml-review-20261009.json. Dữ liệu và quyết định 195 chuỗi tại 195-xml-data.json; script prepare-195-xml.py / apply-195-xml.mjs / verify-195-xml.py. Chi tiết F0 dùng nguồn đang có đã đối chiếu; không khẳng định import chỉ dùng bộ xuất mới. 57 F0 thiếu chi tiết trong bộ xuất mới vẫn chưa được bổ sung bởi hai ZIP này.

Đã xác minh 1.064.309 ô cũ: chỉ 982 ô thuộc phạm vi thay đổi (195 dòng x 5 ô, 7 ô hướng dẫn). Giá trị, công thức, format số, font/màu/viền, căn lề ngoài phạm vi, bảng, ô gộp, pane và data validation được bảo toàn. Đã render các phần sửa. 195 hash XML đối chiếu lại, 390 STT không trùng, tổng chi tiết từng hóa đơn bằng tổng tiền hàng, cả 195 chuỗi tiền hàng/thuế/thanh toán vẫn 0; không có lỗi công thức. Bằng chứng 195-xml-verification.json. Bản chính chỉ được thay sau khi đạt. verify_handoff.py cập nhật kỳ vọng mới 1990/41268/986.

Không chạy builder cũ prepare-adjusted-import.py / add-adjusted-import.mjs để ghi lại sheet: chúng còn quy tắc bỏ toàn bộ chuỗi đã 0 và số đếm 1600, sẽ làm mất 195 chuỗi mới nếu tái tạo không cập nhật logic. Đọc mục này và chạy verify_handoff.py trước khi tiếp tục. Các bổ sung ngày09/10 vẫn cục bộ, chưa commit/push.

---
# Chốt ngày00012885 theo xác nhận mới của người dùng09/10/2026

Người dùng xác nhận18/08/2025 làngàyđúng, thayxácnhậncũ05/07/2025. Đã sửa workbookchính `HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx`: `Dieu chinh!I903`=18/08/2025, `AB903`ghi xácnhậnmới vàgiữtoànbộghi chú237; `Chi tiet nguon!M1822`bổ sung căn cứ, `E1822`vốnđã18/08/2025 nênkhôngđổi. Sửa `USER_CONFIRMED_DATES` trongprepare_invoice_reports.py vàquyướcngày trongaudit_outputs_new_data.py đểkhôngquayvềngàycũ. Khôngchạy builder.

Artifact-tool sửa bản tạm; đã rendertrước/sau vàso sánh629.330ôcó dữliệu/toàn bộsheet: chỉI903,AB903,M1822thayđổi, thànhtiền/thuế/thanh toán/import giữnguyên. Bằngchứng `scripts/data/check-review/date-12885-verification.json`. Đãthayfilechínhsaukhiđối chiếu. Các kếtquảaudittrướcđó ghi05/07vsl18/08 làsnapshottrước xácnhận, đọc cậpnhậtnày trước. Khôngcònđểngày12885chờuserchốt; việcđọcXMLđộc lập vẫnchờlôtảixong.

Giải thích“thiếuchi tiết”: data mới cóhóađơn ởbảngtổng nhưngkhôngcódòngtrongbảngchi tiết, trongkhinguồncũcó. 8F0ởimportchưađiềuchỉnhvà49F0ởimportđãđiềuchỉnhđangdùngnguồncũ; khôngphảidòngimportbịbỏmất. Chưayêucầungườidùngxuấtthêmphầnnày.

---
# Kiểm tra trực tiếp toàn bộ outputs với data mới09/10/2026

Đã đối chiếu3workbook chính và2workbook lịch sử đang còn trongoutputs. Bằng chứng chi tiết/hash: `scripts/data/check-review/outputs-vs-new-data-20261009.json`; script chính `scripts/audit_outputs_new_data.py`. Không sửa output.

- Chưađiềuchỉnh:8.410F0 đối chiếu header;6giảđịnh thanh toántrống=0 giữ đúnguser. Import8.404hóađơn/13.543dòng khớptổnggiảmdấu, thamchiếu, ngàyF0, tên/ngườimua/MST vàhànghóa/SL/đơngiá/thànhtiền theo dữliệu kiểmđược.8hóađơn import lấychi tiết từnguồncũ vìbộmớithiếu;945hóađơn cókhác0/trống nguồn mới nhưng giữsố0nguồncũ, không tự chuyểntrống=0.
- Đãđiềuchỉnh:1.047F0/2.113hóađơn cũ khớp tiền tổng. I903/00012885 vẫn05/07/2025 theo xácnhậnuser trongkhi nguồn mới18/08/2025; chờXML, khôngsửangày. Import1.600hóađơn/40.777dòng gồm791giảmgốc và809đảo khoảncũ, không thấy lệchthamchiếungàyF0, người mua/MST hoặc tiền/dòng trongphạmvi kiểm.49gốc import dùngchi tiết cũ vìbộmớithiếu;534gốc cókhác0/trống.
- Ngoại lệ hiện tại:238F0/326hóađơn khớp tiền vàngày. Liênkết60→61 năm2024 chưa cótrongbộxuấtmới, nhưng đãcócăn cứGit9c50b042:DS bị thay thế2024 dòng7, giữnguyên. Liênkết61→2605 cótrongbảngthaythếmới2025 dòng19; khôngnhầm yêu cầuF2thamchiếutrựctiếpF0.
- Output lịch sử `outputs/ngoai-le-meinvoice-2024/NGOAI_LE_2024_THEO_CHUOI.xlsx`:190khối hóađơn, mộtkhối dòng112 là1C25TUV/00012764 còntrốngcả3tiền; datamớicó−416667/−33333/−450000, workbookngoạilệchính đãcósố. Khôngdùngfilelịch sửthayfilechính.
- Output lịch sử `outputs/kiem-tra-20261008/CAC_TRUONG_HOP_CAN_KIEM_TRA.xlsx`:19hóađơn phần tiền vàngày khớpbảngtổngmới; ghi chú508“lấyXML/chưakếtluận ởhóađơn haybảngkê” đãcũvì cóXML09/10. Ghi chú12885cầnđối chiếuXML vẫnápdụng nhưng cầnđọc kèmkếtquảnguồnmới đồngnhất18/08.

Không kếtluận mọioutputkhớp100%; không ghiđènguồncũ/ghi chúKT. ChờZIP237đểrà nội dungvà cậpnhậtAB khi cóbằngchứng.

---
# Bổ sung liên kết năm 2026 — đã kiểm tra ngày09/10/2026

Đã nhận bảng điều chỉnh `Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_1791535046109.xls` và thay thế `Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_1791535075764.xls`, đều xuất theo hóa đơn mới01/01/2026–09/10/2026. Đọc được1.288liên kết điều chỉnh và416liên kết thay thế, tổng1.704. Lần theo cả liên kết2024/2025/2026 không tìm thấy liên kết2026 thuộc nhómF0 ngày01/01/2024–30/06/2025. Kiểm tra chéo định danh với3workbook hiện có cũngkhông cóthamchiếu2026 tới các hóađơn trongbộđangxửlý.

Có60liên kết2026 thamchiếu hóađơn năm2025, nhưng ngàycủacác hóađơnđó từ11/10/2025–31/12/2025, ngoài kỳF0;1.644liên kếtcòn lại thamchiếuhóađơn2026. Vì vậy khôngcần thêmhóađơn2026 vàocácchuỗi/import hiện tại dựatrên bộxuấtđến09/10/2026. Bằng chứng vàhashnguồn: `scripts/data/check-review/links-2026-review.json`. Ghi chú trước đây “chưa đủliênkết2026” đã đượcgiảiquyết bằngđợtxuất này; không suy ra lịch sử sau09/10/2026. Khôngsửaworkbook hoặcpháthành. VẫnchờZIP/XML đểrà nội dung237chuỗi vàchốtngày00012885.

---
# Đối chiếu bộ xuất lại 09/10/2026 — tiền và ngày

Phạm vi người dùng chốt: chỉ F0 01/01/2024–30/06/2025 và hóa đơn liên quan; không xử lý mọi hóa đơn 2026. Đã đọc toàn bộ Excel/ZIP trong `data/Du lieu Meinvoice/Xuat_lai_2026-10-09`, đối chiếu với nguồn nền. Script chỉ đọc: `scripts/audit_source_refresh.py`; bằng chứng/hash và vị trí dòng: `scripts/data/check-review/source-refresh-review-20261009.json`.

- Bảng tổng trong phạm vi có tiền hàng/thuế/thanh toán, ngày và trạng thái không đổi so với nguồn cũ. Có 2.677 hóa đơn khác chi tiết chỉ do giá trị 0 ở nguồn cũ xuất thành ô trống ở nguồn mới; không tự áp giả định trống=0, không dùng bộ mới ghi đè builder/import. Không phát hiện thay đổi tên, đơn vị, số lượng hay số tiền khác ngoài nhóm 0/trống này ở các chi tiết cùng có trong hai bộ.
- Bộ mới có 65 hóa đơn tổng nhưng thiếu chi tiết; trong phạm vi gốc/chuỗi xác định bằng liên kết hiện có là 64. Cả 65 đều có chi tiết ở nguồn cũ nên bảo toàn nguồn cũ, không bỏ hóa đơn thiếu chi tiết khi cập nhật.
- 00012885 (gốc 00001314): chi tiết mới dòng35386, bảng tổng mới phần5 dòng1628 và liên kết mới dòng1238 đều ghi18/08/2025. Nguồn cũ chi tiết05/07/2025, tổng/liên kết18/08/2025. Bộ xuất mới đã đồng nhất nội bộ, nhưng vẫn khác ngày05/07/2025 người dùng xác nhận theo ảnh. Bảng chính `Dieu chinh!I903` vẫn05/07/2025; giữAB903 và chờXML, không tuyên bố ngày pháp lý đã được chốt hoặc đã sửa workbook.
- Không phát hiện lệch ngày nội bộ khác trong phạm vi xác định hiện tại. Chi tiết12885 là trường hợp duy nhất đổi ngày giữa bộ cũ và mới.
- Liên kết gốc77→186 nay xuất được trong bảng2024; workbook đã có tại dòng68 từ nguồn lịch sử, không phải chuỗi mới cần cộng thêm lần nữa.
- Chưa có liên kết điều chỉnh/thay thế2026: không thể khẳng định không có xử lý2026 liên quanF0 trong phạm vi chỉ từ bảng tổng2026. Không yêu cầu toàn bộ chi tiết2026; cần bổ sung liên kết nếu có phát sinh.
- `scripts/verify_handoff.py` vẫnpassed,18nguồn nền khớpmanifest; import chưađiềuchỉnh8404hóađơn/13543dòng và đãđiềuchỉnh1600hóađơn/40777dòng/791chuỗi giữ nguyên. Rà này không sửa workbook, không import/ký/phát hành. ZIP tải hàngloạt đang chờmeInvoice xử lý; sau khi cóXML mới cập nhật kếtluậnAB về237chuỗi.

---
# Bổ sung bằng chứng XML ngày09/10/2026

Người dùng đã cung cấp ZIP gốc00000003 và điều chỉnh00000508. Đã lưu nguyênZIP/XML dưới `data/Bo sung/meInvoice/2026-10-09-00000003-00000508`; đọc `KET_QUA_DOI_CHIEU.md` và `DOI_CHIEU_XML.json`. Các ghi chú lịch sử “chưa cóXML3/508” phía dưới đã được thay bằng kết quả này. Manifest18file vẫn là bộ nguồn nền; hashfilebổsung nằm trongJSON riêng.

508thamchiếu đúngF0; XML vàHTML đều có dòngTChat1=Hànghóa,dịchvụ, SL/đơngiá/thànhtiền0, thuếdòng−28487; phần tổng và tổnghợpthuếsuất ghi tiềnhàng−356085,thuế−28487,thanhtoán−384572. Chênh lệch đã xác nhận ngay trong dữliệuhóađơn, không chỉExcel. Không có trường tổngthanhtoáncấpdòng chuẩn ghi−28487 trongXML. Gốc7dòng có tổngchi tiết khớptổng11334130/906730/12240860.

Đang chờ chốt cách sửa nội dung/tínhchất/thànhtiền dòng và phần tổng vớiKT/MISA; chưa tự đưa chuỗi này vàoimport hoặc phát hành. Không kết luận cả195hóađơn khác giống508 nếu chưaXML. Vụlệchngày12885 vẫn chưa cóXML.

Đã rà lại237điềuchỉnh ngày09/10: tổng tất cả đảo đúngF0;196dạngtổng (195lệchchi tiết/bảngtổng);41cótênhàng,5kháctênmô tả gồm2đổi rõmáy/lõilọc: gốc3548→13072 và5905→13095. Có5đơngiáâm cầnđối chiếunghiệpvụ. Căn cứ và records tại `scripts/data/check-review/zero-compliance-review.json`. Rà này không sửaworkbook/AB, không chốtpháplý thayXML.

Các bổ sung09/10 hiện cục bộ, chưa commit/push; commitbe9ba64 đã push là bản bàn giao08/10.

---
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
