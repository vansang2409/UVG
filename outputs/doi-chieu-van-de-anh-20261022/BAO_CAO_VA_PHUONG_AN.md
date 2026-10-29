# UVG — đối chiếu vấn đề trong ảnh và phương án xử lý

Ngày kiểm tra: 22/10/2026. Chỉ đọc nguồn; kết quả kiểm tra lưu riêng. Không sửa Google Sheets, dữ liệu kế toán, hai file import; chưa nhập, ký hoặc phát hành hóa đơn.

## 1. Kết luận đề xuất

Chọn **cách 2 có kiểm soát, kết hợp rà thủ công các ngoại lệ của cách 1**: lấy dữ liệu bán hàng AMIS làm tập mục tiêu đã thống nhất, lập đầy đủ chuỗi hóa đơn meInvoice, đối soát tự động các trường hợp rõ và chỉ giao kế toán xác minh nhóm nghi vấn.

Không áp dụng nguyên tắc “không khớp AMIS thì mặc định hóa đơn sai, điều chỉnh về 0”. Một liên kết chưa khớp có thể do ghi sai mã, thiếu chứng từ, khác kỳ, xuất gộp/tách, hoàn trả hoặc chuỗi điều chỉnh/thay thế. Việc thay đổi hóa đơn cần kết luận trên chứng từ cụ thể và phương án nghiệp vụ kế toán đã chốt.

Điểm then chốt: danh sách xuất lại phải căn cứ **phần nghĩa vụ đã được bao phủ bởi hóa đơn còn giữ sau phương án xử lý**, không chỉ kiểm tra đơn hàng từng có hóa đơn hay chưa. Hóa đơn cũ dự kiến được giảm về 0 không thể tiếp tục dùng làm căn cứ loại đơn khỏi danh sách cần xuất lại.

## 2. Nguồn và mức kiểm chứng

- Đã đọc hướng dẫn hiện hành `D:/UVG/TIEP_TUC_UVG.md`, báo cáo liên kết ngày 21/10 và hồ sơ chuỗi H2 cập nhật 22/10.
- Đã chạy lại toàn bộ phần tính toán của `scripts/audit_order_links_20261021.py` trong bộ nhớ, dừng trước mọi lệnh ghi của script gốc. Kết quả khớp hoàn toàn summary đã lưu; kiểm SHA-256 của 42 phần snapshot nguồn khớp bản audit.
- Snapshot Google Sheets được đọc ngày **21/10/2026**, gồm 65.495 dòng AMIS, 37.768 hóa đơn meInvoice T7–T12/2025, 26.550 dòng đối soát nội bộ và 65.436 dòng ánh xạ MISA. Lượt này không đọc lại Google trực tiếp; các thay đổi sau snapshot chưa được kiểm.
- Đã đọc thực tế hai workbook import hiện tại, kiểm số thứ tự hóa đơn và SHA-256. Bản chưa điều chỉnh có 8.410 hóa đơn dự kiến; bản đã điều chỉnh có 2.032, tổng 10.442. Hash cả hai khớp kết quả kiểm đã lưu ngày 14/10. Đây là số hóa đơn dự kiến trong file, không phải số đã nhập/phát hành.
- Đã kiểm dữ liệu chuỗi H2 hiện có: 37.087 gốc, 37.758 hóa đơn trong chuỗi, gồm 671 hóa đơn liên quan và 662 chuỗi có liên quan. Liên kết có theo sang 2026 để biết lịch sử; không đồng nghĩa tự mở rộng phạm vi import cũ sang 2026.

## 3. Đối chiếu trực tiếp hai vấn đề trong ảnh

### A. Đơn trước tháng 7 đã xuất hóa đơn trong T7–T12/2025

Tái tính tìm được **116 ứng viên khớp các trường đối soát ba nguồn**, tổng thanh toán hóa đơn **133.039.951 đồng**:

- Ngày nhận theo NB và ngày hạch toán AMIS đều nằm trong 23/05–30/06/2025.
- Ngày hóa đơn nằm trong 01/07–22/07/2025.
- Không có mã nào của nhóm này lặp ở hóa đơn khác trong snapshot meInvoice H2.
- Cả 116 khóa đều có trạng thái “Hoá đơn mới” trong nguồn Excel ghép hiện có; chưa thấy hóa đơn liên quan trong dữ liệu chuỗi đã ghép.
- **Không có khóa nào trong 116 trùng với các khóa hóa đơn ở sheet đối chiếu F0/F1/F2 của hai workbook import cũ.** Điều này phù hợp vì chúng có gốc phát hành từ tháng 7; tuy nhiên đây là kiểm khóa hóa đơn, chưa chứng minh không tồn tại đơn bán hàng trùng với một hóa đơn khác trong tập cũ.
- Có 35 người mua bị che một phần tên. Đây vẫn là ứng viên khớp dữ liệu, chưa phải 116 đơn được xác nhận tuyệt đối từ chứng từ giao hàng.

Ví dụ `1C25TUV/00010299`: hóa đơn 01/07/2025, nhận hàng và hạch toán 26/06/2025, chứng từ AMIS `BHHS2025010042`, thanh toán 2.779.000 đồng. Vị trí snapshot: meInvoice dòng 231; NB dòng 2; AMIS dòng 36.420.

Ngoài ra còn **1 ứng viên thiếu AMIS**: `1C25TUV/00010101`, NB dòng 2.395 ghi nhận 30/06/2025 nhưng mã chưa tìm được AMIS. Tổng nhóm đang nhận diện là 117 ứng viên, không phải 117 trường hợp đủ điều kiện loại khỏi xuất lại.

**Xử lý đề xuất:** đưa 116 vào danh sách cần kiểm để giữ hóa đơn H2, tạm chặn xuất lại cùng phần đơn trong khi xác minh chứng từ. Khi xác nhận hóa đơn được giữ và đúng đơn, mới loại phần tương ứng khỏi danh sách xuất lại. Trường hợp 10101 giữ riêng để bổ sung chứng từ.

### B. Một đơn có nhiều hóa đơn hoặc mã đơn ghi sai

Có **629 mã ghi chú xuất hiện trên nhiều hóa đơn**. Khi ghép với chuỗi hóa đơn hiện có:

| Phân loại tái kiểm | Số mã | Ý nghĩa và xử lý |
|---|---:|---|
| Các hóa đơn thuộc cùng một chuỗi gốc đã xác định | 517 | Không thể coi mã lặp là bằng chứng xuất trùng; xét vai trò gốc/điều chỉnh/thay thế và giá trị còn hiệu lực |
| Các hóa đơn thuộc nhiều gốc khác nhau | 112 | Ưu tiên rà chứng từ; có thể trùng thực, tách hóa đơn hoặc ghi nhầm mã |

Trong toàn bộ 629 nhóm, **108 nhóm có ít nhất hai hóa đơn trạng thái mới** theo snapshot Google. Đây là tín hiệu ưu tiên, chưa phải kết luận 108 đơn xuất trùng. Các nhóm nghi vấn cần hợp nhất theo khóa trước khi giao kiểm, không cộng số nhóm một cách cơ học.

Có **276 hóa đơn mới** mà mã note tìm được AMIS nhưng thuế và/hoặc thanh toán không khớp. Nhóm này có thể giao với 112 mã nhiều gốc; không cộng 276 + 112 thành số lỗi.

Ví dụ `1C25TUV/00011305`, ngày 07/07/2025, thanh toán **1.379.000 đồng**. Cả meInvoice và NB cùng note mã `250707F76E0BV1`, nhưng AMIS mà mã đó trỏ tới là `BHHS2025010663`, thanh toán **2.479.000 đồng**. Hai nơi ghi cùng mã không chứng minh liên kết đúng. Vị trí: meInvoice dòng 1.237, NB dòng 377, AMIS dòng 34.508.

**Xử lý đề xuất:** tìm lại chứng từ đơn/giao hàng, chi tiết hàng và hoàn trả để xác định đúng liên kết trước; chỉ lập phương án sửa hóa đơn sau khi xác nhận lỗi thực tế. Không tự xóa hóa đơn khỏi dữ liệu để làm tổng khớp.

## 4. Khoảng trống ảnh hưởng việc chốt xuất lại

1. **Thiếu tập đơn AMIS năm 2024 đầy đủ ở dữ liệu đang dùng.** Toàn bộ 65.495 dòng AMIS snapshot có ngày hạch toán năm 2025. Thư mục AMIS local hiện chỉ có bảng kê chi tiết hóa đơn đã sử dụng 2024; đây không phải sổ bán hàng đầy đủ thay thế cho AMIS 2024. Hai sổ AMIS cũ đang bị xóa cục bộ; không phục hồi hoặc thay thế trong lượt này. Do đó 116 chỉ là nhóm tìm thấy, không phải tổng mọi đơn 2024–T6/2025 đã xuất trong H2.
2. **Snapshot Google thiếu 251 khóa H2 có trong dữ liệu Excel đã ghép.** Các khóa này vẫn phải được đưa vào quần thể hóa đơn khi kiểm chống trùng. Không dùng Google làm toàn bộ lịch sử hóa đơn. Đã kiểm tập khóa trực tiếp; 321 khóa Google nằm ngoài chuỗi gốc H2 là phạm vi khác, cần giữ liên kết với gốc trước tháng 7, không tự bỏ.
3. **11.398 hóa đơn không có số tương ứng trong NB**, nên NB không đủ để làm nguồn bao phủ toàn bộ. Chưa khớp NB không có nghĩa chưa có đơn thực tế.
4. **436 hóa đơn trong dữ liệu chuỗi có thanh toán trống**; giữ trống, không thay bằng 0 hoặc chốt tổng của chuỗi là đầy đủ.
5. Bảng tổng hóa đơn xuất đến 09/10/2026, báo cáo liên kết bổ sung đến 22/10/2026. Cần đồng bộ thời điểm dữ liệu trước khi chốt thực thi, đặc biệt nếu đã phát sinh thao tác sau các mốc này.

## 5. So sánh phương án

| Phương án | Công sức | Độ tin cậy và rủi ro |
|---|---|---|
| Rà từng hóa đơn thủ công như cách 1 | Cao; khó đáp ứng nhanh khi có hàng chục nghìn hóa đơn | Có thể làm rõ từng trường hợp nhưng dễ lặp việc đã có dữ liệu chuỗi |
| Đối chiếu mã rồi đưa mọi trường hợp không khớp về 0 như cách 2 nguyên bản | Nhanh ở bước đầu, tăng khối lượng sửa/xuất lại | Rủi ro xử lý sai vì mã note không đáng tin, thiếu dữ liệu 2024 và chưa đủ độ phủ |
| **Đối soát tự động + rà ngoại lệ** | Tận dụng dữ liệu và bộ import đã có; tập trung nhân sự vào nhóm khó | Phù hợp nhất với hiện trạng; cần bổ sung nguồn thiếu và chốt bằng chứng cho nhóm nghi vấn |

## 6. Trình tự thực hiện đề xuất

### Bước 1 — Chốt bộ nguồn và tình trạng thực thi

- Lấy sổ bán hàng AMIS 2024 và 2025 đầy đủ, có mã đơn, chứng từ, ngày hạch toán, tiền hàng/thuế/thanh toán; bổ sung hoàn/trả, giảm giá và chi tiết hàng khi cần.
- Xác nhận hai file import giảm hiện tại mới là dự thảo hay đã có phần nhập/ký/phát hành. Kiểm tra này quyết định giá trị trước/sau; không suy ra từ file local.
- Chốt một mốc đối soát thống nhất; giữ nguyên bản nguồn và khóa tham chiếu. Không cần thay toàn bộ hệ thống hoặc viết ứng dụng mới.

### Bước 2 — Lập bảng nối đơn và chuỗi hóa đơn

- Khóa hóa đơn gồm ký hiệu + số; dùng ngày để kiểm. Khóa bán hàng gồm định danh đơn/chứng từ đã xác minh; mã note là tín hiệu tìm kiếm.
- Mỗi liên kết lưu dòng nguồn, căn cứ khớp, mức tin cậy, người xác nhận và quyết định xử lý. Cho phép một đơn–nhiều hóa đơn và một hóa đơn–nhiều đơn khi có bằng chứng phân bổ.
- Bắt đầu bằng 116 ứng viên; rà 112 nhóm nhiều gốc và 276 hóa đơn lệch tiền sau khi hợp nhất danh sách. Sau đó mở rộng tìm trên nguồn 2024 và 251 khóa Excel bổ sung.
- Khớp số tiền là điều kiện hỗ trợ, chưa đủ: kiểm ngày, người mua, chi tiết hàng và giao/hoàn khi trường hợp có rủi ro. Dung sai làm tròn của audit chỉ dùng phân loại, không sửa tiền gốc.

### Bước 3 — Phân loại để xử lý

| Nhóm đơn/phần đơn | Hướng xử lý |
|---|---|
| Có hóa đơn đúng được giữ sau phương án | Không xuất lại phần đã được bao phủ |
| Hóa đơn cũ dự kiến về 0, chưa có hóa đơn giữ lại khác bao phủ | Đưa vào tập cần lập phương án xuất lại sau khi xác nhận xử lý cũ |
| Nhiều gốc, sai mã/sai tiền hoặc nghi xuất trùng | Chờ kế toán xác minh và chốt phương án sửa cụ thể |
| Thiếu nguồn, thiếu liên kết hoặc thiếu tiền | Chờ bổ sung; không gán về 0 hay tự chọn chứng từ giống tiền |

### Bước 4 — Kiểm trước/sau theo đơn và kiểm tổng

Lập riêng hai trạng thái: thực tế đã phát hành và dự kiến sau phương án. Hóa đơn điều chỉnh phải tính cùng gốc; hóa đơn thay thế phải áp dụng quan hệ thay thế, không cộng gộp cả hóa đơn bị thay thế như doanh thu còn giữ.

Với đơn đã xác minh và có thể phân bổ đầy đủ, đối soát riêng từng chỉ tiêu:

`Giá trị bán hàng mục tiêu = Giá trị hóa đơn được giữ sau phương án + Giá trị xuất lại dự kiến`

Chênh lệch âm/dương là tín hiệu kiểm tra, không là lệnh tự động điều chỉnh. Kiểm theo đơn trước, sau đó theo tháng, thuế và tổng; hai đơn lệch ngược chiều có thể triệt tiêu ở tổng cộng.

### Bước 5 — Chốt đợt xử lý

Chỉ đưa nhóm đủ căn cứ vào danh sách thực thi; giữ nhóm chờ riêng cùng lý do. Giữ các ngoại lệ đã chốt của bộ import hiện tại. Cách này cho phép xử lý từng đợt có kiểm soát, không buộc rà tay toàn bộ hoặc đưa mọi khoản chưa khớp về 0.

## 7. Nguồn kiểm tra và đầu ra

- Báo cáo tái kiểm máy đọc: `D:/UVG/outputs/doi-chieu-van-de-anh-20261022/ket_qua_tai_kiem.json`, có đủ 116 khóa cùng dòng NB/AMIS/meInvoice, ngày và tiền.
- Snapshot và audit: `D:/UVG/scripts/data/link-audit-20261021/`.
- Chuỗi cập nhật: `D:/UVG/scripts/data/h2-chain-readiness-20261022/workbook-input.json` và `KET_QUA_KIEM_TRA.md` cùng thư mục. Đoạn cập nhật đầu báo cáo thay thế số lịch sử bên dưới.
- Hai file chỉ đọc để kiểm số lượng và hash: `D:/UVG/outputs/HOA_DON_CHUA_DIEU_CHINH_IMPORT_2_SHEET.xlsx`, `D:/UVG/outputs/HOA_DON_DA_DIEU_CHINH_IMPORT_2_SHEET.xlsx`.
- Script tái kiểm riêng: `D:/UVG/scripts/review_screenshot_scope_20261022.py`.

Đây là đề xuất xử lý dữ liệu dựa trên nguồn hiện có; chưa chốt số đơn cần xuất lại, số hóa đơn cần sửa hoặc số tiền cần phát hành.
