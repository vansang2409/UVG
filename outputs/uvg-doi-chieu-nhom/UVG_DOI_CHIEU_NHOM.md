# UVG – đối chiếu doanh thu theo nhóm

Dữ liệu AMIS đã được khách hàng chốt là mục tiêu. Chỉ đọc nguồn, không dùng mã đơn sàn và không phát hành hóa đơn. Phạm vi nguồn: năm 2024 và tháng 1–6/2025; chưa bao gồm các lần sửa sau kỳ xuất dữ liệu.

## Kết quả tổng thể

- Tổng thanh toán AMIS sau giảm trừ: **93.194.013.744 đồng**.
- Tổng thanh toán meInvoice tính được theo trạng thái và liên kết trong file: **95.594.027.824 đồng**.
- Chênh lệch AMIS − meInvoice: **-2.400.014.080 đồng**. Đây là chênh lệch cần giải thích, không phải số tiền cần xuất bổ sung hoặc điều chỉnh.
- 10,829 hóa đơn đóng góp hoặc cần kiểm tra; loại khỏi tổng 24 hóa đơn có trạng thái đã hủy/đã bị thay thế. Hóa đơn thiếu số tiền được giữ riêng, không được coi là giá trị bằng 0.

## Tổng hợp theo kỳ gốc

AMIS lấy ngày chứng từ. meInvoice lần theo bảng liên kết về ngày hóa đơn gốc; đây chỉ là quy đổi về kỳ hóa đơn gốc, chưa xác định được kỳ bán hàng của từng giao dịch trong hóa đơn gộp. Khoản không có liên kết được giữ ở dòng chưa xác định kỳ.

| Kỳ | Thanh toán AMIS | Thanh toán meInvoice theo file | Chênh lệch AMIS − meInvoice |
|---|---:|---:|---:|
| 2024-01 | 3.968.742.868 | 922.353.899 | 3.046.388.969 |
| 2024-02 | 1.940.417.064 | 1.915.829.108 | 24.587.956 |
| 2024-03 | 2.754.233.589 | 1.908.746.799 | 845.486.790 |
| 2024-04 | 5.278.284.182 | 5.791.779.884 | -513.495.702 |
| 2024-05 | 3.624.735.231 | 2.065.119.602 | 1.559.615.629 |
| 2024-06 | 4.450.947.091 | 6.640.277.772 | -2.189.330.681 |
| 2024-07 | 5.072.720.381 | 1.805.662.564 | 3.267.057.817 |
| 2024-08 | 5.256.787.164 | 7.458.497.270 | -2.201.710.106 |
| 2024-09 | 5.399.778.133 | 5.482.207.368 | -82.429.235 |
| 2024-10 | 5.506.787.240 | 3.691.741.578 | 1.815.045.662 |
| 2024-11 | 5.246.303.972 | 4.652.736.755 | 593.567.217 |
| 2024-12 | 5.424.441.483 | 8.838.064.336 | -3.413.622.853 |
| 2025-01 | 5.036.161.740 | 1.362.727.514 | 3.673.434.226 |
| 2025-02 | 7.244.115.869 | 4.247.666.645 | 2.996.449.224 |
| 2025-03 | 7.330.548.837 | 7.296.732.651 | 33.816.186 |
| 2025-04 | 5.988.519.553 | 15.795.347.870 | -9.806.828.317 |
| 2025-05 | 6.616.051.535 | 6.851.976.312 | -235.924.777 |
| 2025-06 | 7.054.437.812 | 8.160.721.146 | -1.106.283.334 |
| CHƯA XÁC ĐỊNH KỲ | 0 | 705.838.751 | -705.838.751 |

## Phần còn thiếu căn cứ

- Dòng đủ tiền khớp tổng; dòng thiếu tiền giữ riêng, chưa phân bổ lượng: 1,597 hóa đơn. Một hóa đơn có thể có nhiều vấn đề.
- Thiếu số tiền trên bảng tổng: 14 hóa đơn. Một hóa đơn có thể có nhiều vấn đề.
- Chưa phân bổ chắc chắn tiền theo sản phẩm: 1,176 hóa đơn. Một hóa đơn có thể có nhiều vấn đề.
- Tiền hàng cộng thuế không bằng tổng thanh toán: 8 hóa đơn. Một hóa đơn có thể có nhiều vấn đề.
- Thiếu liên kết tới hóa đơn nguồn: 299 hóa đơn. Một hóa đơn có thể có nhiều vấn đề.
- Chưa xác định kỳ hóa đơn gốc: 299 hóa đơn. Một hóa đơn có thể có nhiều vấn đề.
- Tổng thanh toán chưa phân bổ chắc chắn theo sản phẩm: 8.631.141.501 đồng (số cộng đại số, gồm cả khoản tăng/giảm).
- Tổng thanh toán chưa xác định kỳ hóa đơn gốc: 705.838.751 đồng.
- AMIS có 0 dòng mà công thức doanh thu sau giảm trừ + thuế chưa khớp thanh toán; chi tiết ở file kiểm tra.

- Ngoài nhóm chưa phân bổ sản phẩm, còn 1.219.544.281 đồng thuộc nhóm chưa nhận diện được họ model từ tên hàng.

## Cách tính và giới hạn

- AMIS: doanh thu = doanh số bán − chiết khấu − giá trị trả lại − giá trị giảm giá; số lượng = lượng bán − lượng trả lại. Thuế và thanh toán dùng dấu có sẵn trong báo cáo. Đã kiểm tra tổng dòng với dòng Tổng cộng của hai file.
- meInvoice: loại hóa đơn có trạng thái đã bị thay thế/đã hủy; giữ hóa đơn gốc bị điều chỉnh, cộng khoản điều chỉnh, giữ bản thay thế còn hiện hành theo trạng thái file. Các liên kết đưa khoản sửa về cùng chuỗi; trường hợp thiếu liên kết hoặc vòng lặp được đánh dấu.
- Phân bổ tiền meInvoice vào model từ các dòng có số tiền và thuế thực tế khi tổng các dòng đó khớp bảng tổng. Nếu còn dòng trống tiền, giữ chúng ở khoản chưa phân bổ lượng, không tự coi từng dòng là 0 đồng. Nếu không khớp tổng hoặc thiếu chi tiết, giữ toàn bộ giá trị hóa đơn ở nhóm chưa phân bổ sản phẩm. Không phân bổ tùy ý khoản điều chỉnh chung vào model.
- Họ model phân biệt máy/thiết bị và lõi lọc. Không suy ra biến thể Pro/Light, không tự kết luận cùng SKU. Số lượng trên dòng điều chỉnh cần kiểm tra bản chất: số lượng diễn giải không luôn là lượng tăng/giảm thực tế.
- Kênh chỉ nhận diện khi tên khách hàng/người mua ghi rõ TikTok, Shopee/Shoppe/SP hoặc Lazada. Phần còn lại giữ Khác/chưa xác định kênh, không tự coi là ngoài sàn. File kênh trình bày hai nguồn riêng, không lấy kênh chưa xác định để bù sang một sàn.
- Dữ liệu có thể chưa đủ CQT và các lần sửa sau tháng 6/2025; số dư là theo bộ file hiện có, chưa phải số dư đã xác minh đầy đủ.
- Bảng chênh lệch không phải file import hoặc hướng dẫn giảm toàn bộ hóa đơn. Cần giải thích lệch kỳ, lệch sản phẩm, giảm trừ và lịch sử trước khi chốt hướng xử lý.

## File để kiểm tra

1. DOI_CHIEU_THANG_MODEL.csv: doanh thu, thuế, thanh toán và lượng theo tháng/họ model của hai nguồn.
2. SO_LIEU_THEO_KENH.csv: số liệu theo kênh nhận diện, giữ phần không xác định riêng.
3. HOA_DON_VA_KY_GOC.csv: từng hóa đơn, trạng thái, chuỗi nguồn, kỳ gốc và vấn đề cần kiểm tra.
4. SO_DU_THEO_CHUOI_HOA_DON.csv: số dư theo chuỗi, liệt kê các hóa đơn đóng góp.
5. Các file chi tiết AMIS/meInvoice chứa file nguồn và số dòng để truy ngược.

## Việc tiếp theo

Ưu tiên giải thích hóa đơn thiếu liên kết và khoản chưa phân bổ theo sản phẩm. Sau đó rà soát từng nhóm tháng/model lớn, kiểm tra bảng ánh xạ hàng và thời điểm ghi nhận. Phần không đủ căn cứ giữ lại; chưa phát hành theo chênh lệch tổng.