# UVG – đối chiếu chứng từ AMIS và hóa đơn meInvoice

## Kết quả

- Đã đọc 65,042 chứng từ AMIS và 10,853 hóa đơn meInvoice trong bảng tổng.
- 8,568 hóa đơn đủ điều kiện tìm ứng viên trực tiếp; hóa đơn đã bị sửa, thiếu dữ liệu hoặc lệch tổng chi tiết được giữ riêng.
- Chưa có liên kết nào được xác nhận bằng mã đơn hàng. Không có chứng từ nào được tự kết luận là cần xuất mới.

| Nhóm | Số chứng từ AMIS |
|---|---:|
| Chưa tìm thấy ứng viên trong tập hóa đơn đủ dữ liệu | 18,069 |
| Có ứng viên nhưng trùng lặp hoặc chỉ khớp mức tổng/model | 40,404 |
| Ứng viên duy nhất: khớp từng dòng, khác ngày | 51 |
| Chỉ khớp tiền: chưa khớp nhóm model và số lượng | 6,102 |
| Trả lại/giảm giá hoặc chứng từ cần kiểm tra riêng | 349 |
| Ứng viên duy nhất: khớp từng dòng và cùng ngày | 67 |

## Cách đối chiếu

1. Gộp dòng AMIS theo năm và số chứng từ. Loại dòng Tổng cộng khỏi dữ liệu. Tính tiền hàng sau chiết khấu, thuế và thanh toán.
2. Với meInvoice, chỉ tìm ứng viên trực tiếp trong hóa đơn mới/thay thế chưa thấy liên kết sửa tiếp, có bảng tổng và chi tiết khớp. Hóa đơn gộp bị điều chỉnh cần phân tích chuỗi riêng, không bị coi là không tồn tại.
3. Yêu cầu khớp chính xác ba số: tiền hàng, thuế, tổng thanh toán. Sau đó so cùng số dòng, họ model máy/lõi và số lượng từng dòng. Nhóm khớp từng dòng còn yêu cầu tiền hàng và thuế từng dòng bằng nhau.
4. Kiểm tra số ứng viên theo cả hai chiều. Một hóa đơn có thể cùng khớp nhiều chứng từ; không tự chọn hóa đơn gần ngày nhất.
5. Cùng ngày hoặc ứng viên duy nhất chỉ làm tăng mức ưu tiên kiểm tra, không chứng minh cùng đơn hàng. Ngày AMIS là ngày chứng từ, không được tự coi là ngày giao hàng.

## Giới hạn quan trọng

- AMIS do khách hàng chốt là số liệu mục tiêu theo xác nhận của người dùng. Phép ghép chưa chứng minh tính pháp lý của phương án sửa.
- Chuẩn hóa sản phẩm mới tới họ model và loại máy/lõi. Các biến thể Pro/Light, hàng tặng hoặc cách tách giá máy/lõi có thể khác nhau; không tự suy ra tương đương SKU.
- Không ghép nhiều đơn thành hóa đơn gộp bằng việc chọn tổ hợp số tiền. Khi thiếu bảng kê đơn gốc, cách đó có nhiều đáp án và không tạo bằng chứng chắc chắn.
- Thiếu mã đơn sàn, CQT và lịch sử phát hành/sửa sau tháng 6/2025. Chưa tìm thấy ứng viên không đồng nghĩa chưa xuất hóa đơn.
- CSV là dữ liệu phân tích để kiểm tra, không phải file import meInvoice.

## Hóa đơn giữ riêng khỏi bước ghép trực tiếp

- Có lịch sử sửa hoặc không phải hóa đơn mới/thay thế hiện hành: 2,270.
- Thiếu tổng tiền hoặc chi tiết: 15.

## Ví dụ ứng viên duy nhất khớp từng dòng

| Chứng từ AMIS | Ngày AMIS | Hóa đơn | Ngày HĐ | Tổng thanh toán |
|---|---|---|---|---:|
| BHHT2024000860 | 01/01/2024 | 1C25TUV / 00009855 | 30/06/2025 | 1,678,000 |
| BHHT2024000861 | 01/01/2024 | 1C25TUV / 00009841 | 30/06/2025 | 1,728,000 |
| BHHS2024000940 | 05/03/2024 | 1C25TUV / 00009745 | 30/06/2025 | 2,050,000 |
| BHHN2024000121 | 18/03/2024 | 1C24TUV / 00000113 | 20/03/2024 | 2,000,000 |
| BHHT2024001822 | 28/03/2024 | 1C25TUV / 00009633 | 30/06/2025 | 2,099,000 |
| BHHN2024000100 | 27/04/2024 | 1C24TUV / 00000175 | 22/04/2024 | 85,000,000 |
| BHHT2024008884 | 08/06/2024 | 1C25TUV / 00004811 | 11/06/2025 | 1,839,000 |
| BHHS2024004293 | 09/08/2024 | 1C24TUV / 00000243 | 27/05/2024 | 1,134,000 |
| BHHN2024000070 | 21/08/2024 | 1C24TUV / 00000394 | 21/08/2024 | 1,450,000 |
| BHHN2024000060 | 22/08/2024 | 1C24TUV / 00000387 | 20/08/2024 | 78,427,440 |

## Việc cần làm tiếp

- Kiểm tra từng ứng viên duy nhất bằng mã đơn sàn hoặc chứng từ giao dịch gốc và bảng ánh xạ SKU được khách hàng xác nhận.
- Với hóa đơn gộp/đã sửa: tính số dư theo toàn bộ chuỗi, đối chiếu với nhóm giao dịch thực tế; chưa lập file giảm hoặc xuất lại khi chưa xác định được nhóm giao dịch.
- Bổ sung CQT và các lần sửa sau tháng 6/2025 trước khi chốt danh sách phát hành.

## File kết quả

- AMIS_PHAN_NHOM_DOI_CHIEU.csv: một dòng cho mỗi chứng từ AMIS, có nguồn và dòng.
- UNG_VIEN_AMIS_MEINVOICE.csv: các cặp ứng viên duy nhất khớp từng dòng, có nguồn, dòng và độ lệch ngày.
- HOA_DON_THEO_NHOM_UNG_VIEN.csv: danh mục hóa đơn theo mã nhóm ứng viên. Nối với cột Mã nhóm ứng viên của bảng AMIS để xem đầy đủ các hóa đơn có thể khớp; không chọn một cặp khi nhóm có nhiều đáp án.
- ANH_XA_MODEL_AMIS.csv: họ model suy ra từ mã và tên hàng để rà soát.