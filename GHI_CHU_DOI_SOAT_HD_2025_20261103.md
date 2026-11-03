# Đối soát HD 2025 — bàn giao 03/11/2026

## 25 cặp kế toán đã kiểm tra, chưa thực hiện xử lý

- File Google Sheets `HD 2025`: https://docs.google.com/spreadsheets/d/1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw/edit . Tab `25 cap nghi trung - KT check`, gid `700000031`.
- Kế toán đã điền đủ 25 dòng H:J: giữ HĐ 1, ghi `bỏ HĐ` đối với HĐ 2. Tổng thanh toán của 25 HĐ được ghi chú bỏ là **23.275.929đ**. Cột K `Bằng chứng / HĐ xử lý` trống.
- **Người dùng xác nhận rõ: kế toán mới kiểm tra và note, chưa làm gì trên meInvoice.** Trạng thái bàn giao: **đã chốt phương án, chờ thực hiện**. Không diễn giải từ `bỏ HĐ` thành đã hủy, đã điều chỉnh hoặc đã thay thế; chưa tự chọn cách xử lý pháp lý hay tạo import.
- Bố cục đã làm gọn: mỗi cặp một dòng; D/E là liên kết mở HĐ nguồn; H:J nền vàng để kế toán ghi nhận; K:N là chi tiết thu gọn bằng dấu +. Ghi chú của kế toán giữ nguyên, không tự chép sang tab Ktoan.
- 25 cặp này là STT 1–25 trong ghi chú 27/10. So bảng kê chi tiết MTT 2025: 24 cặp khớp tất cả trường hàng hóa được xuất; 1 cặp `1C25MUV/00012320` và `00012473` chỉ khác khoảng trắng tên hàng. Mã hàng và cờ KM nguồn trống vẫn giữ trống.
- Ba cặp STT 26–28 đã có nhãn `Trùng mã (2 HĐ)` và hướng xử lý từ trước nên không đưa vào tab 25 cặp. Cặp `00012353/00012540` khác cách ghi combo, số lượng và đơn giá; dòng lõi lọc ghi **1 combo 2 lõi** so với **1 lõi**, không gọi là khớp tuyệt đối. Hai cặp `00012347/00012526` và `00012352/00012538` khớp cả dòng máy và lõi lọc 0đ.
- Nhãn `Trùng mã (2 HĐ)` là ghi chú cố định khi cùng mã xuất hiện ở hai F0 tại lần rà 22/10; không tự cập nhật khi một mã được xóa, không phải bằng chứng đã xử lý hóa đơn. 25 cặp còn lại không có nhãn này trong lần rà lịch sử và trên sheet khi kiểm.

## Tab gộp đã phân lại theo ký hiệu + số hóa đơn

- Đã cập nhật **đúng tab cũ** `HD gop nhieu chung tu`, gid `700000030`. Giữ đủ **22.022 chứng từ**, tổng thanh toán **30.694.466.425đ**; không sửa AMIS, MISA_HD_GOC hoặc Ktoan.
- Phạm vi là 22.022 chứng từ trong bảng gộp cũ, **không phải toàn bộ AMIS 2025**. 89 chứng từ đã có khóa `1C26MUV/số HĐ`; còn **21.933 chứng từ thiếu ký hiệu**. Không suy ra ký hiệu từ số HĐ, ngày hoặc việc khớp tiền.
- Bảng chính A6:S512 có **506 dòng theo dõi**: 89 khóa HĐ thật và 417 nhóm tạm chưa có ký hiệu. Trong 417 nhóm tạm: **379 nhóm nhiều CT**, **38 nhóm chỉ 1 CT**. Không gọi 506 dòng là 506 HĐ xuất gộp đã xác nhận.
- 89 khóa thật gồm **88 khớp VAT/thanh toán** và **1 chênh theo số dư chuỗi**. HĐ `1C26MUV/00000031` ngày 02/01/2026, mã đơn `251229MWJ0WUKR`, AMIS `BHHS2025023987` dòng 358 là 679.000đ; HĐ điều chỉnh `1C26MUV/00058521` ngày 03/07/2026 giảm hết ba khoản. Nguồn ghi cả hai HĐ hợp lệ; Ktoan dòng 37108 và AMIS L358 đã lấy đúng số dư **0**. Cần kế toán đối chiếu nghiệp vụ AMIS; không tự sửa sổ hoặc phát hành lại.
- Với dòng thiếu ký hiệu, tiền HĐ và chênh lệch **để trống**. Các cột F/L tra đúng ký hiệu + số HĐ trong Ktoan, giữ phân biệt ô trống và số 0, không lấy kết quả khi khóa thiếu hoặc trùng. Cột G/M chỉ tính chênh khi có tiền chuỗi dạng số.
- Tổng CT và tiền AMIS ở bảng chính dùng COUNTIF/SUMIF theo khóa nhóm ở S, lấy từ chi tiết hàng **517–22538**; tiêu đề chi tiết hàng **516**. 89 CT năm 2026 đã được tách khỏi nhóm số cũ, không cộng hai lần.
- Bố cục chính A:J; J để kế toán ghi chú. K:S (VAT, căn cứ và ghi chú cũ) thu gọn; các dòng một CT/đã khớp ở **387–512** và phần chi tiết **516–22538** thu gọn bằng dấu +. Bộ lọc A6:S512, cố định 6 hàng.
- Đã đọc lại toàn bộ 506 dòng chính và 22.022 dòng chi tiết: giá trị/công thức/ghi chú khớp dữ liệu chuẩn bị, 22.022 mã CT duy nhất, 22.022 liên kết mở AMIS, giữ tổng tiền, không lỗi công thức. Ghi chú cũ được chuyển vào phần căn cứ, không dùng ứng viên chưa chốt làm kết quả hiện tại.
- Kiểm bố cục qua API: cột A:J có độ rộng 65/125/115/65/155/175/165/170/115/280px, hàng chính 34px, bộ lọc/nhóm thu gọn đúng. **Lượt cập nhật cuối chưa kiểm được ảnh trình duyệt** do công cụ trình duyệt và chụp cửa sổ không phản hồi. Các bản bố cục trước của tab 25 cặp đã xem trực tiếp.

## Căn cứ cục bộ và cách tiếp tục

- Thư mục `outputs/ra-soat-hd-gop-20261103/` giữ kế hoạch 25 cặp, chi tiết so hàng, 417 khóa ứng viên cũ, đối chiếu chuỗi, 89 tham chiếu 2026, bằng chứng HĐ31/58521, cấu trúc cũ và JSON kiểm chứng cập nhật.
- `summary.json` là kết quả **rà soát chỉ đọc trước khi cập nhật**; `update-verification.json` là kiểm chứng **sau cập nhật tab**. Không dùng trạng thái trong summary cũ thay cho kết quả sau cập nhật.
- Nguồn so hàng: `data/Du lieu Meinvoice/Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls`. Các file bảng kê/quan hệ nguồn giữ nguyên. 49 hash nguồn/kế hoạch của bộ kiểm sau phát hành trước đây vẫn khớp khi kiểm.
- `before-cells-*.json` là sao lưu cục bộ lớn, được bỏ qua bởi Git; `before-structure.json` giữ cấu trúc để đối chiếu. Không có thông tin xác thực trong dữ liệu bàn giao.
- Khi tiếp tục, đọc lại Google Sheets theo khóa ký hiệu + số HĐ và mã CT, không dùng vị trí dòng cũ nếu sheet đã được sắp xếp. Không gửi lại nội dung cũ đè lên các ô kế toán đã ghi. Không khôi phục Excel cũ lên Google Sheets.
- Việc tiếp theo: đối chiếu ký hiệu còn thiếu bằng hồ sơ nối từng CT với HĐ; kế toán thực hiện phương án 25 cặp khi được yêu cầu rõ, rồi lấy bảng kê/quan hệ sau xử lý để kiểm chứng. Hiện **chưa nhập, ký, phát hành, hủy hay điều chỉnh HĐ** trong lượt này.
