# Bàn giao phân loại hóa đơn năm 2025 — 05/11/2026

Kế toán giữ hóa đơn đúng và xử lý phần sai hoặc thiếu. Bảng phân loại có 65.494 chứng từ AMIS, tổng thanh toán 89.077.087.173đ và VAT 6.598.300.351đ. Đã đọc lại file Excel để kiểm số dòng, mã đơn dạng chữ, tiền, VAT và tổng hợp.

| Việc cần làm | Chứng từ |
|---|---:|
| Đề xuất giữ theo đối soát mã đơn và tiền | 34.560 |
| Kiểm tra và sửa liên kết AMIS–hóa đơn | 112 |
| Rà soát hóa đơn và số dư chuỗi | 327 |
| Ứng viên xuất lại hoặc bổ sung | 29.353 |
| Cần kế toán xác minh thêm | 1.142 |

Bắt đầu với tab **Sửa liên kết**. Khi một nhóm có chứng từ nghi ghép nhầm, sửa liên kết trước rồi đối chiếu lại tiền và VAT. Nhóm giữ cần so chi tiết hàng hóa. Nhóm ứng viên cần xác minh nghiệp vụ, liên kết và các hóa đơn khác trước khi lập. Tab **HĐ cần kiểm riêng** gồm 1.601 HĐ, trong đó có 25 HĐ kế toán đã ghi chú bỏ, chờ đối chiếu bằng chứng xử lý. Tổng ở trên tính theo chứng từ AMIS.

File làm việc tại máy là `outputs/phan-loai-2025-20261105/PHAN_LOAI_XU_LY_HD_2025.xlsx`. Nguồn là ảnh chụp AMIS ngày 05/11/2026, bảng kê meInvoice đến 29/10/2026 và hồ sơ bổ sung ngày 01/11/2026. Dữ liệu nguồn và file Excel chi tiết lưu cục bộ. Bản Google Sheets tạo nhầm đã được chuyển vào thùng rác. Trạng thái công việc là phân loại, chờ kế toán chốt.

Khi có đủ ảnh chụp nguồn tại máy, chạy lần lượt các script trong `scripts/`: `classify_2025_20261105.py`, `prepare_classification_workbook_20261105.py`, `write_classification_streaming_20261105.py`, `verify_classification_xlsx_20261105.py`. Kiểm tra dữ liệu mới nhất trước khi lập hóa đơn.
## Kiểm tra map tiền trên bản sao của kế toán

- Chỉ xử lý các dòng AMIS có đủ số và ký hiệu HĐ nhưng chưa lên tiền. Đã xử lý 478 dòng gồm 421 dòng thiếu tiền ở một bước trong chuỗi Ktoan, 55 dòng có HĐ năm 2026 được bổ sung từ bảng kê meInvoice và 2 dòng cần dò theo HĐ thay thế. Cột L đã được đọc lại sau khi cập nhật.
- Cột O của AMIS ghi trạng thái tại thời điểm kiểm tra. Đã kiểm 35.646 dòng dữ liệu gồm 35.559 dòng đã map tiền, 58 dòng điều chỉnh về 0, 2 dòng thay thế về 0, 1 dòng HĐ gốc 0 đ, 25 dòng thiếu số hoặc ký hiệu và 1 dòng chưa map vì hai số HĐ chung một ô.
- Dòng chưa map vẫn chờ sửa công thức L. Hai HĐ cùng ký hiệu đều có trong Ktoan, tổng tiền 200.000 đ.

## BTL 2025 trong file đối soát

- Tab BTL 2025 có 321 phiếu và 411 dòng chi tiết. Tất cả đã ghép theo Số CT bán hàng lấy trực tiếp từ AMIS Kế toán và khớp mã hàng với chứng từ bán gốc.
- Giá trị trả lại 441.952.027 đ, VAT giảm 35.356.147 đ và thanh toán giảm 477.308.174 đ. Đã đọc lại tab và kiểm tổng. Phần đối chiếu các khoản trả với hóa đơn điều chỉnh trên meInvoice còn cần thực hiện.

## Tiến độ đối soát ngày 06/11/2026

- Tab AMIS trong bản sao HD 2025 chỉ có chứng từ hạch toán tháng 7–12/2025. Khi dò HĐ cho chứng từ AMIS năm 2025, dùng cả bảng kê meInvoice năm 2026. Đã xác nhận có HĐ phát hành tháng 01/2026 cho các chứng từ hạch toán tháng 4–5/2025.
- Tab **Sửa liên kết** có 114 dòng. Đã cập nhật và đánh dấu **Đã thực hiện** cho 30 dòng trên file đối soát, gồm 12 dòng sửa ký hiệu và bỏ ngày khỏi ô số HĐ, 17 dòng bổ sung liên kết và dòng 120 khớp tên khách hàng, ngày, tiền và VAT. Cột L trên AMIS bản gốc đã trả đúng tiền meInvoice cho cả 30 dòng. Đã bổ sung 28 HĐ còn thiếu vào tab Ktoan của bản gốc. Dòng 22 giữ số tiền theo bảng kê, lệch AMIS 1 đ VAT và thanh toán, có khả năng do làm tròn.
- Trong 84 dòng còn lại, đã khoanh 69 dòng thuộc 57 nhóm nghi gắn nhầm. Sau khi loại các dòng này khỏi phép tính, chứng từ còn lại trong từng nhóm khớp tiền và VAT với HĐ. Lượt dò mã đơn chưa xác nhận được HĐ khác cho 69 dòng. Riêng dòng 49 và 104, tên người mua trên HĐ khớp chứng từ còn lại trong từng cặp. Cần xác minh qua đơn TikTok Shop hoặc phiếu giao hàng trước khi sửa AMIS.
- Tab **Giữ HĐ** đã rà 475 dòng, gồm 106 dòng có ghi chú kế toán và 372 dòng đánh dấu lệch hàng, giao nhau 3 dòng. Khi AMIS phân bổ giá combo cho từng sản phẩm, so mã hoặc tên hàng, biến thể, số lượng, tổng thanh toán và VAT. Dòng 9173 cần xác minh số lõi lọc thực giao.
- Tab **6 đơn trả hàng** ghi mã đơn, ngày hạch toán AMIS và chuỗi HĐ đã điều chỉnh về 0. Chưa tìm thấy chứng từ bán gốc tương ứng trong 411 dòng BTL năm 2025. Cần đối chiếu chứng từ trả hàng hoặc hoàn tiền trước khi chốt xuất bổ sung.

## Đối chiếu AMIS xuất lại ngày 08/11/2026

- Hai file nguồn tại `data/0811/` gồm `Ban_hang.xlsx` và `So_chi_tiet_ban_hang.xlsx`; đã đọc đủ hai tab mỗi file. Danh sách bán hàng có 65.493 chứng từ năm 2025, trong đó 35.636 chứng từ tháng 7–12. Tiền và VAT của toàn bộ chứng từ bán hàng khớp số cộng từ sổ chi tiết. SHA-256 nguồn lưu tại `outputs/doi-chieu-amis-0811-20261108/manifest.json`.
- Đối chiếu với tab AMIS của [bản sao HD 2025](https://docs.google.com/spreadsheets/d/1C5hoPDvZKB1vGFSUG9rzChcORt2L9NnsqETVHxPui9I/edit?gid=1264954379#gid=1264954379): link có 35.646 chứng từ; có 10 chứng từ tháng 12 không tìm thấy trong cả hai bản xuất mới, tổng thanh toán 17.221.701 đ và VAT 1.275.682 đ. Trong 35.636 chứng từ còn khớp số, tiền và VAT không đổi; riêng BHHS2025023296 đổi diễn giải từ 19/12 thành 30/12, ngày hạch toán vẫn 30/12/2025. Danh sách chi tiết lưu tại `membership-differences.json` và `changed-fields.json` trong cùng thư mục kết quả.
- Lọc cột M bỏ 0 trên bản sao có 576 dòng, gồm 564 lệch tiền và 12 lỗi `#N/A`. Nếu dùng danh sách AMIS mới và giữ bảng map meInvoice của bản sao, còn 567 dòng, gồm 555 lệch tiền và 12 lỗi; đây là tính lại trên dữ liệu, chưa thay nguồn của bản sao.

## Đối chiếu HD CTY 2025 theo số dư cuối chuỗi — 08/11/2026

- Đã cập nhật [tab HD CTY 2025 của bản gốc](https://docs.google.com/spreadsheets/d/1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw/edit?gid=622685968#gid=622685968), vùng `S4:AB412`. Có 408 dòng hóa đơn thuộc 391 chuỗi; 407 dòng ghép được chứng từ AMIS. Cột I–K vẫn là số của từng hóa đơn, không phải số dư cuối chuỗi.
- S–T là VAT và thanh toán AMIS xuất 08/11. X–Z là doanh thu, VAT và thanh toán cuối chuỗi. U, V, AA lần lượt là chênh VAT, thanh toán và doanh thu theo AMIS trừ số dư cuối chuỗi. W ghi kết quả; AB ghi các hóa đơn trong chuỗi và nguồn.
- Điều chỉnh được cộng có dấu; hóa đơn đã bị thay thế không cộng vào số dư. Đã xét phát sinh liên quan năm 2026. Nguồn meInvoice đến 29/10/2026, bổ sung phát hành 01/11/2026; tab Ktoan đọc lại 08/11. 272 dòng đối chiếu được với số hiện tại AD:AF của Ktoan và khớp phép dựng đầy đủ; 136 dòng trước tháng 7 bổ sung từ bảng kê. Cả 408 dòng đều xác định được chuỗi và số dư.
- Kết quả cuối: 349 dòng khớp, 30 dòng lệch 1–2 đ, 12 dòng lệch lớn hơn 2 đ, 16 dòng cần kiểm tra liên kết chứng từ với nhiều chuỗi và 1 dòng chưa có chứng từ AMIS. Các nhóm này loại trừ nhau. Đếm trực tiếp chênh số: 36 dòng lệch thanh toán, 47 dòng lệch VAT, 49 dòng lệch ít nhất một khoản; có dòng lặp cùng chuỗi nên không cộng các số này hoặc cộng tiền chuỗi theo từng dòng.
- Có 13 dòng khớp tổng thanh toán nhưng lệch VAT và doanh thu: 8, 55, 64, 76, 81, 92, 104, 127, 136, 140, 150, 172, 339. Trong đó 12 dòng lệch 1 đ; dòng 64, HĐ 1C25MUV/00031644, VAT AMIS cao hơn 3 đ và doanh thu thấp hơn 3 đ. Đã đọc lại doanh thu thực tế từ sổ chi tiết cho cả 13 chứng từ.
- Dòng 309 và 388 cùng chuỗi 1C25TUV/00000040, đã về 0 sau các HĐ điều chỉnh 1C26MUV/00110703 và 00110704; AMIS còn 1.039.000 đ. Dòng 391 ghi `13150` tại BHH Amis, ghi chú hàng tặng ở chứng từ nghiệp vụ khác, chưa ghép được với bản xuất bán hàng.
- Đã đọc lại đủ 408 dòng, kiểm tra công thức, xác nhận A–T không đổi khi chuyển sang số cuối chuỗi và xem trực tiếp vùng cập nhật trên Google Sheets. Kết quả cuối, snapshot giá trị/công thức và biên bản kiểm tra nằm tại `outputs/map-hd-cty-amis-20261108/chain-current/`; dùng `final-verification.json` và `final-chain-results.json` làm kết quả hiện hành. Các kết quả đối chiếu từng hóa đơn ở thư mục cha là bước trước, đã được thay bằng đối chiếu theo chuỗi.
- Kiểm tra bản bàn giao trên máy khác bằng `python outputs/map-hd-cty-amis-20261108/chain-current/verify_saved_results.py`. Lệnh đọc các snapshot đã lưu và kiểm SHA-256 hai nguồn; không ghi lên Google Sheets.
