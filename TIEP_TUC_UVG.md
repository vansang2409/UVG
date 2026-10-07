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

Đã phân tích file và lập danh sách cần kiểm tra. Chưa kết nối CQT, chưa xác minh từng hóa đơn bằng XML, chưa chốt số lượng cần phát hành, chưa import, ký hay phát hành hóa đơn. Xem commit mới nhất trên GitHub để biết phiên bản tài liệu đã nhận.

Các file bàn giao đã được commit trong dự án. Ngày 07/10/2026, quyền chạy lệnh Git ngoài sandbox đã cho phép Codex hoàn tất bước commit. Trên máy ở nhà, nhận phiên bản mới bằng clone/pull như hướng dẫn ở trên, rồi kiểm tra:

```powershell
git log -1 --oneline
Test-Path TIEP_TUC_UVG.md
Test-Path outputs/uvg-kiem-tra-uu-tien/HOA_DON_CAN_KIEM_TRA_UU_TIEN.xlsx
```

Commit bàn giao có thông điệp `Add invoice priority review and home continuation guide`. Hai lệnh Test-Path cần trả về True. Commit dự án nền trước đó là `4e78390`.
