# UVG – trạng thái hiện tại ngày 07/10/2026

Ưu tiên xử lý hóa đơn gốc năm 2024. Dữ liệu nguồn giữ nguyên, 11 file khớp SHA-256. Tạm dùng bộ nguồn hiện có; chưa cập nhật đến hiện tại. Bảng tổng đến 30/06/2025, DS có ít nhất một ngoại lệ điều chỉnh ngày 15/07/2025.

## Bảng đang dùng

1. `outputs/ngoai-le-meinvoice-2024/NGOAI_LE_2024_THEO_CHUOI_GON.xlsx`: 118 gốc, 190 hóa đơn trong chuỗi; nhóm công ty/cá nhân có MST giữ nguyên. Đây là nhận diện từ file, chưa xác thực MST. Có lịch sử sửa năm 2025 liên quan gốc 2024.
2. `outputs/hoa-don-dieu-chinh-2024/HOA_DON_DA_DIEU_CHINH_2024_BO_SUNG_TEN.xlsx`: 590 gốc đã bị điều chỉnh ngoài ngoại lệ, 1.187 hóa đơn. Bố cục 2 hàng tiêu đề phân màu F0/F1/F2; bỏ tên/MST và cột tham chiếu trên bảng chính, chi tiết vẫn ở tab nguồn. Không xử lý chuỗi thay thế.

## Quyết định người dùng và phần còn lại

- Mục tiêu người dùng: đưa nhóm ngoài ngoại lệ về 0, sau đó xuất mới từ AMIS; loại doanh thu đã nằm trong ngoại lệ giữ nguyên để tránh trùng. Chưa chốt nghiệp vụ phát hành.
- Chỉ xử lý nhóm đã điều chỉnh trong bước hiện tại; chưa xử lý thay thế hoặc nhóm chưa sửa.
- 5 chuỗi gốc 93, 228, 283, 285, 535 có lệch số học cần xác minh.
- Riêng 1C24TUV/00000574: người dùng xác nhận tạm tính thanh toán gốc trống = 0. Cộng dồn 1.839.000 đồng; giữ ô nguồn trống và ghi chú quy ước. Không áp dụng mọi ô trống.
- Đã đối chiếu mẫu hóa đơn số 3: cộng dồn 11.856.288 đồng; hai hóa đơn trong file mẫu cộng -11.856.288 đồng. Mẫu không được tính là đã phát hành.
- Bước tiếp: làm mẫu chuỗi số 3, tính phương án đưa về 0 cho nhóm còn lại; chưa tạo file import, ký hoặc phát hành.
- Mỗi nhóm giữ một file Excel đang dùng. Đóng Excel trước khi lưu đè; không tự tạo nhiều bản mới. Báo cáo cũ đã xóa theo yêu cầu, có thể tra lịch sử Git.

Tái tính: `python scripts/prepare_adjusted_2024.py`. Xuất lại đúng bố cục: `node scripts/layout-tools/build-adjusted.mjs` với Node và artifact-tool từ runtime Codex. Xem README.md.

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
