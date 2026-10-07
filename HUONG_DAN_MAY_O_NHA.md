# UVG – thiết lập máy ở nhà và tiếp tục công việc

Cập nhật ngày 08/10/2026. Tài liệu này dùng khi thiết lập hoặc tiếp tục dự án trên máy khác. Repo: https://github.com/vansang2409/UVG.git, nhánh `master`.

## 1. Lấy dự án về máy ở nhà

Cài Git và ứng dụng Codex. Trong PowerShell, chạy:

```powershell
git clone https://github.com/vansang2409/UVG.git "$HOME\Documents\UVG"
Set-Location "$HOME\Documents\UVG"
git log -3 --oneline
```

Nếu GitHub yêu cầu đăng nhập, dùng tài khoản có quyền truy cập repo. Repo đã chứa dữ liệu nguồn, script, hai bảng Excel đang dùng và tài liệu bàn giao. Không cần sao chép riêng dữ liệu từ máy công ty.

Nếu đã clone dự án, cập nhật trong thư mục dự án:

```powershell
git status --short
git pull --ff-only origin master
```

Nếu có thay đổi chưa lưu vào Git hoặc lệnh pull báo lỗi, nhờ Codex kiểm tra trước; không xóa hay reset thay đổi.

## 2. Cho Codex hiểu công việc

Mở thư mục `Documents\UVG` làm project trong Codex. Dán nguyên nội dung sau vào chat:

```text
Đọc AGENTS.md nếu có, HUONG_DAN_MAY_O_NHA.md, phần trạng thái hiện tại ở đầu TIEP_TUC_UVG.md và README.md.

Đây là dự án UVG chuyển giữa hai máy. Thư mục làm việc là project đang mở; đường dẫn D:\UVG trong tài liệu cần đổi theo thư mục hiện tại. Kiểm tra SOURCE_MANIFEST.json và trạng thái hiện tại ở đầu TIEP_TUC_UVG.md; không khôi phục nguồn cũ đã được thay bằng bộ xuất mới.

Phạm vi: F0 từ 01/01/2024 đến 30/06/2025; hóa đơn liên quan đến những F0 này đến 31/12/2025. Nhóm đang xử lý là đã điều chỉnh ngoài ngoại lệ; chưa xử lý thay thế. Bảng ngoại lệ có 238 F0 / 326 hóa đơn; bảng đã điều chỉnh có 1.047 F0 / 2.113 hóa đơn.

Ngoại lệ là hóa đơn xuất cho công ty hoặc cá nhân có MST, giữ nguyên. Khi xuất mới từ AMIS sau này phải loại doanh thu đã nằm trong ngoại lệ để tránh xuất trùng. Nhận diện ngoại lệ hiện từ file nguồn, chưa xác thực MST bên ngoài.

Dữ liệu nguồn giữ nguyên. Bộ meInvoice đã xuất lại đến 31/12/2025; ba file chi tiết khớp 48.874 hóa đơn trong 10 phần bảng tổng. Kiểm 18 file bằng SOURCE_MANIFEST.json. Script có thêm 8 liên kết lịch sử năm 2024 từ Git tại commit cố định 9c50b0424a242695fa31182ab826613bf0fa699f; provenance ở tab nguồn. Không tự tính các lần điều chỉnh/thay thế năm 2026 theo nhãn trạng thái hiện tại.

Riêng 1C24TUV/00000574: người dùng đã đồng ý tạm tính thanh toán gốc trống = 0; thanh toán cộng dồn 1.839.000 đồng. Ô nguồn vẫn trống. Không áp dụng quy ước này cho mọi ô trống. Report điều chỉnh có 16 chuỗi có ghi chú; ngoại lệ có 2. Đọc cột Cần kiểm tra và tab nguồn để xác minh tiền/ngày.

Phần đang dở: kiểm tra 8 hóa đơn lệch giữa Tổng tiền (R) và Doanh số bán chưa thuế (P) + Thuế GTGT (Q), có đủ ký hiệu/số và giá trị trong phần đầu TIEP_TUC_UVG.md; ưu tiên 1C25TUV/00000649 và 00000561. Ngày bảng chính của 1C25TUV/00012885 đã dùng 05/07/2025 theo xác nhận và ảnh người dùng, nguồn tổng vẫn ghi 18/08/2025; chưa đối chiếu XML. Bảy hóa đơn có Tổng tiền trống ở bảng tổng nhưng thanh toán 0 ghi rõ ở bảng chi tiết; report chưa dùng số 0 này để bổ sung cộng dồn. Sau khi xác minh ghi chú mới xem phương án đưa nhóm ngoài ngoại lệ về 0, trong đó có mẫu điều chỉnh của gốc số 3 chưa được phát hành. Chưa tạo file import, ký hoặc phát hành hóa đơn.

Mỗi nhóm chỉ giữ một Excel, trực tiếp ở outputs, không tạo thư mục con. Nếu Excel khóa file thì yêu cầu tôi đóng, không tạo thêm bản. Giữ format chuẩn hai hàng tiêu đề phân màu F0/F1/F2, mở F3 khi cần. Bảng điều chỉnh không thêm tên/MST hoặc loại/tham chiếu trên bảng chính; bảng ngoại lệ thêm đúng hai cột đầu Tên người mua và MST/CCCD chủ hộ nguyên bản. Chi tiết ở tab nguồn, hướng dẫn ở tab Huong dan. Chỉ đọc bàn giao trước, chưa tự tái tạo Excel hay tạo import, ký hoặc phát hành.
```

Lịch sử chat cũ không tự nằm trong bản clone. Các tài liệu bàn giao và lời nhắn trên cung cấp ngữ cảnh để Codex mới tiếp tục.

## 3. Hai file Excel đang dùng

- Ngoại lệ giữ nguyên: `outputs/NGOAI_LE_F0_2024_T6_2025.xlsx`.
- Nhóm đã điều chỉnh: `outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx`.

Đây là hai report đã làm lại theo bộ xuất mới. Output cũ đã dọn; chưa tạo import hoặc phát hành hóa đơn.

Nguồn nằm trong `data/Du lieu Amis/` và `data/Du lieu Meinvoice/`. Phần bàn giao cũ phía dưới TIEP_TUC_UVG.md chỉ dùng tra lịch sử; trạng thái ở đầu file được ưu tiên.

## 4. Khi cần chạy lại script

Chỉ đọc tài liệu và mở Excel thì chưa cần cài Python/Node. Khi cần tái tính, nhờ Codex kiểm tra runtime sẵn có và chuẩn bị thư viện theo README.md.

- Phân tích: Python với thư viện trong `requirements.txt`.
- Xuất Excel: Node với `@oai/artifact-tool`; thư viện không nằm trong Git. Codex cần tìm runtime đi kèm hoặc báo rõ nếu thiếu.

Các lệnh cho workflow đang làm, chạy từ thư mục project:

```powershell
python scripts/prepare_invoice_reports.py
node scripts/layout-tools/build-adjusted.mjs
```

Đóng bảng Excel trước khi xuất lại. Không chạy `run.py` để tiếp tục nhóm này vì lệnh đó tái tạo các báo cáo đối chiếu cũ.

## 5. Chuyển qua lại giữa hai máy

Khi kết thúc trên một máy, yêu cầu Codex commit và push đúng các thay đổi cần giữ. Trên máy còn lại, pull trước khi làm tiếp. File chưa commit/push sẽ không tự xuất hiện trên máy kia.
