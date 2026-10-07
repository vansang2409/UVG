# UVG – thiết lập máy ở nhà và tiếp tục công việc

Cập nhật ngày 07/10/2026. Tài liệu này dành cho máy chưa có dự án. Repo: https://github.com/vansang2409/UVG.git, nhánh `master`.

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

Đây là dự án UVG chuyển từ máy công ty sang máy ở nhà. Kiểm tra lịch sử Git có commit 179d5a5 chứa workflow hiện tại. Thư mục làm việc trên máy này là thư mục project đang mở; đường dẫn D:\UVG trong tài liệu là đường dẫn máy công ty, cần đổi theo thư mục hiện tại.

Phạm vi hiện tại: hóa đơn gốc năm 2024 đã bị điều chỉnh, ngoài nhóm ngoại lệ. Không xử lý hóa đơn thay thế trong đợt này. Bảng ngoại lệ có 118 gốc / 190 hóa đơn; bảng đã điều chỉnh có 590 gốc / 1.187 hóa đơn.

Ngoại lệ là hóa đơn xuất cho công ty hoặc cá nhân có MST, giữ nguyên. Khi xuất mới từ AMIS sau này phải loại doanh thu đã nằm trong ngoại lệ để tránh xuất trùng. Nhận diện ngoại lệ hiện từ file nguồn, chưa xác thực MST bên ngoài.

Dữ liệu nguồn giữ nguyên; tạm dùng bộ hiện có, chưa cập nhật đến ngày hiện tại. Bảng tổng đến 30/06/2025; danh sách điều chỉnh có một ngoại lệ liên quan ngày 15/07/2025. Kiểm tra 11 file bằng SOURCE_MANIFEST.json trước khi tính lại.

Riêng 1C24TUV/00000574: người dùng đã đồng ý tạm tính thanh toán gốc trống = 0; thanh toán cộng dồn 1.839.000 đồng. Ô nguồn vẫn trống. Không áp dụng quy ước này cho mọi ô trống. Còn 5 chuỗi gốc 93, 228, 283, 285, 535 lệch số học cần xác minh.

Tiếp theo: đọc mẫu hóa đơn điều chỉnh của gốc số 3, trình bày cách tính và phương án đưa chuỗi về 0 để tôi xem. Mẫu chưa được phát hành. Chưa tạo file import, ký hoặc phát hành hóa đơn.

Mỗi nhóm chỉ giữ một file Excel làm việc. Sửa đúng file hiện dùng; nếu Excel khóa file thì yêu cầu tôi đóng, không tạo thêm bản. Giữ bố cục F0/F1/F2, hai hàng tiêu đề phân màu, không thêm lại tên người mua/MST hoặc cột loại/tham chiếu trên bảng chính. Thông tin chi tiết nằm ở tab nguồn và hướng dẫn ở tab Huong dan.
```

Lịch sử chat cũ không tự nằm trong bản clone. Các tài liệu bàn giao và lời nhắn trên cung cấp ngữ cảnh để Codex mới tiếp tục.

## 3. Hai file Excel đang dùng

- Ngoại lệ giữ nguyên: `outputs/ngoai-le-meinvoice-2024/NGOAI_LE_2024_THEO_CHUOI_GON.xlsx`.
- Nhóm đã điều chỉnh: `outputs/hoa-don-dieu-chinh-2024/HOA_DON_DA_DIEU_CHINH_2024_BO_SUNG_TEN.xlsx`.

Tên file thứ hai còn cụm `BO_SUNG_TEN` từ phiên bản trước; bảng chính hiện đã bỏ tên/MST theo yêu cầu. Không cần đổi tên hay tạo bản mới.

Nguồn nằm trong `data/Du lieu Amis/` và `data/Du lieu Meinvoice/`. Phần bàn giao cũ phía dưới TIEP_TUC_UVG.md chỉ dùng tra lịch sử; trạng thái ở đầu file được ưu tiên.

## 4. Khi cần chạy lại script

Chỉ đọc tài liệu và mở Excel thì chưa cần cài Python/Node. Khi cần tái tính, nhờ Codex kiểm tra runtime sẵn có và chuẩn bị thư viện theo README.md.

- Phân tích: Python với thư viện trong `requirements.txt`.
- Xuất Excel: Node với `@oai/artifact-tool`; thư viện không nằm trong Git. Codex cần tìm runtime đi kèm hoặc báo rõ nếu thiếu.

Các lệnh cho workflow đang làm, chạy từ thư mục project:

```powershell
python scripts/prepare_adjusted_2024.py
node scripts/layout-tools/build-adjusted.mjs
```

Đóng bảng Excel trước khi xuất lại. Không chạy `run.py` để tiếp tục nhóm này vì lệnh đó tái tạo các báo cáo đối chiếu cũ.

## 5. Chuyển qua lại giữa hai máy

Khi kết thúc trên một máy, yêu cầu Codex commit và push đúng các thay đổi cần giữ. Trên máy còn lại, pull trước khi làm tiếp. File chưa commit/push sẽ không tự xuất hiện trên máy kia.
