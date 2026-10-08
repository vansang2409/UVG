# UVG – tiếp tục ở nhà, cập nhật 08/10/2026

Repo: https://github.com/vansang2409/UVG.git – nhánh master.

## Nhận bản mới

Nếu đã clone, chạy trong thư mục UVG:

```powershell
git status --short
git pull --ff-only origin master
git log -1 --oneline
```

Nếu có thay đổi cục bộ, nhờ Codex kiểm tra và bảo toàn trước khi pull; không reset/xóa thay đổi. Nếu chưa clone:

```powershell
git clone https://github.com/vansang2409/UVG.git
cd UVG
```

## Cho Codex tiếp tục đúng ngữ cảnh

Dán đoạn này vào chat trong project UVG:

> Đọc TIEP_TUC_UVG.md từ đầu, README.md và HUONG_DAN_MAY_O_NHA.md. Đây là bàn giao toàn dự án, không chỉ 00000508. Ưu tiên trạng thái mới nhất hơn lịch sử. Kiểm tra nguồn theo SOURCE_MANIFEST.json và mở workbook hiện tại; không tái chạy builder cũ ghi đè ghi chú/sheet import. AMIS là mục tiêu, meInvoice là lịch sử; F0 thuộc 01/01/2024–30/06/2025, liên quan đến31/12/2025. Nhóm ngoại lệ giữ riêng; thay thế chưa xử lý. Người dùng đã gửi kế toán kiểm tra, chờ phản hồi về237chuỗi bằng0 và19chuỗi đang chờ. Hai sheet import đã chuẩn bị: nhóm chưa điều chỉnh8404hóa đơn/13543dòng (6hóa đơn sốlượng0 táchriêng); nhóm đã điều chỉnh791chuỗi/1600hóa đơn dựkiến/40777dòng. Với00000508 phải kiểm cả dòng chi tiết, tính chất dòng và tổng trongXML, không chỉ số học về0. Chưa import/ký/phát hành. Giữ nguồn, ôtrống và số0đầu. Đường dẫnD:\UVG trong script cũ cần đổi theo project máy này, tra runtime bằng load_workspace_dependencies. Đọc phần việc còn lại rồi tiếp tục theo yêu cầu tôi.

## File chính

- outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx
- outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx
- outputs/NGOAI_LE_F0_2024_T6_2025.xlsx
- outputs/kiem-tra-20261008/CAC_TRUONG_HOP_CAN_KIEM_TRA.xlsx (bảng kiểm tra đã tạo trước khi chuyển sang sheet).

Chỉ mở file/tài liệu thì không cần tái tạo Excel. Script/dataJSON hiện hành được bàn giao; ảnhpreview/log và bảnxlsx trunggian là filetạm, không cần chuyển. Các verify cũ có thể phụ thuộc snapshot và ghi đè output: không chạy tùy tiện. Dùng kiểm tra chỉ đọc mới:

```powershell
python scripts/verify_handoff.py
```

Cần Python với openpyxl; có thể dùng runtime Codex từ load_workspace_dependencies. Script mới tự xác định root repo, kiểm manifest18nguồn, sốlượngimport, số0đầu, tênKH=ngườimua, hàngkhuyếnmãi, loạiSL0 và sốhọc về0 nhóm791chuỗi. Đây là kiểm tra dữ liệu, không xác nhận quy định hay nội dungXML.

Trước sửa hoặc tạo file: xác minh phản hồiKT, phiênbảnworkbook và ngày phát hành dự kiến. Ngày08/10/2026 là ngày trong bản dự thảo hiện tại.