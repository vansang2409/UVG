# UVG – tiếp tục ở nhà, cập nhật 11/10/2026

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

> Đọc TIEP_TUC_UVG.md từ đầu, README.md và HUONG_DAN_MAY_O_NHA.md; ưu tiên cập nhật 11/10/2026 hơn lịch sử. AMIS là mục tiêu, meInvoice là lịch sử; F0 thuộc 01/01/2024–30/06/2025, liên quan đến31/12/2025. Import hiện tại: chưa điều chỉnh 8.404 hóa đơn/13.543 dòng; đã điều chỉnh 986 chuỗi/1.990 hóa đơn dự kiến/41.268 dòng, gồm 195 chuỗi đã bổ sung theo XML. Đã bổ sung nguồn cho57F0 và rà thêm608/733 cùng ngoại lệ609; không còn thiếu chi tiết trong import. Đã sửa11đơn giá thuộc8F0 theoXML. Rà toàn bộ10.394hóa đơn/54.811dòng đạt kiểm tra nguồn/kế hoạch;986chuỗi về0. Còn2.185dòng/514F0 có SL×đơn giá khác thành tiền nguồn, phù hợp đơn giá làm tròn; giữ số tiền nguồn và kiểm tùy chọn tự tính khi xem trước meInvoice. Còn19chuỗi chờ,11chuỗi cầnchốt,6chuỗi thaythế; không nhầm19chuỗi chờ với19XMLgốc bổ sung. 237XMLđiềuchỉnh đã có và rà hash/tổng/thamchiếu, ngày12885 đã chốt18/08/2025. Chưa import/ký/phát hành hoặc xác minh mật mã chữ ký/CQT trực tiếp. Không chạy builder/verifier lịch sử ghi đè workbook; giữ số0đầu, ôtrống và dữliệunguồn. Dùng script chỉđọc hiện tại, tra runtime bằng load_workspace_dependencies; đọc phần việc còn lại và tiếp tục theo yêu cầu tôi.

## File chính

- outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx
- outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx
- outputs/NGOAI_LE_F0_2024_T6_2025.xlsx
- outputs/kiem-tra-20261008/CAC_TRUONG_HOP_CAN_KIEM_TRA.xlsx (bảng kiểm tra đã tạo trước khi chuyển sang sheet).

Chỉ mở file/tài liệu thì không cần tái tạo Excel. Script/dataJSON hiện hành được bàn giao; ảnhpreview/log và bảnxlsx trunggian là filetạm, không cần chuyển. Các verify cũ có thể phụ thuộc snapshot và ghi đè output: không chạy tùy tiện. Dùng kiểm tra chỉ đọc mới:

```powershell
python scripts/verify_handoff.py
python scripts/audit_supplement_57.py
python scripts/audit_current_import_plan.py
python scripts/audit_current_sources.py
```

Cần Python 3.11+ với openpyxl/lxml theo requirements.txt; có thể dùng runtime Codex từ load_workspace_dependencies. Các script trên tự xác định root repo, không cần Downloads máy cũ và không ghi Excel; chỉ ghi JSON kết quả kiểm tra. Giữ lịch sử Git để kiểm liên kết thay thế60→61 từ commit9c50b042. verify_handoff kiểm manifest18nguồn, sốlượngimport, số0đầu, tênKH=ngườimua, hàngkhuyếnmãi, loạiSL0 và sốhọc về0 của986chuỗi. Đây là kiểm tra dữ liệu, không xác nhận quy định hoặc chữ ký/CQT.

Trước sửa hoặc tạo file: xác minh phản hồiKT, phiên bảnworkbook và ngày phát hành dự kiến. Import cũ dự thảo ngày08/10/2026;390hóađơn bổ sung dự thảo09/10/2026. Không tự đổi thành ngày máy hiện tại.
