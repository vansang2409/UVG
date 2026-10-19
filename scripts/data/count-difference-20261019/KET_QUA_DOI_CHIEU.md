# Đối chiếu chênh số lượng theo ảnh kế toán — 19/10/2026

Chỉ đọc nguồn và workbook, không sửa dữ liệu nguồn/import. Đếm theo khóa ký hiệu + số hóa đơn, tránh trừ trùng nhóm.

## Cầu nối theo nguồn hiện có

| Nhóm | Số hóa đơn |
|---|---:|
| Bảng tổng có ngày 01/01/2024–30/06/2025 | 10.855 |
| Cộng hóa đơn liên quan sau 30/06/2025 | 321 |
| Tổng khóa nguồn xem xét | 11.176 |
| Loại điều chỉnh cho gốc năm 2022–2023 (37 gốc năm2022,250 năm2023; số lượng ở đây là số điều chỉnh) | -287 |
| Loại hóa đơn hủy theo trạng thái nguồn | -10 |
| Giữ nguyên nhóm ngoại lệ theo workbook hiện tại | -338 |
| Loại hóa đơn thay thế/bị thay thế chưa tính trong ngoại lệ | -30 |
| Giữ31chuỗi đã khớp, gồm31gốc+31điều chỉnh | -62 |
| Giữ3chuỗi cùng sản phẩm theo chốt kế toán, gồm6hóađơn | -6 |
| Giữ điều chỉnh tên00009186, chỉ xử lý gốc00009072 | -1 |
| Còn lại, khớp số hóa đơn mới trong2sheetimport | **10.442** |

Các tập loại trừ không giao nhau. Số cuối khớp8.410+2.032; đây là kiểm đếm số lượng, không thay thế kiểm số tiền/nội dung hoặc xác nhận phần mềm đã nhập thành công.

## Nhóm công ty325 và ngoại lệ338

- Có325hóađơn thuộc tiêu chí tên công ty/định danh MST trong kỳ01/01/2024–30/06/2025.
- Thêm1hóađơn liên quan sau kỳ:1C25TUV/00012764,15/07/2025,CÔNG TY CỔ PHẦN VGENTECH; bảng tổng phần3,dòng4385.
- Thêm12hóađơn thuộc6cặp CC01–CC06 được quyết định giữ riêng, dù tên gốc là khách lẻ.
- 325+1+12=338. Không gọi cả338là hóa đơn công ty.
- 33hóađơn thay thế/bị thay thế có3hóađơn đã nằm trong nhóm ngoại lệ:1C24TUV/00000060,1C24TUV/00000061,1C25TUV/00002605. Nếu đã trừ338ngoại lệ thì chỉ trừ thêm30, không trừ33lần nữa.

## Giới hạn khi đối chiếu ảnh kế toán

Ảnh tính10.855−2+321−10−4−2−325−33=10.800;10.800−10.442=358.

Chưa có số định danh của2hóađơn “hủy nội bộ” và6hóađơn điều chỉnh tên/địa chỉ trong bảng KT. Không tự trừ thêm2hoặc6khỏi cầu nối theo nguồn: cần danh sách để xác định chúng đã nằm trong10hủy/338ngoại lệ/9186hay nhóm nào khác. Không kết luận thiếu358hóađơn trong import từ ảnh tổng số. Muốn chốt chênh từng dòng KT, cần fileExcel/danh sách định danh mà KT dùng.

Căn cứ đầy đủ, gồm danh sách287điều chỉnh ngoài kỳ gốc,338ngoại lệ,các nhóm loại trừ và vị trí nguồn: `initial-audit.json` cùng thư mục. Script: `scripts/audit_count_difference_20261019.py`.
