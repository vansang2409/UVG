# UVG – file làm việc ngoại lệ năm 2024

Mở **NGOAI_LE_2024_THEO_CHUOI_GON.xlsx**. Đây là bản đang dùng theo format người dùng chọn.

- Mỗi dòng là một hóa đơn gốc năm 2024: 118 dòng gốc, chứa đủ 190 hóa đơn trong nguồn hiện có.
- Tên người mua và MST ở đầu; F0 là gốc, F1/F2 là các hóa đơn sửa có liên kết rõ ràng. Các lần sửa sau năm 2024 vẫn giữ cùng dòng với gốc.
- Mỗi F gồm ký hiệu, số, ngày, loại sửa/hóa đơn tham chiếu, tiền hàng, thuế và thanh toán.
- Bảng Lien ket nguon chứa căn cứ file và dòng nguồn. MST mới được nhận diện cấu trúc, chưa tra cứu xác thực.
- Không cộng hóa đơn đã bị thay thế với bản thay thế để tính doanh thu; tiền thiếu giữ trống. Chưa xác minh lịch sử đầy đủ đến hiện tại.

Đã dọn bản Excel/CSV cũ và ảnh/file kiểm tra trung gian ngày 07/10/2026. Giữ workbook_data.json và chains_data.json làm dữ liệu trung gian cho script, không dùng làm bảng làm việc.

Dữ liệu nguồn trong data, Du lieu Amis và Du lieu Meinvoice giữ nguyên. Chưa điều chỉnh, thay thế, import, ký hoặc phát hành hóa đơn.
