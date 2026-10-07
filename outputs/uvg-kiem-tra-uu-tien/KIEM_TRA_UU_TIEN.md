# UVG – danh sách kiểm tra ưu tiên

Có **321 hóa đơn riêng biệt** được chọn từ 10829 dòng hóa đơn trong báo cáo hiện có.

| Vấn đề | Số hóa đơn |
|---|---:|
| Thiếu liên kết tới hóa đơn nguồn | 299 |
| Thiếu số tiền trên bảng tổng | 14 |
| Tiền hàng cộng thuế không bằng tổng thanh toán | 8 |

Các nhóm có thể trùng nhau. Không cộng các số nhóm để ra số hóa đơn riêng biệt.

Đây là danh sách cần kiểm tra, chưa phải danh sách hóa đơn sai hoặc cần phát hành.
Các hóa đơn ngoài danh sách vẫn có thể thiếu chi tiết sản phẩm, thiếu lượng hoặc lệch kỳ.

## Cách dùng

1. Mở HOA_DON_CAN_KIEM_TRA_UU_TIEN.xlsx, lọc theo Nhóm kiểm tra ưu tiên.
2. Tra cứu bằng Ký hiệu + Số hóa đơn. File nguồn và Dòng giúp mở lại đúng dòng trong báo cáo nguồn.
3. Bổ sung tài liệu ở cột Tài liệu cần bổ sung. Giữ XML và tài liệu liên quan trong data/Bo sung.
4. Điền các cột từ Tình trạng kiểm tra trở đi. Ghi đường dẫn tương đối của bằng chứng để máy khác mở được.
5. Chỉ ghi hướng xử lý khi đã đủ căn cứ. Đề xuất không đồng nghĩa với được phép phát hành.

CSV giữ thông tin gốc dạng chữ, số hóa đơn có số 0 đầu. Nếu mở CSV trong Excel, dùng Data > From Text/CSV và chọn kiểu Text cho Ký hiệu, Số hóa đơn, Ký hiệu gốc, Số HĐ gốc.
Giá trị tiền bị thiếu được giữ trống. Không coi ô trống là 0.

## Tái tạo

```powershell
py scripts/prepare_priority_review.py
```

Lệnh trên ghi lại CSV và báo cáo thống kê từ kết quả đối chiếu hiện tại. Nó không đọc các ghi chú bạn đã nhập trong Excel.
Excel là bản làm việc đã tạo riêng. Lưu bản có ghi chú với tên khác trước khi nhận hoặc tạo bản mới.
