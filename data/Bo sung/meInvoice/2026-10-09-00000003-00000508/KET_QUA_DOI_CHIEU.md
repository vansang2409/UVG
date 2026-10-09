# Đối chiếu XML gốc00000003 và điều chỉnh00000508 – 09/10/2026

Đã đọc hai ZIP do người dùng cung cấp trong Downloads, lưu bản sao ZIP và XML nguyên bytes tại thư mục này. DOI_CHIEU_XML.json chứa trường trích xuất và SHA256. Không sửa XML, không kiểm chứng mật mã chữ ký, không thao tác meInvoice.

## Kết quả xác nhận từ XML và HTML đi kèm ZIP508

- Gốc1C24TUV/00000003 ngày10/01/2024:7dòng, đềuTChat1, sốlượng1; cộng thành tiền11334130, thuế906730, thanh toán12240860. Tổng chi tiết khớp phần tổng.
- Điều chỉnh1C25TUV/00000508 ngày26/04/2025: tham chiếu đúng mẫu1/kýhiệuC24TUV/số00000003/ngày10/01/2024, tínhchấthóađơnđiềuchỉnhTCHDon2.
- Có1dòng: “Điều chỉnh giảm thành tiền hoá đơn số00000003”; TChat1 (Hàng hóa,dịch vụ), không phải4 (Ghi chú/diễn giải); sốlượng0, đơngiá0, ThTien0, thuếsuất8%.
- Thuế dòng trong TTKhac/VATAmount vàVATAmountOC đều−28487. XML không có trường tổng thanh toán dòng chuẩn ghi−28487; số này trongExcelchi tiết phù hợp cách lấy tiềnhàng0+thuếdòng−28487.
- TToan/THTTLTSuat/LTSuat:ThTien−356085,TThue−28487. TToan/TgTCThue−356085,TgTThue−28487,TgTTTBSo−384572. Tổng đúng phép cộng.
- HTML trongZIP cũng hiển thị dòngTínhchấtHànghóa,dịchvụ, Thànhtiền0; phần tổng−356085/−28487/−384572. Như vậy chênh lệch có ngay trongXML/HTML đã cung cấp, không chỉ bảngExcel xuất.
- Gốc+508 theo phần tổng:10978045tiềnhàng,878243thuế,11856288thanhtoán. Không tự giảm khoản thuế−28487 lần nữa hoặc coi tổng508chỉ−28487.

## Đánh giá và bước còn lại

Đã xác nhận một điểm chưa nhất quán giữa dòngHHDV(TChat1,tiềnhàng0) và phần tổng tiềnhàngâm. MISA hiện hướng dẫn điều chỉnh thành tiền toàn hóa đơn dùngGhi chú/diễngiải; đây là căn cứ hỗ trợ kiểmtra cách lập, chưa phải kếtluận đầy đủ về pháp lý hóa đơn lập26/04/2025. CầnKT/MISA xác định cách sửa cả nội dung/tínhchất/dòng thành tiền, không chỉ làm tổngchuỗi0. Kế hoạch đảoF1 rồi giảmchi tiếtF0 vẫn làdựthảo, chưa tự phát hành hoặc chuyểnchuỗi này khỏi nhómchờ.

Không suy ra195hóađơn khác mắc lỗiXML giống508 khi chưa đọcXML của chúng. Chưa cóXML12885 đểchốt vụlệchngày05/07/2025–18/08/2025.

Nguồn kỹ thuật:PhụlụcIV,QĐ1450/QĐ-TCT:TChat1=Hànghóa,dịchvụ;4=Ghi chú/diễngiải.
https://luatvietnam.vn/thue/quyet-dinh-1450-qd-tct-bo-tai-chinh-210675-d1.html
Hướng dẫnMISA vềđiềuchỉnhgiảmthànhtiền:
https://helpv4.meinvoice.vn/kb/cach-viet-thong-tin-tren-hoa-don-dieu-chinh-giam-thanh-tien-khong-dieu-chinh-so-luong-don-gia/