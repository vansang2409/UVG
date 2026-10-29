# Kiểm tra sau phát hành ngày 29/10/2026

**Đã về 0 cả tiền hàng, VAT và thanh toán trong phạm vi KT thực hiện.** Ngày30/10/2026 người dùng xác nhận ô tổng thanh toán trống nghĩa là0. Áp dụng đúng24ô của15chuỗi đang treo (9điều chỉnh +6thay thế), giữ nguyên ô trống trong nguồn. Đủ6liên kết thay thế. **1C25TUV/00001943** được KT chủ động loại vì **TĐ không hợp lệ**, có ảnh và xác nhận người dùng; khoản nguồn645.051đ tách riêng, không tính là thiếu xử lý.

Đã đọc 40 file nguồn: 33 phần bảng kê đã sử dụng (159,589 hóa đơn duy nhất), 5 phần điều chỉnh (13,400 liên kết) và 2 file thay thế (468 liên kết). Bộ ban đầu chọn kỳ 01/01/2024–29/10/2026; file bổ sung xuất theo ngày HĐ thay thế 29/10/2026. Không trùng ký hiệu/số trên bảng kê, không phát hiện xung đột tiền giữa các nguồn đã so. Các phần đều có STT liên tục. Hash của các nguồn và kế hoạch giữ nguyên trước/sau kiểm tra.

| Bộ kiểm tra | Số gốc/chuỗi | Về 0 đủ tiền hàng/VAT/thanh toán | Trường hợp còn lại |
|---|---:|---:|---|
| Đã điều chỉnh: 2.045 HĐ theo kế hoạch | 1.015 | 1.015 | Không còn chênh |
| Chưa điều chỉnh: 8.410 HĐ trong file, KT loại 1 | 8.409 trong phạm vi | 8.409 | Không còn chênh |
| Thay thế 6 chuỗi | 6 | 6 | Khớp đủ 6 cặp; đủ ba khoản0 |

Tổng hai bộ điều chỉnh sau loại trừ: **9.424/9.424 gốc trong phạm vi** khớp0 đủ ba khoản; cộng6chuỗi thay thế thành **9.430/9.430 chuỗi đã về0**. Đủ **10.454/10.454 HĐ điều chỉnh trong phạm vi** theo liên kết và kế hoạch; STT1168/gốc1943 nằm ngoài đợt thực hiện theo xác nhận KT. Toàn bộ10.454 HĐ điều chỉnh năm2026 và6HĐ thay thế mới được bảng kê ghi **HĐ hợp lệ**. Đây là trạng thái trong file xuất, chưa tra cứuCQT trực tiếp hoặc kiểm mật mã chữ ký.

## 13 hóa đơn bổ sung trên 8 chuỗi

Đã tìm đủ13HĐ **1C26MUV/00110692–00110704**, ngày29/10/2026, đúng liên kết8gốc, đều ghi HĐ hợp lệ và khớp đủ3khoản với kế hoạch. HĐ00110696/STT2037 thanh toán nguồn trống được tính0 theo xác nhận30/10. Tổng tiền hàng bổ sung -233.262.118đ vàVAT -18.660.968đ đã triệt tiêu số còn dư trước đó của8chuỗi. Chưa kiểm chi tiết số lượng, đơn giá, mặt hàng của lần phát hành mới vì các file này là bảng kê tổng.

| STT kế hoạch | Gốc | HĐ phát hành | Kết quả so tiền |
|---:|---|---|---|
| 2033 | 1C24TUV/00000003 | 1C26MUV/00110692 | Khớp 3 khoản |
| 2034 | 1C24TUV/00000003 | 1C26MUV/00110693 | Khớp 3 khoản |
| 2035 | 1C24TUV/00000498 | 1C26MUV/00110694 | Khớp 3 khoản |
| 2036 | 1C24TUV/00000498 | 1C26MUV/00110695 | Khớp 3 khoản |
| 2037 | 1C24TUV/00000574 | 1C26MUV/00110696 | Khớp 3 khoản |
| 2038 | 1C24TUV/00000574 | 1C26MUV/00110697 | Khớp 3 khoản |
| 2039 | 1C25TUV/00001348 | 1C26MUV/00110698 | Khớp 3 khoản |
| 2040 | 1C25TUV/00002377 | 1C26MUV/00110699 | Khớp 3 khoản |
| 2041 | 1C25TUV/00002378 | 1C26MUV/00110700 | Khớp 3 khoản |
| 2042 | 1C24TUV/00000478 | 1C26MUV/00110701 | Khớp 3 khoản |
| 2043 | 1C24TUV/00000478 | 1C26MUV/00110702 | Khớp 3 khoản |
| 2044 | 1C25TUV/00000040 | 1C26MUV/00110703 | Khớp 3 khoản |
| 2045 | 1C25TUV/00000040 | 1C26MUV/00110704 | Khớp 3 khoản |

## Hóa đơn KT đã loại khỏi đợt xử lý

- Gốc **1C25TUV/00001943**, ngày 29/05/2025; STT1168 của bộ chưa điều chỉnh.
- Giá trị nguồn vẫn giữ 597.269đ tiền hàng + 47.782đ VAT =645.051đ; file import online và bản cục bộ có khoản giảm tương ứng.
- Người dùng xác nhận KT chủ động bỏ HĐ này vì TĐ không hợp lệ; ảnh mới cũng hiển thị trạng thái này. Việc không có HĐ điều chỉnh phát hành được giải thích bằng quyết định loại trừ, không còn ghi là khoản thiếu xử lý.
- Chỉ đổi phạm vi báo cáo. Giữ nguyên dữ liệu nguồn, file import và bản online; không kết luận thay cho việc xử lý thuế của hóa đơn không hợp lệ.

## 9 chuỗi điều chỉnh đã chốt thanh toán trống bằng0

Người dùng xác nhận trực tiếp ngày30/10/2026: “trống là 0 á”. Áp dụng vào phép đối chiếu, giữ nguyên giá trị trống trên file gốc. 9chuỗi dưới đã khớp0 cả ba khoản; không còn yêu cầu XML chỉ vì những ô thanh toán này.

| Gốc | HĐ điều chỉnh có thanh toán trống |
|---|---|
| 1C24TUV/00000574 | 1C26MUV/00110696 |
| 1C25TUV/00000006 | 1C26MUV/00101944 |
| 1C25TUV/00000016 | 1C26MUV/00101946 |
| 1C24TUV/00000671 | 1C26MUV/00102243 |
| 1C24TUV/00000744 | 1C26MUV/00102257 |
| 1C25TUV/00000391 | 1C26MUV/00102465 |
| 1C25TUV/00000761 | 1C26MUV/00102587 |
| 1C25TUV/00000835 | 1C26MUV/00102613 |
| 1C25TUV/00004748 | 1C26MUV/00105733 |

## Sáu hóa đơn thay thế — đã xác nhận liên kết ngày 30/10/2026

Các tham chiếu **1C25TUV/00002601, 00002610, 00002602, 00002603, 00002604, 00002617** trong bảng kê mới đều ghi **Hóa đơn đã bị thay thế**.

File bổ sung `Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_1793293386111.xls` có **6/6 cặp** đúng tham chiếu dự kiến. HĐ cũ ngày31/05/2025, HĐ mới ngày29/10/2026. Số tiền từng HĐ cũ khớp nguồn và tổng371.344.277đ tiền hàng +29.707.543đVAT =401.051.820đ thanh toán. HĐ mới tiền hàng/VAT bằng0; cộtW thanh toán trống được tính0 theo xác nhận người dùng30/10. Tất cả6HĐ mới được bảng kê tổng ghi **HĐ hợp lệ**.

| HĐ cũ (1C25TUV) | HĐ thay thế (1C26MUV) | Tiền hàng mới | VAT mới | Thanh toán mới |
|---|---|---:|---:|---|
| 00002601 | 00110705 | 0 | 0 | 0 theo xác nhận |
| 00002602 | 00110707 | 0 | 0 | 0 theo xác nhận |
| 00002603 | 00110708 | 0 | 0 | 0 theo xác nhận |
| 00002604 | 00110709 | 0 | 0 | 0 theo xác nhận |
| 00002610 | 00110706 | 0 | 0 | 0 theo xác nhận |
| 00002617 | 00110710 | 0 | 0 | 0 theo xác nhận |

**Kết luận:** đủ6liên kết thay thế và cả15chuỗi từng treo thanh toán đã chốt0 theo xác nhận. Không còn trường hợp chênh/thiếu liên kết trong phạm vi đợt đối soát này. Phạm vi kết luận là số tiền, số lượng hóa đơn và liên kết theo các nguồn đã nhận cùng xác nhận người dùng; chưa kiểm chữ ký hoặc toàn bộ chi tiết mặt hàng phát hành mới.

## Cách đối chiếu

- Khóa hóa đơn: ký hiệu + số hóa đơn, không ghép chỉ theo số.
- Lần theo đầy đủ liên kết; giữ gốc bị điều chỉnh và cộng các khoản điều chỉnh có liên quan.
- Loại giá trị bản bị thay thế/hủy; chỉ kết luận khi có đầy đủ bản thay thế và liên kết.
- Khớp số lượng HĐ và từng bộ tiền với kế hoạch, không chỉ tổng ròng.
- 24ô thanh toán trống trên15chuỗi được tính0 theo xác nhận30/10; giữ ô nguồn trống, lưu căn cứ riêng. Các ô trống khác không thuộc xác nhận vẫn cần chứng cứ. Ngoại lệ XML lịch sử không tính lại thanh toán từ tiền hàng + VAT.
- Đọc dữ liệu cục bộ. Không nhập/phát hành lại, sửa meInvoice, sửa AMIS hoặc chỉnh workbook nguồn.

## Tệp chi tiết

- `CAC_TRUONG_HOP_CAN_KIEM_TRA.csv`: không còn dòng cần xử lý (chỉ còn tiêu đề).
- `LOAI_TRU_THEO_XAC_NHAN_KT.csv`: 1 gốc1943, giữ giá trị nguồn và lý do KT loại, không cộng vào số dư trong phạm vi.
- `TAT_CA_CHUOI.csv`: đủ9.431dòng, gồm9.430chuỗi trong phạm vi về0 và1gốc1943đã loại. Số dư của gốc đã loại để trống vì ngoài phạm vi.
- JSON căn cứ, nguồn/dòng, liên kết và SHA-256: `scripts/data/post-issue-20261029/audit.json`.

## Nguồn mới đã đọc

| Tệp | Số dòng hóa đơn/liên kết |
|---|---:|
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_1.xls | 4988 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_10.xls | 4781 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_11.xls | 4840 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_12.xls | 4891 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_13.xls | 4726 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_14.xls | 4765 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_15.xls | 4810 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_16.xls | 4644 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_17.xls | 4957 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_18.xls | 4663 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_19.xls | 4983 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_2.xls | 4983 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_20.xls | 4474 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_21.xls | 4415 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_22.xls | 4544 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_23.xls | 4957 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_24.xls | 4768 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_25.xls | 3972 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_26.xls | 4005 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_27.xls | 4674 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_28.xls | 4912 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_29.xls | 4517 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_3.xls | 4963 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_30.xls | 4977 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_31.xls | 4809 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_32.xls | 1237 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_33.xls | 10838 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_4.xls | 4862 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_5.xls | 4994 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_6.xls | 4934 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_7.xls | 4863 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_8.xls | 4919 |
| data/Meinv_SAU_PHAT_HANH_20261029/Bảng kê hóa đơn đã sử dụng/Bang_ke_hoa_don_da_su_dung_1793291495460_9.xls | 4924 |
| data/Meinv_SAU_PHAT_HANH_20261029/DIEU_CHINH/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_bi_dieu_chinh_1793291608325_1.xls | 4371 |
| data/Meinv_SAU_PHAT_HANH_20261029/DIEU_CHINH/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_bi_dieu_chinh_1793291608325_2.xls | 2505 |
| data/Meinv_SAU_PHAT_HANH_20261029/DIEU_CHINH/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_bi_dieu_chinh_1793291608325_3.xls | 2497 |
| data/Meinv_SAU_PHAT_HANH_20261029/DIEU_CHINH/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_bi_dieu_chinh_1793291608325_4.xls | 2307 |
| data/Meinv_SAU_PHAT_HANH_20261029/DIEU_CHINH/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_bi_dieu_chinh_1793291608325_5.xls | 1720 |
| data/Meinv_SAU_PHAT_HANH_20261029/THAY_THE/Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_1793293386111.xls | 6 |
| data/Meinv_SAU_PHAT_HANH_20261029/THAY_THE/Thay_the.xls | 462 |
