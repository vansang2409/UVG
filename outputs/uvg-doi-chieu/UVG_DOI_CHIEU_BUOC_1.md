# Đối chiếu UVG – bước 1

Ngày phân tích: 06/10/2026. Chỉ đọc dữ liệu nguồn; chưa sửa hoặc phát hành hóa đơn.

AMIS được dùng làm số liệu mục tiêu đã được khách hàng chốt theo xác nhận của người dùng. meInvoice là lịch sử phát hành trong các file cung cấp. Chưa có dữ liệu CQT hoặc mã đơn sàn.

## Phạm vi và kiểm tra dữ liệu

- AMIS: 79,338 dòng, 65,042 số chứng từ tính riêng từng tháng. Không đồng nghĩa số đơn hàng hoặc hóa đơn cần phát hành.
- meInvoice: 10,853 hóa đơn trong bảng tổng; 10,800 khóa hóa đơn trong bảng chi tiết.
- Thiếu chi tiết so với bảng tổng: 55; thiếu bảng tổng so với chi tiết: 2.

## Tổng hợp theo tháng

AMIS phân kỳ theo ngày chứng từ. meInvoice phân kỳ theo ngày hóa đơn, bao gồm lần sửa hóa đơn kỳ cũ. Hai cột không cùng cơ sở thời gian; chênh lệch không phải số tiền cần xuất bổ sung. Tổng thanh toán meInvoice quy đổi dưới đây loại hóa đơn có trạng thái đã bị thay thế/đã hủy, giữ hóa đơn bị điều chỉnh và cộng các hóa đơn điều chỉnh hiện có. Đây là phép tính theo file, chưa xác minh XML, dữ liệu CQT hoặc chuỗi sửa đầy đủ.

| Tháng | Chứng từ AMIS | Tổng thanh toán AMIS | Hóa đơn meInvoice | Tổng thanh toán meInvoice quy đổi | HĐ thiếu giá trị tổng |
|---|---:|---:|---:|---:|---:|
| 2024-01 | 2,196 | 3.968.742.868 | 30 | 907.280.313 | 0 |
| 2024-02 | 1,023 | 1.940.417.064 | 45 | 1.799.106.331 | 1 |
| 2024-03 | 1,609 | 2.754.233.589 | 56 | 1.657.595.309 | 0 |
| 2024-04 | 3,056 | 5.278.284.182 | 86 | 5.139.623.145 | 0 |
| 2024-05 | 2,504 | 3.624.735.231 | 49 | 1.881.441.370 | 0 |
| 2024-06 | 3,070 | 4.450.947.091 | 71 | 6.046.067.643 | 0 |
| 2024-07 | 3,310 | 5.072.720.381 | 28 | 1.648.254.843 | 0 |
| 2024-08 | 3,456 | 5.256.787.164 | 72 | 6.760.398.095 | 0 |
| 2024-09 | 3,529 | 5.399.778.133 | 62 | 4.975.754.344 | 0 |
| 2024-10 | 4,075 | 5.506.787.240 | 58 | 3.303.686.102 | 0 |
| 2024-11 | 3,614 | 5.246.303.972 | 76 | 4.226.106.025 | 1 |
| 2024-12 | 3,572 | 5.424.441.483 | 153 | 8.097.170.163 | 2 |
| 2025-01 | 3,292 | 5.036.161.740 | 66 | 1.263.930.938 | 2 |
| 2025-02 | 5,353 | 7.244.115.869 | 99 | 4.086.057.219 | 0 |
| 2025-03 | 5,150 | 7.330.548.837 | 240 | 7.416.793.150 | 1 |
| 2025-04 | 4,378 | 5.988.519.553 | 127 | 15.793.420.290 | 0 |
| 2025-05 | 5,366 | 6.616.051.535 | 2,209 | 11.724.782.647 | 2 |
| 2025-06 | 6,489 | 7.054.437.812 | 7,326 | 8.866.559.897 | 5 |

Đơn vị tiền: VND. Dòng thiếu tổng thanh toán không được tự coi là hóa đơn 0 đồng; cần kiểm tra chi tiết/XML.

## Kiểm tra tổng dòng AMIS với dòng Tổng cộng

| File | Lệch doanh số | Lệch thuế | Lệch thanh toán |
|---|---:|---:|---:|
| So_chi_tiet_ban_hang 2024.xlsx | 0 | 0 | 0 |
| So_chi_tiet_ban_hang T1-6.2025.xlsx | 0 | 0 | 0 |

## Hóa đơn tổng tiền âm sau các điều chỉnh trong bảng liên kết

Đây là danh sách cần xác minh, chưa kết luận sai phạm. Số dư = giá trị gốc tại bảng liên kết + các khoản điều chỉnh liệt kê.

| Ký hiệu | Số hóa đơn | Giá trị gốc | Tổng điều chỉnh | Số dư | Nguồn và dòng |
|---|---|---:|---:|---:|---|
| 1C25TUV | 00000003 | 518.400 | -557.799 | -39.399 | DS bị điều chỉnh 2025.xls, dòng 8 |
| 1C25TUV | 00000009 | 1.379.000 | -1.720.000 | -341.000 | DS bị điều chỉnh 2025.xls, dòng 12 |
| 1C25TUV | 00000013 | 3.678.000 | -5.178.001 | -1.500.001 | DS bị điều chỉnh 2025.xls, dòng 13 |
| 1C25TUV | 00000386 | 17.277.991 | -19.983.506 | -2.705.515 | DS bị điều chỉnh 2025.xls, dòng 158 |
| 1C25TUV | 00001112 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 293 |
| 1C25TUV | 00001113 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 294 |
| 1C25TUV | 00001114 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 295 |
| 1C25TUV | 00001115 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 296 |
| 1C25TUV | 00001116 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 297 |
| 1C25TUV | 00001117 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 298 |
| 1C25TUV | 00001121 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 299 |
| 1C25TUV | 00001122 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 300 |
| 1C25TUV | 00001123 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 301 |
| 1C25TUV | 00001124 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 302 |
| 1C25TUV | 00001125 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 303 |
| 1C25TUV | 00001126 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 304 |
| 1C25TUV | 00001127 | 645.051 | -645.052 | -1 | DS bị điều chỉnh 2025.xls, dòng 305 |
| 1C25TUV | 00001328 | 438.999 | -439.000 | -1 | DS bị điều chỉnh 2025.xls, dòng 352 |
| 1C25TUV | 00001740 | 2.355.050 | -2.639.098 | -284.048 | DS bị điều chỉnh 2025.xls, dòng 372 |
| 1C25TUV | 00001742 | 1.747.050 | -2.355.050 | -608.000 | DS bị điều chỉnh 2025.xls, dòng 374 |
| 1C25TUV | 00002884 | 2.264.807 | -2.264.808 | -1 | DS bị điều chỉnh 2025.xls, dòng 385 |
| 1C25TUV | 00007999 | 569.000 | -679.000 | -110.000 | DS bị điều chỉnh 2025.xls, dòng 471 |

## Hóa đơn điều chỉnh nhiều lần

- 1C24TUV / 00000026: 1C25TUV / 00000512, 1C25TUV / 00002316
- 1C24TUV / 00000034: 1C25TUV / 00002608, 1C25TUV / 00002643
- 1C24TUV / 00000064: 1C25TUV / 00002609, 1C25TUV / 00009375
- 1C24TUV / 00000077: 1C24TUV / 00000186, 1C25TUV / 00002599
- 1C24TUV / 00000122: 1C25TUV / 00002611, 1C25TUV / 00009436
- 1C24TUV / 00000129: 1C25TUV / 00002612, 1C25TUV / 00002645
- 1C24TUV / 00000228: 1C25TUV / 00002619, 1C25TUV / 00002646
- 1C24TUV / 00000353: 1C25TUV / 00002031, 1C25TUV / 00002569
- 1C24TUV / 00000544: 1C25TUV / 00002162, 1C25TUV / 00002163
- 1C24TUV / 00000659: 1C25TUV / 00002639, 1C25TUV / 00012764
- 1C25TUV / 00000380: 1C25TUV / 00000696, 1C25TUV / 00000700
- 1C25TUV / 00000455: 1C25TUV / 00012911, 1C25TUV / 00012914, 1C25TUV / 00012933
- 1C25TUV / 00000481: 1C25TUV / 00012915, 1C25TUV / 00012928
- 1C25TUV / 00000485: 1C25TUV / 00012829, 1C25TUV / 00013042
- 1C25TUV / 00000720: 1C25TUV / 00009456, 1C25TUV / 00013052
- 1C25TUV / 00000748: 1C25TUV / 00009457, 1C25TUV / 00013036
- 1C25TUV / 00000755: 1C25TUV / 00012830, 1C25TUV / 00013060
- 1C25TUV / 00000766: 1C25TUV / 00012957, 1C25TUV / 00013026, 1C25TUV / 00013030
- 1C25TUV / 00000767: 1C25TUV / 00012831, 1C25TUV / 00013061
- 1C25TUV / 00000771: 1C25TUV / 00013037, 1C25TUV / 00013062

## Chuỗi thay thế tiếp tục được sửa

- 1C24TUV / 00000060 → 1C24TUV / 00000061 → Thay thế: 1C25TUV / 00002605
- 1C24TUV / 00000080 → 1C24TUV / 00000188 → Thay thế: 1C25TUV / 00002601
- 1C24TUV / 00000087 → 1C24TUV / 00000189 → Thay thế: 1C25TUV / 00002610
- 1C24TUV / 00000116 → 1C24TUV / 00000190 → Thay thế: 1C25TUV / 00002602
- 1C24TUV / 00000117 → 1C24TUV / 00000191 → Thay thế: 1C25TUV / 00002603
- 1C24TUV / 00000118 → 1C24TUV / 00000192 → Thay thế: 1C25TUV / 00002604
- 1C24TUV / 00000184 → 1C24TUV / 00000216 → Thay thế: 1C25TUV / 00002617

## Khóa hóa đơn chưa khớp giữa các bảng

Có trong tổng, không có trong chi tiết:

- 1C24TUV / 00000594 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 599)
- 1C24TUV / 00000595 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 600)
- 1C24TUV / 00000596 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 601)
- 1C24TUV / 00000597 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 602)
- 1C24TUV / 00000598 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 603)
- 1C24TUV / 00000599 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 604)
- 1C24TUV / 00000600 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 605)
- 1C24TUV / 00000601 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 606)
- 1C24TUV / 00000602 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 607)
- 1C24TUV / 00000603 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 608)
- 1C24TUV / 00000604 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 609)
- 1C24TUV / 00000605 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 610)
- 1C24TUV / 00000606 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 611)
- 1C24TUV / 00000607 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 612)
- 1C24TUV / 00000608 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 613)
- 1C24TUV / 00000609 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 614)
- 1C24TUV / 00000610 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 615)
- 1C24TUV / 00000611 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 616)
- 1C24TUV / 00000612 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 617)
- 1C24TUV / 00000613 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 618)
- 1C24TUV / 00000614 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 619)
- 1C24TUV / 00000615 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 620)
- 1C24TUV / 00000616 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 621)
- 1C24TUV / 00000617 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 622)
- 1C24TUV / 00000618 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 623)
- 1C24TUV / 00000619 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 624)
- 1C24TUV / 00000620 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 625)
- 1C24TUV / 00000621 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 626)
- 1C24TUV / 00000622 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 627)
- 1C24TUV / 00000623 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 628)
- 1C24TUV / 00000624 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 629)
- 1C24TUV / 00000625 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 630)
- 1C24TUV / 00000626 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 631)
- 1C24TUV / 00000627 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 632)
- 1C24TUV / 00000628 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 633)
- 1C24TUV / 00000629 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 634)
- 1C24TUV / 00000630 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 635)
- 1C24TUV / 00000631 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 636)
- 1C24TUV / 00000632 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 637)
- 1C24TUV / 00000633 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 638)
- 1C24TUV / 00000634 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 639)
- 1C24TUV / 00000635 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 640)
- 1C24TUV / 00000636 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 641)
- 1C24TUV / 00000637 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 642)
- 1C24TUV / 00000638 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 643)
- 1C24TUV / 00000639 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 644)
- 1C24TUV / 00000640 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 645)
- 1C24TUV / 00000641 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 646)
- 1C24TUV / 00000642 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 647)
- 1C24TUV / 00000643 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 648)
- 1C25TUV / 00002140 (Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls, dòng 2929)
- 1C25TUV / 00004203 (Bang_ke_hoa_don_da_su_dung_T6.2025.xls, dòng 6)
- 1C25TUV / 00004204 (Bang_ke_hoa_don_da_su_dung_T6.2025.xls, dòng 7)
- 1C25TUV / 00006108 (Bang_ke_hoa_don_da_su_dung_T6.2025.xls, dòng 1911)
- 1C25TUV / 00009426 (Bang_ke_hoa_don_da_su_dung_T6.2025 L2.xls, dòng 246)

Có trong chi tiết, không có trong tổng:

- 1C25TUV / 00001943
- 1C25TUV / 00002060

## Công việc tiếp theo

1. Giải thích các dòng thiếu giá trị tổng và khóa hóa đơn chưa khớp trước khi dùng tổng meInvoice.
2. Lập bảng ánh xạ sản phẩm AMIS và meInvoice, ghép có điều kiện theo số lượng, tiền hàng, thuế và thời gian. Không coi ghép bằng số tiền đơn thuần là bằng chứng chắc chắn.
3. Các giao dịch có nhiều ứng viên ghép hoặc không ghép được giữ ở nhóm chưa xác định; không tự chuyển sang xuất mới.
4. Bổ sung dữ liệu CQT và lịch sử sửa sau tháng 6/2025 trước khi chốt phương án phát hành.

## Nguồn

- D:\UVG\Du lieu Amis\So_chi_tiet_ban_hang 2024.xlsx
- D:\UVG\Du lieu Amis\So_chi_tiet_ban_hang T1-6.2025.xlsx
- D:\UVG\Du lieu Meinvoice\Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls
- D:\UVG\Du lieu Meinvoice\Bang_ke_hoa_don_da_su_dung_T6.2025 L2.xls
- D:\UVG\Du lieu Meinvoice\Bang_ke_hoa_don_da_su_dung_T6.2025.xls
- D:\UVG\Du lieu Meinvoice\Bảng kê chi tiết hóa đơn 2024.xls
- D:\UVG\Du lieu Meinvoice\Bảng kê chi tiết hóa đơn 2025.xls
- D:\UVG\Du lieu Meinvoice\DS bị thay thế 2024.xls
- D:\UVG\Du lieu Meinvoice\DS bị điều chỉnh 2024.xls
- D:\UVG\Du lieu Meinvoice\DS bị điều chỉnh 2025.xls