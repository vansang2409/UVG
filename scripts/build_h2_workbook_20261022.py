"""Ghi file cuối nhẹ hơn sau khi Artifact Tool không hoàn tất xuất workbook lớn.

Bố cục và công thức đã được tạo, tính và xem ảnh bằng Artifact Tool.
Lần xuất lớn phải dừng sau thời gian chờ kéo dài; chỉ dùng XlsxWriter cho
bước ghi cuối, sau đó kiểm tra toàn bộ dữ liệu bằng công cụ đọc độc lập.
"""
import json
from pathlib import Path
from datetime import datetime
import xlsxwriter

ROOT=Path(__file__).resolve().parents[1]
DIR=ROOT/'scripts/data/h2-chain-readiness-20261022'
OUT=ROOT/'outputs/HOA_DON_GOC_T7_T12_2025_F0_F1_F2.xlsx'
assert not OUT.exists(), 'Không ghi đè file hiện có'
data=json.loads((DIR/'workbook-input.json').read_text(encoding='utf-8'))
book=xlsxwriter.Workbook(str(OUT),{'strings_to_formulas':False,'strings_to_urls':False,'default_date_format':'dd/mm/yyyy'})
book.set_calc_mode('auto')
main=book.add_worksheet('F0, F1, F2, ...');src=book.add_worksheet('Chi tiet nguon')
base={'font_name':'Arial','font_size':10,'valign':'vcenter'}
body=book.add_format(base)
num=book.add_format(dict(base,num_format='#,##0;[Red](#,##0);0'))
identifier=book.add_format(dict(base,num_format='00000000'))
day=book.add_format(dict(base,num_format='dd/mm/yyyy'))
head=book.add_format(dict(base,bold=True,font_color='white',bg_color='#203746',text_wrap=True,align='center'))
title=book.add_format(dict(base,font_size=14,bold=True))
warning=book.add_format({'bg_color':'#FFF2CC','font_color':'#9C5700','bold':True})
headers=['STT','Người mua hàng','MST/CCCD','Mã đơn hàng F0','Số HĐ liên quan','Thanh toán trống tại F',
 'Ký hiệu F0','Số hóa đơn F0','Ngày F0','Tiền hàng F0','Thuế F0','Thanh toán F0',
 'Loại F1','Ký hiệu F1','Số hóa đơn F1','Ngày F1','Tiền hàng F1','Thuế F1','Thanh toán F1',
 'Loại F2','Ký hiệu F2','Số hóa đơn F2','Ngày F2','Tiền hàng F2','Thuế F2','Thanh toán F2',
 'Tiền hàng hiện tại','Thuế hiện tại','Thanh toán hiện tại','Tình trạng chuỗi','Ghi chú']
sh=['Ký hiệu F0','Số hóa đơn F0','F','Ký hiệu','Số hóa đơn','Ngày hóa đơn','Loại',
 'Ký hiệu HĐ được xử lý','Số HĐ được xử lý','Người mua hàng','MST/CCCD','Mã đơn hàng',
 'Tiền hàng','Thuế','Thanh toán','Trạng thái hóa đơn','Trạng thái CQT','Mã CQT',
 'Nguồn bảng tổng','Dòng bảng tổng','Nguồn liên kết','Dòng liên kết']
widths=[7,34,17,24,14,22,15,16,15,19,17,20,16,15,16,15,19,17,20,16,15,16,15,19,17,20,21,19,23,25,55]
sw=[15,16,8,15,16,15,16,19,20,34,17,24,19,17,20,28,18,38,65,16,65,16]
for s,w in [(main,widths),(src,sw)]:
 s.hide_gridlines(2);s.set_default_row(25);s.freeze_panes(6,0)
 for c,width in enumerate(w):s.set_column(c,c,width,body)
 s.set_row(1,30);s.set_row(5,48)
for c in [7,14,21]:main.set_column(c,c,widths[c],identifier)
for c in [8,15,22]:main.set_column(c,c,widths[c],day)
for c in [9,10,11,16,17,18,23,24,25,26,27,28]:main.set_column(c,c,widths[c],num)
for c in [1,4,8]:src.set_column(c,c,sw[c],identifier)
src.set_column(5,5,sw[5],day)
for c in [12,13,14]:src.set_column(c,c,sw[c],num)
main.set_tab_color('#203746');src.set_tab_color('#7594A8')
main.write(1,0,'CHUỖI HÓA ĐƠN GỐC 01/07–31/12/2025',title)
main.write(2,0,'37.087 gốc · 662 chuỗi có liên quan · 671 hóa đơn liên quan, gồm phát sinh năm 2026. Đơn vị: VND.',body)
main.write(3,0,'Thanh toán trống được giữ nguyên và tô vàng; lọc cột F để kiểm tra. Mã đơn hàng là ghi chú nguồn, không dùng để nối chuỗi.',body)
main.write(4,0,'Tiền hiện tại: cộng điều chỉnh; hóa đơn thay thế thay khoản cũ. Thiếu khoản thanh toán đang dùng thì tổng thanh toán để trống.',body)
src.write(1,0,'NGUỒN ĐỐI CHIẾU TỪNG HÓA ĐƠN',title)
src.write(2,0,'Bảng tổng xuất 09/10/2026; báo cáo liên kết 2024–2025 và 2026 đến 22/10/2026. Đường dẫn và dòng nguồn ở cột S–V.',body)
src.write(3,0,'Liên kết 31054 → 34893 bổ sung từ ảnh meInvoice người dùng gửi, lưu tại data/Bo sung/2026-10-22/; các nguồn nguyên bản không sửa.',body)
src.write(4,0,'Giữ trống 436 ô thanh toán; không tự coi là 0. F1/F2 sắp theo liên kết cha trước con, rồi ngày và ký hiệu/số.',body)
def kind(v):return 'Gốc' if not v['relationship'] else 'Thay thế' if v['relationship']['kind']=='replacement' else 'Điều chỉnh'
def col(n):
 s='';n+=1
 while n:s=chr(65+(n-1)%26)+s;n=(n-1)//26
 return s
expected=[];sr=6;first_flag=None;special_row=None
for i,c in enumerate(data['chains']):
 inv=c['invoices'];rt=inv[0];r=i+6;excel=r+1
 row=[i+1,rt['buyer'] or rt['name'],rt['tax_id'] or None,rt['order'] or None,len(inv)-1,None]
 row += rt['key']+[datetime.fromisoformat(rt['date'])]+rt['money']
 for f in (1,2):
  v=inv[f] if f<len(inv) else None
  row += ([kind(v)]+v['key']+[datetime.fromisoformat(v['date'])]+v['money']) if v else [None]*7
 last=0
 for f,v in enumerate(inv):
  if kind(v)=='Thay thế':last=f
 active=list(range(last,len(inv)))
 blank=['F'+str(f) for f,v in enumerate(inv) if v['money'][2] is None]
 totals=[None if any(inv[f]['money'][m] is None for f in active) else sum(inv[f]['money'][m] for f in active) for m in range(3)]
 special=rt['key']==['1C25MUV','00031054']
 row += [None,None,None,'Chưa có HĐ liên quan' if len(inv)==1 else 'Có thay thế' if last else 'Có điều chỉnh',
  'Liên kết 31054 → 34893 xác nhận bằng ảnh meInvoice; 34893 → 2639 theo báo cáo 2026.' if special else None]
 main.write_row(r,0,row)
 for m,total in enumerate(totals):
  cells=[f'{col([9,16,23][f]+m)}{excel}' for f in active];refs=','.join(cells)
  main.write_formula(r,26+m,f'=IF(COUNT({refs})={len(cells)},SUM({refs}),"")',num,'' if total is None else total)
 formula=f'=IF(L{excel}="","F0; ","")&IF(AND(N{excel}<>"",S{excel}=""),"F1; ","")&IF(AND(U{excel}<>"",Z{excel}=""),"F2; ","")'
 main.write_formula(r,5,formula,body,''.join(f+'; ' for f in blank))
 if blank and first_flag is None:first_flag=excel
 if special:special_row=excel
 expected.append(dict(row=excel,root=rt['key'],blank=blank,total=totals,active=active,invoices=[dict(key=v['key'],date=v['date'],money=v['money']) for v in inv]))
 for f,v in enumerate(inv):
  e=v['relationship'];parent=e['parent'] if e else [None,None]
  raw=rt['key']+['F'+str(f)]+v['key']+[datetime.fromisoformat(v['date']),kind(v)]+parent+[
   v['buyer'] or v['name'],v['tax_id'] or None,v['order'] or None]+v['money']+[
   v['status'],v['cqt_status'] or None,v['cqt'] or None,v['source'].replace('\\','/'),v['row'],e['source'].replace('\\','/') if e else None,e['row'] if e else None]
  src.write_row(sr,0,raw);sr+=1
end=len(data['chains'])+5
for s,last,cols,name in [(main,end,headers,'Chuoi_H2_2025'),(src,sr-1,sh,'Nguon_H2_2025')]:
 s.add_table(5,0,last,len(cols)-1,{'name':name,'style':'Table Style Medium 2','columns':[{'header':t,'header_format':head} for t in cols]})
main.conditional_format(6,5,end,5,{'type':'formula','criteria':'=$F7<>""','format':warning})
for series,pay in [('G','L'),('N','S'),('U','Z')]:
 main.conditional_format(f'{pay}7:{pay}{end+1}',{'type':'formula','criteria':f'=AND(${series}7<>"",ISBLANK(${pay}7))','format':warning})
main.conditional_format(f'AC7:AC{end+1}',{'type':'formula','criteria':'=$AC7=""','format':warning})
src.conditional_format(f'O7:O{sr}',{'type':'formula','criteria':'=ISBLANK($O7)','format':warning})
book.close()
(DIR/'expected-workbook.json').write_text(json.dumps(dict(expected=expected,rawCount=sr-6,firstFlag=first_flag,specialRow=special_row),ensure_ascii=False),encoding='utf-8')
(DIR/'export-method.txt').write_text('Artifact Tool: authored, recalculated and rendered main/source/blank views. Large final export did not complete after prolonged wait; stopped only verified exporter PID 20632. Final XLSX written using bundled XlsxWriter, with identical layout/formulas and explicit formula caches; independently verify all rows using verify_h2_workbook_20261022.py.\n',encoding='utf-8')
print(OUT)
