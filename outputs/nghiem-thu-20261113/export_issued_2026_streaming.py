"""Xuất theo luồng sau khi Artifact Tool hết bộ nhớ khi lưu bảng hơn 3 triệu ô."""
from pathlib import Path
from datetime import datetime
from collections import Counter
from decimal import Decimal
import json,sys,math,os
import xlsxwriter
sys.stdout.reconfigure(encoding='utf-8');O=Path(__file__).parent
d=json.loads((O/'issued-2026-report-data.json').read_text(encoding='utf-8'))
v=json.loads((O/'issued-2026-source-verification.json').read_text(encoding='utf-8'));assert v['passed']
CHAINS='--chains' in sys.argv
COMPACT='--compact' in sys.argv or CHAINS
p=O/('TONG_HOP_HD_2026_CHO_DON_2024_2025_CHUOI.xlsx' if CHAINS else 'TONG_HOP_HD_2026_CHO_DON_2024_2025_RUT_GON.xlsx' if COMPACT else 'TONG_HOP_HD_PHAT_HANH_2026_CHO_DON_2024_2025.xlsx')
write_path=p.with_suffix('.xlsx.writing') if CHAINS else p
w=xlsxwriter.Workbook(write_path,{'constant_memory':True,'strings_to_numbers':False,'strings_to_formulas':False,'strings_to_urls':False})
w.set_properties({'title':'HĐ phát hành 2026 cho đơn 2024–2025','subject':'Đối chiếu meInvoice và AMIS đến 13/11/2026'})
navy='#253B53';ink='#233348';blue='#C1E5F5';red='#A22626'
fmts={}
def f(kind='text',shade=False,wrap=False):
 key=(kind,shade,wrap)
 if key not in fmts:
  z={'font_name':'Arial','font_size':10,'font_color':ink,'valign':'vcenter'}
  if shade:z['bg_color']=blue
  if wrap:z['text_wrap']=True
  if kind in ['money','phase_money']:
   z.update(num_format='#,##0;[Red](#,##0);0',align='right')
   if kind=='phase_money':z.update(indent=1,right=1,right_color='#D7E0E8')
  elif kind=='date':z.update(num_format='dd/mm/yyyy',align='center')
  elif kind=='id':z.update(num_format='@',align='left')
  elif kind=='center':z['align']='center'
  fmts[key]=w.add_format(z)
 return fmts[key]
title=w.add_format({'font_name':'Arial','font_size':15,'bold':True,'font_color':ink})
header=w.add_format({'font_name':'Arial','font_size':10,'bold':True,'font_color':'#FFFFFF','bg_color':navy,'align':'center','valign':'vcenter','text_wrap':True})
note=w.add_format({'font_name':'Arial','font_size':10,'italic':True,'font_color':'#5A6878'})
bold=w.add_format({'font_name':'Arial','font_size':10,'bold':True,'font_color':ink})
total_text=w.add_format({'font_name':'Arial','font_size':10,'bold':True,'bg_color':'#E8EEF4','font_color':ink,'valign':'vcenter'})
total_num=w.add_format({'font_name':'Arial','font_size':10,'bold':True,'bg_color':'#E8EEF4','font_color':ink,'valign':'vcenter','num_format':'#,##0;[Red](#,##0);0','align':'right'})
amber=w.add_format({'font_name':'Arial','font_size':10,'font_color':'#875D0F','bg_color':'#FFF1D6','valign':'vcenter'})
amber_num=w.add_format({'font_name':'Arial','font_size':10,'font_color':'#875D0F','bg_color':'#FFF1D6','valign':'vcenter','num_format':'#,##0','align':'right'})
def begin(name,caption,subtitle,heads,widths,rows,freeze):
 s=w.add_worksheet(name);s.hide_gridlines(2);s.set_default_row(23)
 for col,width in enumerate(widths):s.set_column_pixels(col,col,width)
 if COMPACT:
  caption,subtitle={
   'HĐ 2026':('HĐ phát hành 2026','Đơn 2024–2025. Chốt nguồn 13/11/2026.'),
   'BHH và CT liên quan':('BHH và HĐ liên quan','HĐ cũ và HĐ phát hành 2026.'),
   'Chuỗi thiếu BHH':('Chuỗi chưa xác định BHH','Tiền hàng và VAT của các chuỗi đều bằng 0.'),
  }.get(name,(caption,subtitle))
 s.write_string(1,0,caption,title);s.write_string(2,0,subtitle,note)
 s.set_row(4,46)
 for j,x in enumerate(heads):s.write_string(4,j,x,header)
 s.autofilter(4,0,len(rows)+4,len(heads)-1);s.freeze_panes(5,freeze)
 s.set_landscape();s.repeat_rows(0,4)
 return s
def put(s,row,col,value,fmt,dated=False):
 if value is None or value=='':s.write_blank(row,col,None,fmt)
 elif dated:s.write_datetime(row,col,datetime.strptime(value[:10],'%Y-%m-%d'),fmt)
 elif isinstance(value,(int,float)):s.write_number(row,col,value,fmt)
 else:s.write_string(row,col,str(value),fmt)

end=len(d['rows'])+5

def write_summary():
 summary=w.add_worksheet('Tổng hợp');summary.hide_gridlines(2);summary.set_default_row(23);summary.set_tab_color(navy)
 for c,x in enumerate([335,130,130,135,195,175,225]):summary.set_column_pixels(c,c,x)
 summary.write_string(1,0,'Tổng hợp HĐ phát hành 2026 cho đơn 2024–2025',title)
 summary.write_string(2,0,'Ngày chốt nguồn: 13/11/2026. Bao gồm HĐ xuất mới/xuất lại, điều chỉnh và thay thế trong các chuỗi liên quan.',note)
 summary.set_row(4,48)
 for c,x in enumerate(['Loại HĐ phát hành 2026','Đơn 2024','Đơn 2025','Tổng số HĐ','Tiền hàng trên HĐ (đ)','VAT trên HĐ (đ)','Thanh toán có số liệu (đ)']):summary.write_string(4,c,x,header)
 cats=['Xuất mới / xuất lại','Điều chỉnh trong chuỗi về 0','Thay thế bằng HĐ 0','Điều chỉnh khác','Thay thế khác']
 end=len(d['rows'])+5;cache=[];type_col='G' if COMPACT else 'H'
 for excelrow,cat in enumerate(cats,6):
  selected=[r for r in d['rows'] if r[5]==cat]
  vals=[sum(r[1]=='2024' for r in selected),sum(r[1]=='2025' for r in selected),len(selected)]+[sum(r[j] for r in selected if r[j] is not None) for j in [9,10,11]]
  cache.append(vals);summary.write_string(excelrow-1,0,cat,f())
  formulas=[f'=COUNTIFS(\'HĐ 2026\'!$B$6:$B${end},"2024",\'HĐ 2026\'!${type_col}$6:${type_col}${end},$A{excelrow})',f'=COUNTIFS(\'HĐ 2026\'!$B$6:$B${end},"2025",\'HĐ 2026\'!${type_col}$6:${type_col}${end},$A{excelrow})',f'=SUM(B{excelrow}:C{excelrow})']+[f'=SUMIFS(\'HĐ 2026\'!${col}$6:${col}${end},\'HĐ 2026\'!${type_col}$6:${type_col}${end},$A{excelrow})' for col in ['I','J','K']]
  for c,(formula,value) in enumerate(zip(formulas,vals),1):summary.write_formula(excelrow-1,c,formula,f('money'),value)
 summary.set_row(10,28);summary.write_string(10,0,'Tổng HĐ phát hành 2026',total_text)
 totals=[sum(r[c] for r in cache) for c in range(6)]
 assert totals==[40161,42046,82207,*d['summary']['money_sums']]
 for c,value in enumerate(totals,1):summary.write_formula(10,c,f'=SUM({xlsxwriter.utility.xl_col_to_name(c)}6:{xlsxwriter.utility.xl_col_to_name(c)}10)',total_num,value)
 stats=[('BHH / BHDV đã nối được',65241),('CT trả/giảm liên quan',485),('Chuỗi HĐ cũ chưa xác định BHH',2989),('HĐ 2026 thuộc các chuỗi thiếu BHH',3692),('HĐ 2026 chưa có số thanh toán',4022),('HĐ 2026 nối theo mã đơn',4011),('Bảng kê nguồn đã kiểm',235263)]
 if COMPACT:stats=[('Chứng từ AMIS đã nối',65726),('HĐ chưa xác định chứng từ AMIS',3692),('Chuỗi cũ chưa xác định chứng từ',2989),('HĐ chưa có số thanh toán',4022)]
 for rr,(label,value) in enumerate(stats,13):
  highlighted=rr in ({14,15} if COMPACT else {15,16})
  summary.write_string(rr,0,label,amber if highlighted else f())
  if label in ['HĐ 2026 chưa có số thanh toán','HĐ chưa có số thanh toán']:summary.write_formula(rr,1,f'=COUNTBLANK(\'HĐ 2026\'!K6:K{end})',f('money'),value)
  else:summary.write_number(rr,1,value,amber_num if highlighted else f('money'))
 summary.write_string(22,0,'Cách đọc số liệu',bold)
 notes=[
  'Tiền trên HĐ cộng theo dấu của từng HĐ đã phát hành, bao gồm HĐ đã bị thay thế. Tổng này không đại diện doanh thu còn lại.',
  'Kết quả chuỗi ở cột S và số dư T–V lặp trên các HĐ cùng chuỗi. Tra theo F0, tránh cộng lặp số dư chuỗi.',
  'F0, F1, F2 theo liên kết điều chỉnh/thay thế. HĐ xuất lại nối với HĐ cũ qua BHH ở cột AD khi có bằng chứng.',
  'Ô thanh toán trống trên HĐ được giữ nguyên. Số dư chuỗi sử dụng số 0 ở các trường hợp đã có bằng chứng xác nhận riêng.',
  'Chuỗi thiếu BHH có tiền hàng/VAT bằng 0. Chuỗi có thanh toán trống được ghi rõ, không mặc định là đã có đủ ba số tiền 0.',
  'Căn cứ nối BHH tại cột X. Các liên kết theo mã đơn nguyên vẹn được ghi riêng với liên kết chứng từ đã chốt.',
  'Mã đơn chỉ sang BHH khác mapping đã chốt được giữ theo mapping hiện tại. Tab BHH có HĐ cũ ở cột L và HĐ 2026 ở cột M.',
 ]
 if COMPACT:notes=[
  'Mỗi HĐ 2026 chỉ xuất hiện một lần. Chứng từ AMIS có thể lặp khi có nhiều HĐ điều chỉnh, thay thế hoặc xuất lại.',
  'Tra theo BHH ở tab BHH và CT liên quan: cột E là HĐ cũ, F là các HĐ 2026. Khi có nhiều HĐ, cột G ghi HĐ đang map AMIS.',
  'Tiền trên HĐ cộng theo dấu của từng HĐ đã phát hành, bao gồm HĐ đã bị thay thế. Tổng này không đại diện doanh thu còn lại.',
  'Ô thanh toán trống được giữ nguyên. Tab Chuỗi thiếu BHH phân biệt chuỗi đã về 0 và chuỗi còn trống số thanh toán.',
  'Ghi chú chỉ nêu các trường hợp cần rà: nối theo mã đơn, thiếu chứng từ, liên quan nhiều chứng từ hoặc nguồn có xung đột.',
 ]
 if CHAINS:
  notes=[
   'Mỗi HĐ 2026 chỉ xuất hiện một lần ở tab HĐ 2026. Dữ liệu và số tiền của bảng này được giữ nguyên.',
   'F0–F3 trên tab chuỗi là các đợt xử lý: HĐ gốc; ĐC/TT trước 2026; xử lý năm 2026; xuất mới/xuất lại năm 2026.',
   'Mỗi F có HĐ, ngày và tổng thanh toán gồm VAT. Mỗi HĐ có dòng ngày/tiền riêng; HĐ gốc chỉ ghi một lần trong nhóm.',
   'Tổng hợp tiền toàn bộ dùng tab HĐ 2026. Một HĐ có thể liên quan nhiều BHH trong tab chuỗi; HĐ đã bị thay thế là chứng từ cũ.',
   'Ô tiền màu vàng: bảng kê chưa có số thanh toán, không mặc định là 0. Các chuỗi chưa rõ BHH để trống cột chứng từ.',
   'ĐC: điều chỉnh. TT: thay thế. Quan hệ HĐ được xử lý theo bảng kê giữ nguyên ở cột L tab HĐ 2026.',
   'Tiền trên HĐ cộng theo dấu của từng HĐ đã phát hành, gồm HĐ đã bị thay thế. Tổng này không đại diện doanh thu còn lại.',
  ]
 for rr,x in enumerate(notes,23):summary.write_string(rr,0,x,f())
 summary.write_string(32,0,'Nguồn đối chiếu',bold)
 sources=[
  'AMIS: Sổ chi tiết bán hàng 2024 và 2025, bộ dữ liệu 13/11/2026. Vị trí chứng từ tại tab BHH và CT liên quan.',
  'meInvoice: bảng kê HĐ đã sử dụng, chi tiết, điều chỉnh và thay thế ngày 13/11/2026. Dòng nguồn tại cột AC.',
  'Liên kết lịch sử: bảng đối chiếu BHH 09/11, Import HD moi đã bổ sung số HĐ phát hành 11/11, mapping nghiệm thu 13/11.',
  'Bảng xử lý 2025: https://docs.google.com/spreadsheets/d/1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA',
  'Bảng kế toán: https://docs.google.com/spreadsheets/d/1C5hoPDvZKB1vGFSUG9rzChcORt2L9NnsqETVHxPui9I',
  '31 quan hệ lịch sử bổ sung đã được kiểm. Bảng tổng và 18.790 quan hệ hiện tại đã đối chiếu độc lập với ZIP nguồn.',
 ]
 if COMPACT:sources[1]='meInvoice: bảng kê HĐ đã sử dụng, chi tiết, điều chỉnh và thay thế ngày 13/11/2026. Giữ nguyên bộ HĐ và số tiền đã đối chiếu.'
 if CHAINS:sources[0]='AMIS: Sổ chi tiết bán hàng 2024 và 2025, bộ dữ liệu 13/11/2026. Chứng từ và liên kết HĐ nằm tại tab Chuỗi F0 F1 F2.'
 for rr,x in enumerate(sources,33):summary.write_string(rr,0,x,f())
 return totals

totals=[40161,42046,82207,*d['summary']['money_sums']] if CHAINS else write_summary()

order=[0,1,6,7,8,2,3,5,9,10,11,4,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29]
def compact_invoice(r):
 docs=list(dict.fromkeys(x for field in [r[2],r[3]] for x in str(field or '').split('\n') if x))
 notices=[]
 if not docs:notices.append('Chưa xác định chứng từ AMIS')
 if 'Mã đơn' in r[23]:notices.append('Nối theo mã đơn')
 if len(docs)>1:notices.append(f'Liên quan {len(docs)} chứng từ')
 if 'Mã đơn chỉ sang BHH khác' in r[27]:notices.append('Giữ mapping chứng từ đã chốt')
 if 'không hợp lệ' in str(r[25]).lower():notices.append('CQT không hợp lệ')
 assert '\n' not in r[15] and '\n' not in r[16], 'Cần kiểm tra HĐ có nhiều cha'
 parent=r[15]+'/'+r[16] if r[15] and r[16] else ''
 return [r[0],r[1],r[6],r[7],r[8],'\n'.join(docs),r[5],r[4],r[9],r[10],r[11],parent,'; '.join(notices)]
if COMPACT:
 rows=[compact_invoice(r) for r in d['rows']]
 heads=['STT','Năm đơn','Ký hiệu HĐ 2026','Số HĐ 2026','Ngày HĐ','BHH / chứng từ','Loại phát hành','Mã đơn hàng','Tiền hàng (đ)','VAT (đ)','Thanh toán (đ)','HĐ được điều chỉnh/thay thế','Ghi chú cần rà']
 widths=[55,80,110,115,110,195,225,215,150,140,165,260,350]
 dates={4};monies={8,9,10};ids={2,3};centers={0,1};wraps={5,12}
 subtitle='Mỗi HĐ một dòng. Chỉ ghi HĐ được xử lý khi có điều chỉnh/thay thế; liên kết HĐ cũ theo BHH nằm ở bảng tra chứng từ.'
else:
 rows=[[r[j] for j in order] for r in d['rows']];heads=[d['headers'][i] for i in order]
 widths=[62,90,115,120,115,210,180,235,155,145,165,225,120,125,115,130,130,105,270,165,155,180,250,370,215,160,420,630,660,290]
 dates={4,14};monies={8,9,10,19,20,21};ids={2,3,12,13,15,16};centers={0,1,17};wraps={5,6}
 subtitle='Mỗi dòng là một HĐ duy nhất. Cột F: BHH; cột W: HĐ hiện map AMIS; cột AD: HĐ cũ cùng BHH.'
s=begin('HĐ 2026','HĐ phát hành năm 2026 cho đơn 2024–2025',subtitle,heads,widths,rows,4)
s.set_tab_color(navy)
for i,r in enumerate(rows,5):
 shade=i%2==1
 for col,value in enumerate(r):
  kind='date' if col in dates else 'money' if col in monies else 'id' if col in ids else 'center' if col in centers else 'text'
  put(s,i,col,value,f(kind,shade,col in wraps),kind=='date')
 lines=max(str(r[5] or '').count('\n')+1,math.ceil(len(str(r[12] or ''))/52)) if COMPACT else max(str(r[5] or '').count('\n')+1,str(r[6] or '').count('\n')+1)
 if lines>1:s.set_row(i,min(4,lines)*14+8)
 if i%20000==5:print('STREAM HĐ',i-4,flush=True)
s.conditional_format(5,5,end-1,5,{'type':'formula','criteria':'=$F6=""' if COMPACT else '=AND($F6="",$G6="")','format':amber})

if CHAINS:
 chain=json.loads((O/'issued-2026-chain-view.json').read_text(encoding='utf-8'))
 assert chain['summary']['presentation']=='business_phases_with_dates_and_payments'
 chainwidths=[85,85,205]+[250,110,165]*4+[190,430]
 chain_audit=json.loads((O/'issued-2026-chain-view-audit.json').read_text(encoding='utf-8'))
 palette=[('#305E8C','#EAF2FA','#DDEAF8'),('#9A6700','#FFF5DF','#FFEDC4'),('#6B4C9A','#F4EEFA','#E9DFF4'),('#2F6F46','#EAF5EC','#D9EDDE')]
 chain_formats={}
 def cf(col,shade,kind,wrap):
  if col<3 or col==16:return f(kind,shade,wrap)
  key=(col,shade,kind,wrap)
  if key not in chain_formats:
   z={'font_name':'Arial','font_size':10,'font_color':ink,'valign':'vcenter','bg_color':palette[(col-3)//3][1+int(shade)] if col<15 else '#E4F0EF'}
   if wrap:z['text_wrap']=True
   if kind=='date':z.update(num_format='dd/mm/yyyy',align='center')
   if kind=='phase_money':z.update(num_format='#,##0;[Red](#,##0);0',align='right',indent=1,right=1,right_color='#D7E0E8')
   if col==15:z.update(num_format='#,##0;[Red](#,##0);0',align='right',indent=1,bold=True)
   chain_formats[key]=w.add_format(z)
  return chain_formats[key]
 phase_headers=[w.add_format({'font_name':'Arial','font_size':10,'bold':True,'font_color':'#FFFFFF','bg_color':p[0],'align':'center','valign':'vcenter','text_wrap':True,'right':1,'right_color':'#FFFFFF'}) for p in palette]
 s=w.add_worksheet('Chuỗi F0 F1 F2');s.hide_gridlines(2);s.set_default_row(24)
 for col,width in enumerate(chainwidths):s.set_column_pixels(col,col,width)
 s.write_string(1,0,'Chuỗi xử lý HĐ',title)
 s.write_string(1,16,'ĐC: điều chỉnh; TT: thay thế.',note)
 s.write_string(2,0,'Giá trị sau cùng gồm VAT; hiện một lần mỗi chuỗi.',note)
 s.write_string(2,16,'Ô vàng: nguồn trống. Tổng HĐ duy nhất: tab HĐ 2026.',note)
 s.set_row(4,32);s.set_row(5,38)
 phase_leaf=w.add_format({'font_name':'Arial','font_size':10,'bold':True,'font_color':'#FFFFFF','bg_color':navy,'align':'center','valign':'vcenter','text_wrap':True,'right':1,'right_color':'#FFFFFF'})
 s.merge_range(4,0,4,2,'Chứng từ AMIS',header)
 for phase,label in enumerate(chain['phase_titles']):s.merge_range(4,3+phase*3,4,5+phase*3,label,phase_headers[phase])
 s.merge_range(4,15,4,16,'Kết quả sau xử lý',header)
 for col,label in enumerate(chain['headers']):s.write_string(5,col,label,phase_headers[(col-3)//3] if 3<=col<15 else header)
 s.write_comment(5,15,'Giá trị thanh toán còn lại sau xử lý, đã gồm VAT. Cộng HĐ còn hiệu lực và các điều chỉnh có dấu; loại HĐ đã bị thay thế, hủy hoặc CQT không hợp lệ; gồm HĐ xuất lại nối qua BHH. Nếu Thanh toán nguồn trống nhưng Tiền hàng + VAT xác nhận bằng 0 thì tính 0, ghi tại cột Q. Mỗi chuỗi chỉ hiện một lần. Cùng BHH có thể có nhiều chuỗi gốc nên không dùng cột này làm tổng doanh thu toàn bộ.')
 s.autofilter(5,0,len(chain['rows'])+5,16);s.freeze_panes(6,3);s.set_landscape();s.repeat_rows(0,5)
 s.set_tab_color('#366B87')
 last_gid=None
 for i,r in enumerate(chain['rows'],6):
  shade=r[0]%2==1
  for col,value in enumerate(r):
   kind='date' if col in [4,7,10,13] else 'phase_money' if col in [5,8,11,14] else 'center' if col in [0,1] else 'text'
   fmt=cf(col,shade,kind,col in [2,3,6,9,12,16])
   if col in [5,8,11,14] and r[col-2] and value is None:fmt=amber_num
   if col==15 and r[0]!=last_gid:
    rec=chain_audit['records'][r[0]-1]
    if value is None:put(s,i,col,value,amber_num)
    else:s.write_formula(i,col,rec['final_formula'],fmt,value)
   else:put(s,i,col,value,fmt,kind=='date')
  last_gid=r[0]
  logical_lines=[]
  for col in [2,3,6,9,12,16]:
   parts=str(r[col] or '').split('\n')
   logical_lines.append(sum(max(1,math.ceil(len(part)/max(1,int(chainwidths[col]/6.5)))) for part in parts))
  s.set_row(i,min(400,max(23,max(logical_lines)*14+8)))
  if i%20000==6:print('STREAM CHUOI',i-5,flush=True)
 s.conditional_format(6,2,len(chain['rows'])+5,2,{'type':'formula','criteria':'=$C7=""','format':amber})
 w.close()
 os.replace(write_path,p)
 (O/'issued-2026-chain-authoring.json').write_text(json.dumps({'passed':True,'writer':'streaming fallback after Artifact Tool out-of-memory at 6 GB','updated_in_place':True,'totals':totals,'sheets':[['HĐ 2026',82207,13],['Chuỗi F0 F1 F2',len(chain['rows']),len(chain['headers'])]]},ensure_ascii=False,indent=2),encoding='utf-8')
 print('EXPORTED',p.name,p.stat().st_size,flush=True)
 sys.exit(0)

def compact_document(r):
 notices=[]
 newkeys=r[12].split('\n') if r[12] else []
 if r[4] and (len(newkeys)>1 or r[4] not in newkeys):notices.append('Đang map AMIS: '+r[4])
 elif not r[4]:notices.append('Chưa có HĐ đang map AMIS')
 if r[10]:notices.append(r[10])
 return [r[0],r[1],r[2],r[8],r[11],r[12],'; '.join(notices)]
if COMPACT:
 refrows=[compact_document(r) for r in d['bhh_rows']]
 refheads=['Năm đơn','BHH / chứng từ','Ngày chứng từ','Mã đơn hàng','HĐ gốc 2024/2025 liên quan','HĐ phát hành 2026 liên quan','Ghi chú khi có nhiều HĐ / cần rà']
 refwidths=[90,205,115,235,300,460,510];refdates={2};refmonies=set();refwraps={3,4,5,6}
 refsubtitle='Mỗi chứng từ một dòng. HĐ cũ ở cột E, HĐ 2026 ở cột F. Cột G chỉ ghi thêm khi cần phân biệt HĐ đang map hoặc rà nguồn.'
else:
 refrows=[r[:9]+['AMIS | '+r[9].split(' | ',1)[1]]+r[10:] for r in d['bhh_rows']]
 refheads=d['bhh_headers'];refwidths=[95,205,115,410,250,165,155,175,255,620,650,300,460];refdates={2};refmonies={5,6,7};refwraps={3,9,10,11,12}
 refsubtitle='Một dòng mỗi chứng từ. Cột E: HĐ hiện tại; cột L: HĐ cũ; cột M: toàn bộ HĐ phát hành 2026 liên quan.'
s=begin('BHH và CT liên quan','BHH và chứng từ AMIS liên quan',refsubtitle,refheads,refwidths,refrows,2)
for i,rr in enumerate(refrows,5):
 shade=i%2==1
 for c,value in enumerate(rr):
  kind='date' if c in refdates else 'money' if c in refmonies else 'center' if c==0 else 'text'
  put(s,i,c,value,f(kind,shade,c in refwraps),c in refdates)
 lines=max(max(1,math.ceil(len(part)/max(1,int(refwidths[c]/6.5)))) for c in refwraps for part in str(rr[c] or '').split('\n'))
 lines=max(lines,max(str(rr[c] or '').count('\n')+1 for c in refwraps))
 s.set_row(i,min(400,max(23,lines*14+8)))
 if i%20000==5:print('STREAM BHH',i-4,flush=True)

if COMPACT:
 unrows=[r[:7] for r in d['unmapped_rows']];unheads=d['unmapped_headers'][:7];unwidths=[65,95,250,115,460,105,300];unmonies=set()
else:
 unrows=d['unmapped_rows'];unheads=d['unmapped_headers'];unwidths=[65,95,250,115,460,105,290,160,145,175,420,670];unmonies={7,8,9}
s=begin('Chuỗi thiếu BHH','Chuỗi HĐ cũ chưa xác định BHH','Tiền hàng/VAT của các chuỗi đều bằng 0. Cột G phân biệt thanh toán 0 và thanh toán trống.',unheads,unwidths,unrows,3)
for i,r in enumerate(unrows,5):
 shade=i%2==1
 for c,value in enumerate(r):
  kind='date' if c==3 else 'money' if c in unmonies else 'center' if c in [0,1,5] else 'text'
  put(s,i,c,value,f(kind,shade,c==4),c==3)
 s.set_row(i,max(23,(str(r[4]).count('\n')+1)*14+8))
w.close()
(O/('issued-2026-compact-authoring.json' if COMPACT else 'issued-2026-authoring-verification.json')).write_text(json.dumps({'passed':True,'writer':'streaming fallback after Artifact Tool out-of-memory at 6 GB','compact':COMPACT,'totals':totals,'blankPayment':4022,'sheets':[['HĐ 2026',82207,len(heads)],['BHH và CT liên quan',65726,len(refheads)],['Chuỗi thiếu BHH',2989,len(unheads)]]},ensure_ascii=False,indent=2),encoding='utf-8')
print('EXPORTED',p.name,p.stat().st_size,flush=True)
