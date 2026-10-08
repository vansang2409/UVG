import sys,json,collections,openpyxl
from pathlib import Path
from decimal import Decimal
sys.stdout.reconfigure(encoding='utf8');r=Path(r'D:\UVG');sys.path.insert(0,str(r/'scripts'));from prepare_invoice_reports import rows,invoice_key,invoice_day,amount
w=openpyxl.load_workbook(r/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',data_only=True);s=w['Dieu chinh'];chains=[];wanted=set()
for n,v in enumerate(s.iter_rows(min_row=3,max_row=1049,max_col=55,values_only=True),3):
 olds=[]
 for i in range(0,24,6):
  if v[i]:olds.append({'key':list(invoice_key(v[i],v[i+1])),'date':str(v[i+2]),'money':list(v[i+3:i+6])});wanted.add(tuple(olds[-1]['key']))
 state='unknown' if any(x is None for x in v[24:27]) else 'zero' if all(x==0 for x in v[24:27]) else 'nonzero'
 chains.append({'key':olds[0]['key'],'row':n,'old':olds,'prior':list(v[24:27]),'note':v[27],'state':state,'plan':list(v[28:52]),'predicted':list(v[52:55])})
headers={};detail=collections.defaultdict(list);files=sorted((r/'data/Du lieu Meinvoice/Bang_ke_hoa_don_da_su_dung_2024_2025').glob('*.xlsx'))
for f in files:
 for n,v in rows(f):
  if n<6 or amount(v.get('A')) is None:continue
  k=invoice_key(v.get('B'),v.get('C'))
  if k in wanted:headers[k]={'date':str(invoice_day(v.get('D'))),'money':[amount(v.get(c)) for c in ['P','Q','R']],'address':v.get('I'),'taxid':v.get('J'),'buyer':v.get('L'),'customer':v.get('H'),'source':f.name+' | dòng '+str(n)}
for fn in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
 mt='_MTT' in fn;f=r/'data/Du lieu Meinvoice'/fn;files.append(f)
 for n,v in rows(f):
  if n<7 or amount(v.get('A')) is None:continue
  k=invoice_key(v.get('F' if mt else 'D'),v.get('B'))
  if k not in wanted:continue
  detail[k].append({'name':v.get('S' if mt else 'L'),'unit':v.get('T' if mt else 'M'),'qty':amount(v.get('U' if mt else 'N')),'price':amount(v.get('V' if mt else 'O')),'goods':amount(v.get('W' if mt else 'P')),'rate':v.get('Z' if mt else 'S'),'tax':amount(v.get('AA' if mt else 'T')),'pay':amount(v.get('AB' if mt else 'U')),'discount':amount(v.get('Y' if mt else 'R')),'buyer':v.get('J' if mt else 'H'),'address':headers[k]['address'] if mt else v.get('F'),'taxid':v.get('I' if mt else 'G'),'email':v.get('K') if mt else None,'method':v.get('N' if mt else 'I'),'date':str(invoice_day(v.get('D' if mt else 'C'))),'source':fn+' | dòng '+str(n)})
ready=[];pending=[];checks=collections.Counter()
for x in chains:
 if x['state']=='zero':continue
 reasons=[];k=tuple(x['key']);ls=detail[k]
 if x['state']=='unknown':reasons.append('Cộng dồn còn ô tiền trống')
 if x['note']:reasons.append(x['note'])
 if k==('1C24TUV','00000003'):reasons.append('Kế hoạch đang ghi chờ kiểm tra khoản 1C25TUV/00000508')
 for op in x['old']:
  ok=tuple(op['key']);money=headers[ok]['money'];checkmoney=[None if v is None else Decimal(str(v)) for v in op['money']]
  if money!=checkmoney:reasons.append('Bảng hiện tại khác số tổng nguồn '+ '/'.join(ok))
  if any(v is None for v in money):reasons.append('Số tổng nguồn trống '+ '/'.join(ok))
  elif money[0]+money[1]!=money[2]:reasons.append('Tiền hàng + thuế không bằng thanh toán '+ '/'.join(ok))
 if not ls:reasons.append('Không tìm thấy chi tiết gốc')
 else:
  if any(l['qty'] in [None,0] for l in ls):reasons.append('Chi tiết gốc có số lượng 0 hoặc trống')
  if any(not l['name'] or not l['method'] for l in ls):reasons.append('Chi tiết gốc thiếu tên hàng hoặc hình thức thanh toán')
  if any(l['goods'] is None or l['price'] is None for l in ls):reasons.append('Chi tiết gốc thiếu thành tiền hoặc đơn giá')
  if any(l['discount'] not in [None,0] for l in ls):reasons.append('Có chiết khấu cần ánh xạ thêm')
  if [sum(l[c] or 0 for l in ls) for c in ['goods','tax','pay']]!=headers[k]['money']:reasons.append('Chi tiết gốc khác số tổng nguồn')
 for op in x['old']:
  opkey=tuple(op['key']);rates={str(l['rate']).strip() for l in detail[opkey] if l['rate'] not in [None,'']}
  if len(rates)!=1:reasons.append('Thuế suất chưa xác định duy nhất '+ '/'.join(opkey))
  elif not next(iter(rates)).rstrip('%').replace('.','',1).isdigit():reasons.append('Thuế suất đặc biệt cần kiểm tra '+ '/'.join(opkey))
 x['reasons']=list(dict.fromkeys(reasons));x['lines']=ls
 if reasons:pending.append(x)
 else:ready.append(x)
for x in pending:
 for reason in x['reasons']:checks[reason.split(' 1C')[0]]+=1
import hashlib
payload={'chains':chains,'ready':ready,'pending':pending,'zero':sum(x['state']=='zero' for x in chains),'headers':{'/'.join(k):v for k,v in headers.items()},'detail':{'/'.join(k):v for k,v in detail.items()},'hashes':{str(f):hashlib.sha256(f.read_bytes()).hexdigest() for f in files}}
def conv(v):return int(v) if isinstance(v,Decimal) and v==v.to_integral_value() else float(v) if isinstance(v,Decimal) else v
(r/'scripts/data/check-review/adjusted-import-data.json').write_text(json.dumps(payload,ensure_ascii=False,default=conv),encoding='utf8')
print(json.dumps({'chains':len(chains),'already_zero':payload['zero'],'ready':len(ready),'pending':len(pending),'new_invoices':sum(len(x['old']) for x in ready),'lines':sum(len(x['old'])-1+len(x['lines']) for x in ready),'pending_categories':dict(checks),'pending_examples':[{a:x[a] for a in ['key','reasons']} for x in pending[:12]]},ensure_ascii=False))
