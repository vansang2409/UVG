import sys,json,hashlib,unicodedata,re
from pathlib import Path
from collections import Counter
from decimal import Decimal
import openpyxl
r=Path(r'D:\UVG');h=r/'scripts/data/check-review';sys.path.insert(0,str(r/'scripts'))
from prepare_invoice_reports import rows
orig=json.loads((h/'zero-detail-review.json').read_text(encoding='utf8'));root_details={tuple(x['key']):x['detail'] for x in orig['roots']}
w=openpyxl.load_workbook(r/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',data_only=True);s=w['Dieu chinh'];chains=[];wanted=set()
for n,v in enumerate(s.iter_rows(min_row=3,values_only=True),3):
 if not all(z==0 for z in v[24:27]):continue
 children=[{'key':[v[i],v[i+1]],'money':list(v[i+3:i+6])} for i in range(6,24,6) if v[i]]
 wanted.update(tuple(x['key']) for x in children)
 chains.append({'row':n,'root':[v[0],v[1]],'children':children,'old_note':v[27]})
detail={k:[] for k in wanted};hashes={}
for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
 p=r/'data/Du lieu Meinvoice'/name;hashes[str(p.relative_to(r))]=hashlib.sha256(p.read_bytes()).hexdigest();mtt='_MTT' in name
 for n,v in rows(p):
  k=(v.get('F' if mtt else 'D'),v.get('B'))
  if k not in wanted:continue
  def num(c):
   z=v.get(c);return None if z in (None,'') else float(Decimal(z))
  detail[k].append({'row':n,'source':name,'name':v.get('S' if mtt else 'L'),'qty':num('U' if mtt else 'N'),'price':num('V' if mtt else 'O'),'goods':num('W' if mtt else 'P'),'discount':num('Y' if mtt else 'R'),'tax':num('AA' if mtt else 'T'),'payment':num('AB' if mtt else 'U')})
def norm(z):
 z=unicodedata.normalize('NFD',z or '');return re.sub(r'\s+',' ',''.join(c for c in z if not unicodedata.combining(c)).lower().replace('đ','d')).strip()
def lump(a):
 name=norm(a['name']);return 'dieu chinh' in name and (a['qty'] in (0,None)) and (a['price'] in (0,None))
counts=Counter();chain_counts=Counter();names=Counter();results=[]
for x in chains:
 msgs=[];cats=[]
 for y in x['children']:
  k=tuple(y['key']);a=detail[k];names.update(z['name'] for z in a)
  if not a:cat='Không tìm thấy chi tiết'
  elif all(lump(z) for z in a):cat='Chỉ điều chỉnh khoản tổng'
  elif any(lump(z) for z in a):cat='Có cả dòng tổng và dòng khác'
  else:cat='Có dòng chi tiết khác, cần đối chiếu'
  counts[cat]+=1;cats.append(cat)
  sums=[None if not a or any(z[f] is None for z in a) else sum(z[f] for z in a) for f in ['goods','tax','payment']]
  differences=[None if sums[i] is None or y['money'][i] is None else sums[i]-y['money'][i] for i in range(3)]
  y.update(category=cat,detail=a,detail_sums=sums,difference=differences)
  discrepancy=any(z not in (None,0) for z in differences)
  label='/'.join(k);msg=f'{label}: {cat.lower()}, {len(a)} dòng'
  if a and all(lump(z) for z in a):msg+='; SL/đơn giá đều 0 hoặc trống, không ghi lại từng hàng gốc'
  if discrepancy:msg+='; tiền chi tiết khác bảng tổng'
  refs=[]
  for src in dict.fromkeys(z['source'] for z in a):
   ns=[z['row'] for z in a if z['source']==src];refs.append(src+' dòng '+','.join(map(str,ns)))
  msg+=' ('+'; '.join(refs)+')';msgs.append(msg)
  results.append({'root':x['root'],**y,'discrepancy':discrepancy})
 c='Toàn bộ điều chỉnh là khoản tổng' if all(z=='Chỉ điều chỉnh khoản tổng' for z in cats) else 'Có trường hợp cần đối chiếu thêm'
 chain_counts[c]+=1;x['note']=(x['old_note']+'; ' if x['old_note'] else '')+'Rà điều chỉnh cũ: '+' | '.join(msgs)+'. Chuỗi bằng 0 ở bảng tổng không chứng minh đã giảm chi tiết hàng.'
assert len(chains)==237
payload={'roots':chains,'invoice_counts':dict(counts),'chain_counts':dict(chain_counts),'results':results,'hashes':hashes}
(h/'zero-adjustment-review.json').write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({'chains':len(chains),'adjustment_invoices':len(results),'invoice_counts':dict(counts),'chain_counts':dict(chain_counts),'detail_total_mismatch':sum(x['discrepancy'] for x in results),'missing_detail':sum(not x['detail'] for x in results),'common_descriptions':names.most_common(6)},ensure_ascii=False))
print('OTHER',[(x['root'],x['key'],x['category'],x['detail']) for x in results if x['category']!='Chỉ điều chỉnh khoản tổng'][:5])
print('EXAMPLE225',json.dumps(next(x for x in results if x['root']==['1C24TUV','00000225']),ensure_ascii=False))
