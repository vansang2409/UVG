import json,sys,hashlib
from pathlib import Path
from collections import Counter
from decimal import Decimal
r=Path(r'D:\UVG');h=r/'scripts/data/untreated';sys.path.insert(0,str(r/'scripts'))
from prepare_invoice_reports import rows
p=h/'untreated.json';d=json.loads(p.read_text(encoding='utf8'));keys={tuple(x['root']) for x in d['groups']}
counts={k:Counter() for k in keys}; refs={k:[] for k in keys}; dates={k:set() for k in keys}
for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
 f=r/'data/Du lieu Meinvoice'/name;mt='_MTT' in name
 for n,v in rows(f):
  key=(v.get('F' if mt else 'D'),v.get('B'))
  if key not in keys:continue
  c=counts[key];c['lines']+=1;a=v.get('W' if mt else 'P');pay=v.get('AB' if mt else 'U');flag=v.get('AD' if mt else 'W')
  c['missing_goods' if a in (None,'') else 'zero_goods' if Decimal(a)==0 else 'nonzero_goods']+=1
  c['missing_payment' if pay in (None,'') else 'zero_payment' if Decimal(pay)==0 else 'nonzero_payment']+=1
  if flag not in (None,'','0'):c['promo_flag']+=1
  refs[key].append((name,n))
for g in d['groups']:
 k=tuple(g['root']);c=counts[k];assert c['lines']>0;g['line_counts']=dict(c)
 src=[]
 for name in dict.fromkeys(x[0] for x in refs[k]):
  ns=[n for fn,n in refs[k] if fn==name];src.append(name+' | dòng '+(str(ns[0]) if len(ns)==1 else str(min(ns))+'–'+str(max(ns))))
 g['detail_source']='; '.join(src)
 m=g['invoices'][0][4:7]
 if m[2] is None and c['zero_payment']==c['lines']:
  g['notes']+='; Chi tiết có thanh toán 0 tại tất cả '+str(c['lines'])+' dòng; bảng tổng vẫn trống, chưa tự bổ sung 0'
 g['money_class']='Thiếu số tiền bảng tổng' if any(v is None for v in m) else 'Cả 3 khoản bằng 0' if all(v==0 for v in m) else 'Có giá trị khác 0'
 d['existing_report_hashes']={fn:hashlib.sha256((r/'outputs'/fn).read_bytes()).hexdigest() for fn in ['HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx','NGOAI_LE_F0_2024_T6_2025.xlsx']}
p.write_text(json.dumps(d,ensure_ascii=False),encoding='utf8')
print('EXCLUSIONS',dict(Counter(x['reason'] for x in d['exclusions'])))
print('MISSING',[(g['root'],g['invoices'][0][4:7],g['line_counts']) for g in d['groups'] if g['money_class']=='Thiếu số tiền bảng tổng'])
print('TOTAL_DETAIL_LINES',sum(c['lines'] for c in counts.values()),'ROOTS_WITH_ZERO_LINES',sum(c['zero_goods']>0 for c in counts.values()))
