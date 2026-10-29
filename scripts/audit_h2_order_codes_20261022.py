"""Đối chiếu mã đơn của bộ nguồn người dùng chỉ định theo khóa hóa đơn."""
import sys,json,hashlib
from pathlib import Path
from collections import defaultdict,Counter
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount

ROOT=Path(__file__).resolve().parents[1]
DIR=ROOT/'scripts/data/h2-order-codes-20261022'
SRC=ROOT/'data/Du lieu Meinvoice/T7_T12_2025_HD_da_su_dung'
mapping=defaultdict(list);files=[]
for p in sorted(SRC.glob('*.xlsx')):
 count=0
 for n,r in rows(p):
  if n<6 or amount(r.get('A')) is None:continue
  k=invoice_key(r.get('B'),r.get('C'))
  if not k:continue
  mapping[k].append(dict(code=(r.get('F') or '').strip(),date=str(invoice_day(r.get('D'))),money=[r.get(c) for c in ('N','Q','R')],source=str(p.relative_to(ROOT)),row=n));count+=1
 files.append(dict(path=str(p.relative_to(ROOT)),rows=count,sha256=hashlib.sha256(p.read_bytes()).hexdigest()))
payload=json.loads((ROOT/'scripts/data/h2-chain-readiness-20261022/workbook-input.json').read_text(encoding='utf-8'))
stats=Counter();plan=[]
for c in payload['chains']:
 for f,v in enumerate(c['invoices']):
  k=tuple(v['key']);found=mapping.get(k,[]);codes={x['code'] for x in found if x['code']}
  old=v.get('order') or '';flag='';new=old
  if len(codes)>1:flag='Nhiều mã khác nhau trong nguồn mới';stats['conflicting_new_codes']+=1
  elif len(codes)==1:
   new=next(iter(codes))
   if old and old!=new:flag='Mã nguồn mới khác mã cũ';stats['different_old_code']+=1
   if old!=new:stats['changed']+=1;stats['changed_F0' if f==0 else 'changed_related']+=1
  elif not found:stats['not_in_new_source']+=1
  else:stats['new_code_blank']+=1
  if found:
   if any(x['date']!=v['date'] or [amount(z) for z in x['money']]!=[amount(z) for z in v['money']] for x in found):
    flag='Ngày/tiền nguồn mới khác bảng chuỗi; giữ mã cũ';new=old;stats['invoice_field_conflict']+=1
  plan.append(dict(root=c['root'],f=f,key=v['key'],old=old,new=new,flag=flag,source=found))
DIR.mkdir(parents=True,exist_ok=True)
result=dict(files=files,unique_source_invoices=len(mapping),source_duplicate_keys=sum(len(v)>1 for v in mapping.values()),stats=dict(stats),plan=plan)
(DIR/'audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
sys.stdout.reconfigure(encoding='utf-8');print(json.dumps({k:v for k,v in result.items() if k!='plan'},ensure_ascii=False,indent=2))
