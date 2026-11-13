"""Kiểm dữ liệu giữ lại sau khi rút gọn, gồm khóa HĐ, BHH và số tiền từng dòng."""
from pathlib import Path
from zipfile import ZipFile
from datetime import datetime
from collections import Counter
from decimal import Decimal
from openpyxl import load_workbook
import json,hashlib,sys
sys.stdout.reconfigure(encoding='utf-8');O=Path(__file__).parent
d=json.loads((O/'issued-2026-report-data.json').read_text(encoding='utf-8'))
p=O/'TONG_HOP_HD_2026_CHO_DON_2024_2025_RUT_GON.xlsx'
with ZipFile(p) as z:assert z.testzip() is None
w=load_workbook(p,read_only=True,data_only=True)
assert w.sheetnames==['Tổng hợp','HĐ 2026','BHH và CT liên quan','Chuỗi thiếu BHH']
assert [w['Tổng hợp'].cell(11,c).value for c in range(2,8)]==[40161,42046,82207,*d['summary']['money_sums']]
assert w['Tổng hợp']['B17'].value==4022
assert [w[n].max_column for n in w.sheetnames]==[7,13,7,7]
keys=set();money=[Decimal(0)]*3;years=Counter();types=Counter();blanks=0;count=0;errors=[]
for i,row in enumerate(w['HĐ 2026'].iter_rows(min_row=6,max_col=13)):
 r=[c.value for c in row];old=d['rows'][i]
 assert [r[j] for j in [0,1,2,3]]==[old[j] for j in [0,1,6,7]]
 assert isinstance(r[2],str) and isinstance(r[3],str) and len(r[3])==8
 assert isinstance(r[4],datetime) and r[4].date().isoformat()==old[8]
 key=r[2]+'/'+r[3];assert key not in keys;keys.add(key)
 docs=set(x for s in [old[2],old[3]] for x in str(s or '').split('\n') if x)
 assert set(str(r[5] or '').split('\n'))- {''}==docs
 assert r[6]==old[5] and (r[7] or '')==(old[4] or '')
 for j in range(3):
  assert r[8+j]==old[9+j]
  if r[8+j] is not None:money[j]+=Decimal(str(r[8+j]))
 if r[10] is None:blanks+=1
 expected_parent=old[15]+'/'+old[16] if old[15] and old[16] else ''
 assert (r[11] or '')==expected_parent and r[11]!=key
 if not docs:assert 'Chưa xác định' in r[12]
 if 'Mã đơn' in old[23]:assert 'Nối theo mã đơn' in r[12]
 for c in row:
  if c.data_type=='e':errors.append((c.coordinate,c.value))
 years[r[1]]+=1;types[r[6]]+=1;count+=1
assert count==82207 and dict(years)==d['summary']['year_counts'] and dict(types)==d['summary']['purpose_counts']
assert [int(x) for x in money]==d['summary']['money_sums'] and blanks==4022
ref_count=0
for i,row in enumerate(w['BHH và CT liên quan'].iter_rows(min_row=6,max_col=7)):
 r=[c.value for c in row];old=d['bhh_rows'][i]
 assert r[:2]==old[:2]
 assert r[2].date().isoformat()==old[2]
 assert [(r[j] or '') for j in [3,4,5]]==[(old[j] or '') for j in [8,11,12]]
 new=(r[5] or '').split('\n');assert all(k in keys for k in new)
 if old[4]:assert old[4] in new or old[4] in str(r[6])
 if old[10]:assert old[10] in r[6]
 for c in row:
  if c.data_type=='e':errors.append((c.coordinate,c.value))
 ref_count+=1
assert ref_count==65726
un_count=0
for i,row in enumerate(w['Chuỗi thiếu BHH'].iter_rows(min_row=6,max_col=7)):
 r=[c.value for c in row];old=d['unmapped_rows'][i][:7]
 assert r[:3]==old[:3] and r[3].date().isoformat()==old[3] and r[4:]==old[4:]
 assert all(k in keys for k in r[4].split('\n'))
 for c in row:
  if c.data_type=='e':errors.append((c.coordinate,c.value))
 un_count+=1
assert un_count==2989 and not errors
w.close()
wf=load_workbook(p,read_only=True,data_only=False)
for row in wf['Tổng hợp'].iter_rows(min_row=6,max_row=10,min_col=2,max_col=7):
 for j,c in enumerate(row):
  if j!=2:assert "'HĐ 2026'!$G$6:" in c.value,(c.coordinate,c.value)
assert wf['Tổng hợp']['B17'].value=="=COUNTBLANK('HĐ 2026'!K6:K82212)"
wf.close()
with ZipFile(p) as z:
 for n,ref in [(2,'A5:M82212'),(3,'A5:G65731'),(4,'A5:G2994')]:
  first=b'';tail=b''
  with z.open(f'xl/worksheets/sheet{n}.xml') as stream:
   while True:
    chunk=stream.read(1048576)
    if not chunk:break
    if not first:first=chunk[:16384]
    tail=(tail+chunk)[-32768:]
  assert b'<pane ' in first and f'<autoFilter ref="{ref}"'.encode() in tail
original=O/'TONG_HOP_HD_PHAT_HANH_2026_CHO_DON_2024_2025.xlsx'
assert hashlib.sha256(original.read_bytes()).hexdigest()=='487bb64447227e0dd87283aee097c25aaf90d4eebfdd07ea8deb6b51b6f1b2ee'
result={'passed':True,'invoices':count,'document_rows':ref_count,'unmapped_chains':un_count,'columns_before_after':{'HĐ 2026':[30,13],'BHH và CT liên quan':[13,7],'Chuỗi thiếu BHH':[12,7]},'money_and_mapping_unchanged':True,'blank_payments':blanks,'error_cells':0,'filters_and_freeze_panes':True,'summary_formula_references_checked':True,'original_file_unchanged':True,'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
(O/'issued-2026-compact-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False),flush=True)
