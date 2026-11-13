"""Đọc lại file đã xuất, so từng dòng và kiểm số tổng, kiểu dữ liệu, bộ lọc."""
from pathlib import Path
from zipfile import ZipFile
from collections import Counter
from decimal import Decimal
from datetime import datetime,date
import json,sys,hashlib
from openpyxl import load_workbook
sys.stdout.reconfigure(encoding='utf-8');O=Path(__file__).parent
p=O/'TONG_HOP_HD_PHAT_HANH_2026_CHO_DON_2024_2025.xlsx'
d=json.loads((O/'issued-2026-report-data.json').read_text(encoding='utf-8'))
with ZipFile(p) as z:assert z.testzip() is None
w=load_workbook(p,read_only=True,data_only=True)
assert w.sheetnames==['Tổng hợp','HĐ 2026','BHH và CT liên quan','Chuỗi thiếu BHH']
expected=[40161,42046,82207,*d['summary']['money_sums']]
assert [w['Tổng hợp'].cell(11,c).value for c in range(2,8)]==expected
assert w['Tổng hợp']['B18'].value==4022
order=[0,1,6,7,8,2,3,5,9,10,11,4,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29]
def equal(a,b,isdate=False):
 if isdate:return (a.date().isoformat() if isinstance(a,datetime) else a.isoformat() if isinstance(a,date) else a)==b
 return ('' if a is None else a)==('' if b is None else b)
money=[Decimal(0)]*3;missing=0;keys=set();years=Counter();count=0
for i,row in enumerate(w['HĐ 2026'].iter_rows(min_row=6,max_row=5+len(d['rows']),max_col=30,values_only=True)):
 target=d['rows'][i]
 for c,j in enumerate(order):assert equal(row[c],target[j],j in [8,14]),(i+6,c+1,row[c],target[j])
 assert isinstance(row[2],str) and isinstance(row[3],str) and len(row[3])==8
 key=row[2]+'/'+row[3];assert key not in keys;keys.add(key)
 assert isinstance(row[4],datetime) and row[4].year==2026
 years[row[1]]+=1
 for j in range(3):
  if row[8+j] is not None:assert isinstance(row[8+j],(int,float));money[j]+=Decimal(str(row[8+j]))
 if row[10] is None:missing+=1
 count+=1
assert count==82207 and dict(years)==d['summary']['year_counts']
assert [int(x) for x in money]==d['summary']['money_sums'] and missing==4022
for name,rows in [('BHH và CT liên quan',d['bhh_rows']),('Chuỗi thiếu BHH',d['unmapped_rows'])]:
 cnt=0
 for i,r in enumerate(w[name].iter_rows(min_row=6,max_row=5+len(rows),max_col=len(rows[0]),values_only=True)):
  for c,v in enumerate(r):
   expected_value=rows[i][c]
   if name.startswith('BHH') and c==9:expected_value='AMIS | '+expected_value.split(' | ',1)[1]
   assert equal(v,expected_value,c==(2 if name.startswith('BHH') else 3)),(name,i+6,c+1)
  cnt+=1
 assert cnt==len(rows)
errors=[]
for sh in w:
 for row in sh:
  for c in row:
   if c.data_type=='e':errors.append((sh.title,c.coordinate,c.value))
assert not errors
w.close()
# Kiểm tính năng native trực tiếp trong OOXML để không tải lại toàn bộ dữ liệu vào RAM.
with ZipFile(p) as z:
 for n in [2,3,4]:
  first=b'';tail=b''
  with z.open(f'xl/worksheets/sheet{n}.xml') as stream:
   while True:
    chunk=stream.read(1048576)
    if not chunk:break
    if not first:first=chunk[:16384]
    tail=(tail+chunk)[-32768:]
  assert b'<pane ' in first
  assert b'<autoFilter ' in tail or b'<tableParts' in tail
result={'passed':True,'invoices_verified':count,'bhh_references_verified':len(d['bhh_rows']),'unmapped_chains_verified':len(d['unmapped_rows']),'all_cells_compared':True,'totals':expected,'blank_payments':missing,'error_cells':0,'native_filters_and_freeze_panes':True,'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
(O/'issued-2026-final-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False),flush=True)
