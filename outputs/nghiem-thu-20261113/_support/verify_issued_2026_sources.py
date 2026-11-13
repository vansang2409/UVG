"""Kiểm độc lập danh sách HĐ 2026 với bảng kê ZIP và các liên kết gốc."""
from pathlib import Path
from zipfile import ZipFile
from decimal import Decimal
from collections import Counter,defaultdict
import json,io,sys,hashlib
sys.stdout.reconfigure(encoding='utf-8')
R=Path('D:/UVG');O=Path(__file__).parent;sys.path.insert(0,str(R/'scripts'))
from prepare_invoice_reports import rows,invoice_key,invoice_day
d=json.loads((O/'issued-2026-report-data.json').read_text(encoding='utf-8'))
a=json.loads((O/'issued-2026-selection-audit.json').read_text(encoding='utf-8'))
expected={r[6]+'/'+r[7]:r for r in d['rows']};assert len(expected)==len(d['rows'])==82207
def n(v):return None if v in (None,'') else Decimal(str(v))
seen=set();full_count=0;full_2026=0;blank_header_counts=Counter();raw={};manifest=[]
p=next((R/'data/1311/meinvoice').glob('Bang_ke_hoa_don_da_su_dung_*.zip'))
with ZipFile(p) as z:
 for member in z.namelist():
  if member.endswith('/'):continue
  for rn,r in rows(io.BytesIO(z.read(member))):
   if rn<6 or not str(r.get('A') or '').isdigit():continue
   full_count+=1;k='/'.join(invoice_key(r.get('B'),r.get('C')));dt=invoice_day(r.get('D')).isoformat()
   if dt.startswith('2026'):full_2026+=1
   if k not in expected:continue
   assert k not in seen;seen.add(k);x=expected[k]
   assert x[8]==dt and dt.startswith('2026'),(k,'date')
   assert x[28]==member+'!'+str(rn),(k,'source')
   assert x[24]==str(r.get('U') or '') and x[25]==str(r.get('S') or ''),(k,'status')
   amounts=[n(r.get(c)) for c in ['P','Q','R']];raw[k]=amounts
   for j,v in enumerate(amounts):
    if v is not None:assert v==n(x[9+j]),(k,'money',j,v,x[9+j])
    else:blank_header_counts[j]+=1
  print('READ',member,'selected',len(seen),flush=True)
assert set(expected)==seen and full_count==235263 and full_2026==186389
manifest.append({'file':p.name,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
rels={};relationship_rows=0;source_rel_money=defaultdict(lambda:defaultdict(set))
for p in (R/'data/1311/meinvoice').glob('*.zip'):
 kind='adjustment' if 'dieu_chinh' in p.name else 'replacement' if 'thay_the' in p.name else None
 if not kind:continue
 manifest.append({'file':p.name,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
 with ZipFile(p) as z:
  for member in z.namelist():
   if member.endswith('/'):continue
   for rn,r in rows(io.BytesIO(z.read(member))):
    if rn<7 or not str(r.get('A') or '').isdigit():continue
    old='/'.join(invoice_key(r.get('B'),r.get('D')));new='/'.join(invoice_key(r.get('M'),r.get('O')))
    rels[old,new]=kind;relationship_rows+=1
    for k,cols in [(old,['J','K','L']),(new,['U','V','W'])]:
     if k not in expected:continue
     for j,col in enumerate(cols):
      v=n(r.get(col))
      if v is not None:source_rel_money[k][j].add(v)
assert relationship_rows==18790
# Các tiền bổ sung ở ô bảng tổng trống phải có trong liên kết hoặc chi tiết cùng kỳ.
detail_needed={k for k,x in expected.items() if any(raw[k][j] is None and x[9+j] is not None and n(x[9+j]) not in source_rel_money[k][j] for j in range(3))}
detail_sums=defaultdict(lambda:[Decimal(0)]*3);detail_known=defaultdict(lambda:[0]*3)
if detail_needed:
 p=next((R/'data/1311/meinvoice').glob('*chi_tiet*.zip'));manifest.append({'file':p.name,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
 with ZipFile(p) as z:
  for member in z.namelist():
   if member.endswith('/'):continue
   for rn,r in rows(io.BytesIO(z.read(member))):
    if rn<6:continue
    kk=invoice_key(r.get('B'),r.get('C'))
    if not kk:continue
    k='/'.join(kk)
    if k not in detail_needed:continue
    for j,col in enumerate(['P','R','S']):
     v=n(r.get(col))
     if v is not None:detail_sums[k][j]+=v;detail_known[k][j]+=1
 for k in detail_needed:
  for j in range(3):
   if raw[k][j] is None and expected[k][9+j] is not None and n(expected[k][9+j]) not in source_rel_money[k][j]:
    assert detail_known[k][j] and detail_sums[k][j]==n(expected[k][9+j]),(k,'fallback',j)
cache_edges=json.loads((O/'edges.json').read_text(encoding='utf-8'))
assert {(e['old'],e['new']):e['kind'] for e in cache_edges}==rels
am=json.loads((O/'amis.json').read_text(encoding='utf-8'));allowed={x['doc'] for x in am.values()}
assert all(not x[2] or set(x[2].split('\n'))<=allowed for x in d['rows'])
assert len({(r[0],r[1]) for r in d['bhh_rows']})==len(d['bhh_rows'])
for p,sha in d['summary']['source_amis_sha256'].items():assert hashlib.sha256(Path(p).read_bytes()).hexdigest()==sha
verified={'passed':True,'header_rows_read':full_count,'headers_2026_in_source':full_2026,'selected_unique_invoices_verified':len(seen),'relationship_rows_verified':relationship_rows,'header_blank_cells':dict(blank_header_counts),'detail_fallback_invoices_verified':len(detail_needed),'mapped_documents':len(d['bhh_rows']),'sources_unchanged':True,'manifest':manifest}
(O/'issued-2026-source-verification.json').write_text(json.dumps(verified,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(verified,ensure_ascii=False,indent=2),flush=True)
