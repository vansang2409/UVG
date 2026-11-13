"""Kiểm tab mới với quan hệ gốc và bảo toàn nguyên bảng HĐ 2026 đã rút gọn."""
from pathlib import Path
from zipfile import ZipFile
from collections import defaultdict
from functools import lru_cache
from openpyxl import load_workbook
from datetime import datetime
from itertools import zip_longest
from decimal import Decimal
import json,re,hashlib,sys,unicodedata
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8');O=Path(__file__).parent
def read(n):return json.loads((O/n).read_text(encoding='utf-8'))
def digest_member(path,name):
 h=hashlib.sha256()
 with ZipFile(path) as z:
  with z.open(name) as f:
   while True:
    b=f.read(1048576)
    if not b:break
    h.update(b)
 return h.hexdigest()
results=O.parent if O.name=='_support' else O
src=O/'TONG_HOP_HD_2026_CHO_DON_2024_2025_RUT_GON.xlsx';p=results/'TONG_HOP_HD_2026_CHO_DON_2024_2025_CHUOI.xlsx'
assert hashlib.sha256(src.read_bytes()).hexdigest()=='403bc88ad1d180de06caa56a6077d2cff6f1c2b4b12906838e7c4589583bc88d'
with ZipFile(p) as z:assert z.testzip() is None
d=read('issued-2026-report-data.json');expected=read('issued-2026-chain-view.json');audit=read('issued-2026-chain-view-audit.json');ledger=read('ledger.json')
ed=read('edges.json');rev=read('reverse-chain-audit.json');ed+=rev['recovered_edges']
resolved={i['key']:i['resolved'] for c in rev['components'] for i in c['invoices']};del rev
replaced={e['old'] for e in ed if e['kind']=='replacement'}
def norm(s):return ''.join(c for c in unicodedata.normalize('NFD',str(s).lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
parent={};children=defaultdict(set);kind={}
for e in ed:
 parent[e['new']]=e['old'];children[e['old']].add(e['new']);kind[e['old'],e['new']]=e['kind']
@lru_cache(None)
def root(k):return root(parent[k]) if k in parent else k
refs={str(r[0])+'/'+r[1]:r for r in d['bhh_rows']}
target=set(audit['invoice_keys']);assert len(target)==82207
w=load_workbook(p,read_only=True,data_only=True)
assert w.sheetnames==['HĐ 2026','Chuỗi F0 F1 F2']
assert [w[n].max_column for n in w.sheetnames]==[13,17]
source_book=load_workbook(src,read_only=True,data_only=True)
invoice_count=0;main_money=[Decimal(0)]*3;main_blanks=0
for a,b in zip_longest(w['HĐ 2026'].iter_rows(min_row=6,max_col=13,values_only=True),source_book['HĐ 2026'].iter_rows(min_row=6,max_col=13,values_only=True)):
 assert a is not None and a==b,'Bảng HĐ 2026 thay đổi'
 assert isinstance(a[3],str) and len(a[3])==8
 for j in range(3):
  if a[8+j] is not None:main_money[j]+=Decimal(str(a[8+j]))
 if a[10] is None:main_blanks+=1
 invoice_count+=1
assert invoice_count==82207 and [int(v) for v in main_money]==d['summary']['money_sums'] and main_blanks==4022
source_book.close();print('UNCHANGED MAIN',invoice_count,flush=True)
assert [c.value for c in w['Chuỗi F0 F1 F2'][6]]==expected['headers']
regex=re.compile(r'\b1[A-Z0-9]+/\d{8}\b');covered=set();documents=set();unknown=set();count=0
events=defaultdict(lambda:defaultdict(list));phase_sums=defaultdict(lambda:[Decimal(0)]*4);missing_payments=0;groups=set();final_values={};formula_cells={};payment_cells=defaultdict(dict)
palette=[('305E8C','EAF2FA','DDEAF8'),('9A6700','FFF5DF','FFEDC4'),('6B4C9A','F4EEFA','E9DFF4'),('2F6F46','EAF5EC','D9EDDE')]
for phase in range(4):assert w['Chuỗi F0 F1 F2'].cell(5,4+phase*3).fill.fgColor.rgb[-6:]==palette[phase][0]
for i,cells in enumerate(w['Chuỗi F0 F1 F2'].iter_rows(min_row=7,max_col=17)):
 r=[c.value for c in cells];exp=expected['rows'][i];gid=r[0];rec=audit['records'][gid-1];groups.add(gid)
 for col,(actual,val) in enumerate(zip(r,exp)):
  if col in [4,7,10,13] and val is not None:assert isinstance(actual,datetime) and actual.date().isoformat()==val,(i+7,col)
  else:assert ('' if actual is None else actual)==('' if val is None else val),(i+7,col)
 assert r[1]==rec['year'] and (r[2] or '')==rec['doc']
 if gid not in final_values:
  assert r[15]==rec['final_value'];final_values[gid]=r[15]
  formula_cells['P'+str(i+7)]=rec['final_formula'][1:]
  assert cells[15].font.bold and cells[15].fill.fgColor.rgb[-6:]=='E4F0EF'
 else:assert r[15] is None
 for col in range(3,15):
  ismissing=col in [5,8,11,14] and r[col-2] and r[col] is None
  assert cells[col].fill.fgColor.rgb[-6:]==('FFF1D6' if ismissing else palette[(col-3)//3][1+int(gid%2==1)])
 for phase in range(4):
  c=3+phase*3;label=r[c];date=r[c+1];amount=r[c+2]
  if not label:assert date is None and amount is None;continue
  keys=regex.findall(label);assert len(keys)==1;k=keys[0]
  assert k in ledger and k not in {x for ks in events[gid].values() for x in ks}
  assert date.date().isoformat()==ledger[k]['day'] and amount==ledger[k]['money'][2]
  if amount is None:
   assert cells[c+2].fill.fgColor.rgb[-6:]=='FFF1D6';missing_payments+=1
  else:
   assert isinstance(amount,(int,float))
   assert cells[c+2].alignment.indent==1 and cells[c+2].border.right.style=='thin'
   phase_sums[gid][phase]+=Decimal(str(amount))
  if phase==0:assert k==rec['root'] and k not in parent
  elif phase==1:assert k in rec['nodes'] and k!=rec['root'] and ledger[k]['day'][:4]<'2026'
  elif phase==2:assert k in rec['nodes'] and k!=rec['root'] and ledger[k]['day'].startswith('2026-')
  else:assert k in rec['reissued'] and k not in parent and ledger[k]['day'].startswith('2026-')
  if phase in [1,2]:assert k+(' (ĐC)' if kind[parent[k],k]=='adjustment' else ' (TT)') in label
  events[gid][phase].append(k)
  payment_cells[gid][k]=['F','I','L','O'][phase]+str(i+7)
  if ledger[k]['day'].startswith('2026-'):covered.add(k)
 if r[2]:documents.add(r[1]+'/'+r[2])
 else:unknown.add(rec['root'])
 for c in cells:assert c.data_type!='e'
 count+=1
assert count==expected['summary']['rows']==82817 and covered==target and documents==set(refs)
assert groups==set(range(1,len(audit['records'])+1))
for gid,rec in enumerate(audit['records'],1):
 for phase in range(4):assert events[gid][phase]==rec['phases'][str(phase)]
 nodes=set(rec['nodes'])
 assert all(children[k]<=nodes for k in nodes)
 assert set(x for ks in events[gid].values() for x in ks)==nodes|set(rec['reissued'])
 # Kiểm độc lập số tiền và tham chiếu công thức cuối với nguồn còn hiệu lực.
 active=[];value=Decimal(0)
 for k in nodes|set(rec['reissued']):
  if k in replaced or 'bi thay the' in norm(ledger[k]['status']) or 'huy' in norm(ledger[k]['status']) or 'khong hop le' in norm(ledger[k]['cqt_status']):continue
  active.append(k);pay=ledger[k]['money'][2]
  if pay is None:
   g,v,rp=resolved.get(k,ledger[k]['money'])
   assert rp==0 or (rp is None and g is not None and v is not None and g+v==0)
   pay=0
  value+=Decimal(str(pay))
 assert value==final_values[gid],(gid,value,final_values[gid])
 assert set(rec['final_active'])==set(active)
 assert set(re.findall(r'[FILO]\d+',rec['final_formula']))=={payment_cells[gid][k] for k in active}
for gid,amount in [(46137,2639000),(46139,1479000)]:
 assert sum(phase_sums[gid][:3])==0 and phase_sums[gid][3]==amount
assert sum(phase_sums[34968][:3])==0
assert [final_values[k] for k in [34968,46137,46139]]==[0,2639000,1479000]
assert unknown=={r[2] for r in d['unmapped_rows']} and len(unknown)==2989
w.close()
with ZipFile(p) as z:
 first=b'';tail=b''
 with z.open('xl/worksheets/sheet2.xml') as f:
  while True:
   b=f.read(1048576)
   if not b:break
   if not first:first=b[:16384]
   tail=(tail+b)[-32768:]
 assert b'<pane ' in first and b'<autoFilter ref="A6:Q82823"' in tail
 for ref in ['A5:C5','D5:F5','G5:I5','J5:L5','M5:O5','P5:Q5']:assert f'<mergeCell ref="{ref}"'.encode() in tail
 ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};actual_formulas={}
 with z.open('xl/worksheets/sheet2.xml') as f:
  for _,row in etree.iterparse(f,events=('end',),tag='{'+ns['s']+'}row'):
   for cell in row:
    form=cell.find('s:f',ns)
    if form is not None:actual_formulas[cell.get('r')]=form.text
   row.clear()
   while row.getprevious() is not None:del row.getparent()[0]
 assert actual_formulas==formula_cells
result={'passed':True,'updated_in_place':True,'sheets':['HĐ 2026','Chuỗi F0 F1 F2'],'removed_sheets':['Tổng hợp','BHH và CT liên quan','Chuỗi thiếu BHH'],'main_values_dates_amounts_identical':True,'invoice_count':invoice_count,'chain_display_rows':count,'chain_groups':len(groups),'documents':len(documents),'phase_count':4,'date_and_payment_for_each_invoice_verified':True,'unknown_document_chains_retained':len(unknown),'all_2026_invoices_represented':True,'business_phases_verified':True,'examples_zeroed_and_reissued_verified':True,'blank_payment_events_marked':missing_payments,'final_formulas_verified':len(actual_formulas),'phase_colors_verified':True,'error_cells':0,'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
(O/'issued-2026-chain-book-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False),flush=True)
