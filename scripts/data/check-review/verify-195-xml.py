"""Xác minh độc lập và chỉ thay file chính khi toàn bộ thay đổi đúng phạm vi."""
from pathlib import Path
from datetime import datetime
from decimal import Decimal as D
from collections import defaultdict
import json,hashlib,shutil,openpyxl,copy
h=Path(__file__).resolve().parent;root=h.parents[2]
d=json.loads((h/'195-xml-data.json').read_text(encoding='utf-8'))
p=root/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx'
sha=lambda f:hashlib.sha256(f.read_bytes()).hexdigest()
assert sha(p)==d['base_sha256'],'Workbook changed concurrently'
a=openpyxl.load_workbook(h/'before-195-xml.xlsx');b=openpyxl.load_workbook(h/'with-195-xml.xlsx');cache=openpyxl.load_workbook(h/'with-195-xml.xlsx',read_only=True,data_only=True)
allowed={('Dieu chinh',f'{col}{x["row"]}') for x in d['changes'] for col in x['values']}
allowed|={('Import dieu chinh',x) for x in ['A3','A4','A7']}
allowed|={('Huong dan',f'A{i}') for i in range(13,17)}
assert a.sheetnames==b.sheetnames
changed=[];checked=0
acache={};bcache={}
def style(c,cache):
 if c.style_id not in cache:cache[c.style_id]=tuple(copy.copy(getattr(c,k)) for k in ['font','fill','border','alignment','protection'])
 return cache[c.style_id]
for name in a.sheetnames:
 s,t=a[name],b[name]
 assert s.freeze_panes==t.freeze_panes and set(map(str,s.merged_cells.ranges))==set(map(str,t.merged_cells.ranges)),name
 assert {k:s.tables[k].ref for k in s.tables}=={k:t.tables[k].ref for k in t.tables},name
 assert str(s.data_validations)==str(t.data_validations),name
 for rr in s.iter_rows():
  for cell in rr:
   other=t[cell.coordinate];checked+=1
   if cell.value!=other.value:
    assert (name,cell.coordinate) in allowed,(name,cell.coordinate,cell.value,other.value)
    changed.append((name,cell.coordinate))
   assert cell.number_format==other.number_format,(name,cell.coordinate,'number format')
   left=style(cell,acache);right=style(other,bcache)
   assert left[:3]==right[:3] and left[4]==right[4],(name,cell.coordinate,'style')
   if (name,cell.coordinate) not in allowed:assert left[3]==right[3],(name,cell.coordinate,'alignment')
 for rownum,dim in s.row_dimensions.items():
  if name!='Dieu chinh' or rownum not in {x['row'] for x in d['changes']}:
   assert dim.height==t.row_dimensions[rownum].height,(name,rownum,'height')
assert set(changed)==allowed,(len(changed),len(allowed),allowed-set(changed))
for x in d['changes']:
 for col,val in x['values'].items():
  expected=datetime(2026,10,9) if col in ['AE','AK'] else val
  assert b['Dieu chinh'][f'{col}{x["row"]}'].value==expected
s=b['Import dieu chinh'];byseq=defaultdict(list);totals=defaultdict(lambda:[D(0)]*3)
for n,(row,expected,meta) in enumerate(zip(s.iter_rows(min_row=d['start_row'],max_row=d['last_row'],max_col=24,values_only=True),d['rows'],d['meta']),d['start_row']):
 expected=list(expected);expected[1]=datetime(2026,10,9)
 assert list(row)==expected,(n,'row mismatch')
 key=tuple(row[9:11]);assert row[2]==row[5] and isinstance(row[10],str) and len(row[10])==8
 assert row[23]==(1 if row[18]==0 else None)
 assert row[16] is None if meta['kind']=='reverse' else row[16] is not None and row[16]<0
 assert row[17] is None if meta['kind']=='reverse' else row[17]>=0
 byseq[row[0]].append(row)
 if meta['first']:
  money=[D(str(row[c])) for c in [19,21,22]];assert money[0]+money[1]==money[2]
  for i,v in enumerate(money):totals[key][i]+=v
 else:assert all(row[c] is None for c in [19,20,21,22])
assert set(byseq)==set(range(1601,1991)) and len(totals)==195
for seq,rs in byseq.items():assert sum(D(str(v[18])) for v in rs)==D(str(rs[0][19])),seq
for x in d['chains']:
 k=tuple(x['root']);assert totals[k]==[D(0)]*3
 n=x['row'];r=b['Dieu chinh']
 assert [r[f'{col}{n}'].value for col in ['AF','AG','AH']]==x['original_money']
 assert [r[f'{col}{n}'].value for col in ['AL','AM','AN']]==[-v for v in x['original_money']]
 xml=root/'data/Bo sung/meInvoice/2026-10-09-237-XML'/('_'.join(x['adjustment'])+'.xml')
 assert sha(xml)==x['xml_sha256']
cd={n:v for n,v in enumerate(cache['Dieu chinh'].iter_rows(min_row=3,max_col=55,values_only=True),3)}
for x in d['chains']:assert list(cd[x['row']][52:55])==[0,0,0]
errors=[]
for sheet in cache:
 for rr in sheet.iter_rows():
  for cell in rr:
   if cell.data_type=='e':errors.append((sheet.title,cell.coordinate,cell.value))
assert not errors,errors[:5]
assert s.max_row==d['last_row'] and s.max_column==24
cache.close()
result={'passed':True,'old_cells_checked':checked,'allowed_changed_cells':len(changed),'new_chains':195,'new_invoices':390,'new_lines':d['added_lines'],'total_invoices':1990,'total_roots':986,'total_lines':d['total_lines'],'all_195_cumulative_zero':True,'old_sheet_values_formulas_styles_preserved':True,'xml_hashes_verified':195,'base_sha256':d['base_sha256'],'output_sha256':sha(h/'with-195-xml.xlsx')}
(h/'195-xml-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
shutil.copy2(h/'with-195-xml.xlsx',p);assert sha(p)==result['output_sha256']
print(json.dumps(result,ensure_ascii=False,indent=2))
