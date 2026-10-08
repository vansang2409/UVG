from pathlib import Path
import json,hashlib,shutil,openpyxl
r=Path(r'D:\UVG');h=r/'scripts/data/check-review';p=r/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx'
d=json.loads((h/'zero-detail-review.json').read_text(encoding='utf8'));expected={x['row']:x['note'] for x in d['roots']}
a=openpyxl.load_workbook(h/'before-zero-notes.xlsx',data_only=False);b=openpyxl.load_workbook(h/'with-zero-notes.xlsx',data_only=False)
assert a.sheetnames==b.sheetnames
changed=[]
for s in a:
 t=b[s.title];assert (s.max_row,s.max_column)==(t.max_row,t.max_column)
 for row in s:
  for v in row:
   u=t.cell(v.row,v.column)
   value=expected[v.row] if s.title=='Dieu chinh' and v.column==28 and v.row in expected else v.value
   assert u.value==value,(s.title,v.coordinate)
   if v.value!=u.value:changed.append((s.title,v.coordinate))
   assert v.number_format==u.number_format
 assert s.freeze_panes==t.freeze_panes
 assert [(x.name,x.ref) for x in s.tables.values()]==[(x.name,x.ref) for x in t.tables.values()]
assert len(changed)==237 and all(c=='Dieu chinh' and pos.startswith('AB') for c,pos in changed)
assert all(hashlib.sha256((r/k).read_bytes()).hexdigest()==v for k,v in d['hashes'].items())
assert all(x['old_note'] is None or x['note'].startswith(x['old_note']+'; ') for x in d['roots'])
assert hashlib.sha256(p.read_bytes()).digest()==hashlib.sha256((h/'before-zero-notes.xlsx').read_bytes()).digest(),'File changed during edit'
shutil.copyfile(h/'with-zero-notes.xlsx',p)
assert hashlib.sha256(p.read_bytes()).digest()==hashlib.sha256((h/'with-zero-notes.xlsx').read_bytes()).digest()
print('PASS: 237 AB notes added; original notes preserved; other values/formulas/number formats/tables/freeze panes unchanged; sources unchanged; saved original workbook')
print('DETAIL',len(d['roots']),d['counts'],'ZERO LINES',sum(a['goods']==0 for x in d['roots'] for a in x['detail']))
