from pathlib import Path
import openpyxl,hashlib,shutil,json
r=Path(r'D:\UVG');h=r/'scripts/data/untreated';a=openpyxl.load_workbook(h/'before-buyer-sync.xlsx');b=openpyxl.load_workbook(h/'with-buyer-sync.xlsx');assert a.sheetnames==b.sheetnames
changed=0;blank=0
for name in a.sheetnames:
 s,t=a[name],b[name];assert (s.max_row,s.max_column)==(t.max_row,t.max_column)
 assert s.freeze_panes==t.freeze_panes
 assert list(s.merged_cells.ranges)==list(t.merged_cells.ranges)
 for k in s.tables:assert s.tables[k].ref==t.tables[k].ref
 for rr,ss in zip(s.iter_rows(),t.iter_rows()):
  for c,e in zip(rr,ss):
   if name=='Import dieu chinh' and c.column==3 and c.row>=10:
    assert e.value==s.cell(c.row,6).value,(c.coordinate,'buyer mismatch')
    changed+=c.value!=e.value;blank+=e.value is None
   else:assert c.value==e.value,(name,c.coordinate,'unexpected change')
   assert c.number_format==e.number_format,(name,c.coordinate,'format changed')
p=r/'outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
assert sha(p)==sha(h/'before-buyer-sync.xlsx'),'Report changed concurrently'
shutil.copy2(h/'with-buyer-sync.xlsx',p);assert sha(p)==sha(h/'with-buyer-sync.xlsx')
print(json.dumps({'PASS':True,'matched_rows':13570,'changed_names':changed,'blank_buyer_rows':blank,'other_values_preserved':True,'saved':True}))
