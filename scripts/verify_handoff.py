"""Kiểm tra chỉ đọc bộ bàn giao UVG; dùng đường dẫn theo vị trí repo."""
from pathlib import Path
from collections import defaultdict
from decimal import Decimal
import hashlib, json, openpyxl
ROOT=Path(__file__).resolve().parents[1]
def amt(v): return Decimal(str(v or 0))
manifest=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf-8-sig'))
for x in manifest:
 p=ROOT/x['path']; assert p.stat().st_size==x['bytes'],x['path']; assert hashlib.sha256(p.read_bytes()).hexdigest()==x['sha256'],x['path']
result={'source_files_verified':len(manifest)}
for name,expected,lines in [('HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx',8404,13543),('HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',1990,41268)]:
 w=openpyxl.load_workbook(ROOT/'outputs'/name,read_only=True,data_only=True)
 seq=set(); roots=set(); totals=defaultdict(lambda:[Decimal(0)]*3); count=promo=0
 for v in w['Import dieu chinh'].iter_rows(min_row=10,max_col=24,values_only=True):
  if v[0] is None: continue
  count+=1; seq.add(v[0]); key=(v[9],v[10]); roots.add(key)
  assert isinstance(v[10],str) and len(v[10])==8,key
  assert v[2]==v[5],(key,'Tên khách hàng'); assert v[16]!=0,(key,'SL0')
  assert v[23]==(1 if v[18]==0 else None),(key,'Khuyến mãi'); promo+=v[23]==1
  if v[19] is not None:
   vals=[amt(v[i]) for i in [19,21,22]]; assert vals[0]+vals[1]==vals[2],key
   for i,a in enumerate(vals): totals[key][i]+=a
 assert count==lines,(name,count); assert seq==set(range(1,expected+1)),name
 if expected==8404:
  excluded={(v[0],v[1]) for v in w['Kiem tra SL 0'].iter_rows(min_row=2,values_only=True) if isinstance(v[0],str) and v[0].startswith('1C')}; assert len(excluded)==6 and not roots.intersection(excluded)
 else:
  old={(v[0],v[1]):[amt(v[i]) for i in [24,25,26]] for v in w['Dieu chinh'].iter_rows(min_row=3,max_col=27,values_only=True) if v[0]}
  assert len(roots)==986
  for key,vals in totals.items(): assert all(a+b==0 for a,b in zip(old[key],vals)),key
 result[name]={'invoices':len(seq),'roots':len(roots),'lines':count,'promo_lines':promo,'sheets':w.sheetnames}
 w.close()
result['passed']=True
print(json.dumps(result,ensure_ascii=False,indent=2))