import sys,json,collections,openpyxl
from pathlib import Path
from io import BytesIO
sys.stdout.reconfigure(encoding='utf8');r=Path(r'D:\UVG');w=openpyxl.load_workbook(r/'outputs/HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx',data_only=True)
roots={tuple(v[:2]):v for v in w['Chua dieu chinh'].iter_rows(min_row=3,values_only=True)};lines=collections.defaultdict(list)
for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
 s=openpyxl.load_workbook(BytesIO((r/'data/Du lieu Meinvoice'/name).read_bytes()),read_only=True,data_only=True).worksheets[0];mt='_MTT' in name
 for n,v in enumerate(s.iter_rows(values_only=True),1):
  k=(v[5] if mt else v[3],v[1])
  if k in roots:lines[k].append({'row':n,'source':name,'v':list(v),'mt':mt})
c=collections.Counter();bad=[]
for k,v in roots.items():
 ls=lines[k]; rates={str(x['v'][25 if x['mt'] else 18]) for x in ls}
 def sums(i,j):return sum(float(x['v'][j if x['mt'] else i] or 0) for x in ls)
 totals=[sums(15,22)-sums(17,24),sums(19,26),sums(20,27)]
 if any(abs(a-b)>0.01 for a,b in zip(totals,v[3:6])):bad.append([k,v[3:6],totals])
 c['mixed_rate']+=len(rates)>1
 c['discount_invoices']+=any(x['v'][24 if x['mt'] else 17] not in [None,'',0] for x in ls)
 c['missing_rate']+=any(x['v'][25 if x['mt'] else 18] in [None,''] for x in ls)
 c['missing_qty']+=any(x['v'][20 if x['mt'] else 13] in [None,''] for x in ls)
print(c,'BAD',len(bad));print(json.dumps(bad[:15],ensure_ascii=False))
p={'roots':[{'key':list(k),'values':list(v),'lines':lines[k]} for k,v in roots.items()],'mismatches':bad,'audit':dict(c)}
(r/'scripts/data/untreated/import-data.json').write_text(json.dumps(p,ensure_ascii=False,default=str),encoding='utf8')
