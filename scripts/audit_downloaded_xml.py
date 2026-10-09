"""Đọc hai lô XML vừa tải, đối chiếu nhóm237; không sửa workbook."""
from pathlib import Path
from collections import Counter
from decimal import Decimal
from zipfile import ZipFile
import xml.etree.ElementTree as E
import json,hashlib,sys
import openpyxl
ROOT=Path(__file__).resolve().parents[1];sys.stdout.reconfigure(encoding='utf-8')
review=json.loads((ROOT/'scripts/data/check-review/zero-compliance-review.json').read_text(encoding='utf-8'))
wanted={tuple(r['key']):r for r in review['records']}
outputs=json.loads((ROOT/'scripts/data/check-review/outputs-vs-new-data-20261009.json').read_text(encoding='utf-8'))
missing_roots={tuple(x['key']) for v in outputs['outputs'].values() for x in v.get('import',{}).get('old_detail_fallback',[])}
wb=openpyxl.load_workbook(ROOT/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',read_only=True,data_only=True)
current_dates={}
for row in wb['Dieu chinh'].iter_rows(min_row=2,min_col=7,max_col=9,values_only=True):
 if row[0] and row[1] and row[2]:
  current_dates[(str(row[0]),str(row[1]).zfill(8))]=row[2].strftime('%Y-%m-%d') if hasattr(row[2],'strftime') else str(row[2])[:10]
wb.close()
found={};inventory=[];errors=[];duplicates=[]
def tag(e):return e.tag.rsplit('}',1)[-1]
def first(e,name):return next((x for x in e.iter() if tag(x)==name),None)
def value(e,name):
 child=next((x for x in e if tag(x)==name),None) if e is not None else None
 return child.text if child is not None else None
def amt(v):return None if v in [None,''] else Decimal(v)
def key(sig,num,prefix):return (str(prefix or '')+str(sig or ''),str(num or '').zfill(8))
for name in ['Hoa_don_dien_tu_09.10.2026_02.51.zip','Hoa_don_dien_tu_09.10.2026_02.52.zip']:
 p=Path(r'C:\Users\tranq\Downloads')/name;z=ZipFile(p);records=[]
 for filename in z.namelist():
  if not filename.lower().endswith('.xml'):continue
  try:
   data=z.read(filename);root=E.fromstring(data);hd=first(root,'DLHDon');tt=first(hd,'TTChung');pay=first(hd,'TToan');rel=first(tt,'TTHDLQuan');k=key(value(tt,'KHHDon'),value(tt,'SHDon'),value(tt,'KHMSHDon'))
   related=key(value(rel,'KHHDCLQuan'),value(rel,'SHDCLQuan'),value(rel,'KHMSHDCLQuan')) if rel is not None else None
   lines=[]
   for l in hd.iter():
    if tag(l)!='HHDVu':continue
    extra={value(x,'TTruong'):value(x,'DLieu') for x in l.iter() if tag(x)=='TTin'}
    lines.append({'nature':value(l,'TChat'),'name':value(l,'THHDVu'),'qty':amt(value(l,'SLuong')),'price':amt(value(l,'DGia')),'goods':amt(value(l,'ThTien')),'discount':amt(value(l,'STCKhau')),'tax_rate':value(l,'TSuat'),'tax_extra':amt(extra.get('VATAmount'))})
   money=[amt(value(pay,t)) for t in ['TgTCThue','TgTThue','TgTTTBSo']]
   r={'key':k,'date':value(tt,'NLap'),'related':related,'related_date':value(rel,'NLHDCLQuan') if rel is not None else None,'kind':value(rel,'TCHDon') if rel is not None else None,'money':money,'lines':lines,'zip':name,'xml':filename,'sha256':hashlib.sha256(data).hexdigest(),'signature_elements':sum(tag(x)=='Signature' for x in root.iter()),'has_tax_code':first(root,'MCCQT') is not None}
   records.append(r)
   if k in found:duplicates.append(k)
   found[k]=r
   if k in wanted:
    d=ROOT/'data/Bo sung/meInvoice/2026-10-09-237-XML';d.mkdir(parents=True,exist_ok=True);(d/(k[0]+'_'+k[1]+'.xml')).write_bytes(data)
  except Exception as e:errors.append({'zip':name,'xml':filename,'error':str(e)})
 inventory.append({'path':str(p),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'xml':len(records),'pdf':sum(n.lower().endswith('.pdf') for n in z.namelist()),'min_date':min(r['date'] for r in records),'max_date':max(r['date'] for r in records),'matched_237':sum(r['key'] in wanted for r in records),'matched_57_originals':sum(r['key'] in missing_roots for r in records)})
matched=[];counts=Counter()
for k,old in wanted.items():
 if k not in found:continue
 r=found[k];r['header_matches_report']=r['money']==[Decimal(str(v)) for v in old['money']];r['reference_matches_report']=r['related']==tuple(old['root']);r['date_matches_prior_review']=r['date']==old['date'][:10];r['date_matches_current_workbook']=r['date']==current_dates.get(k)
 counts['header_matches_report']+=r['header_matches_report'];counts['reference_matches_report']+=r['reference_matches_report']
 counts['date_matches_prior_review']+=r['date_matches_prior_review'];counts['date_matches_current_workbook']+=r['date_matches_current_workbook']
 monetary=[l for l in r['lines'] if l['nature'] in ['1','2','3']]
 goods=sum((l['goods'] or Decimal(0) for l in monetary),Decimal(0));r['monetary_line_goods_sum']=goods;r['line_goods_vs_header_mismatch']=goods!=r['money'][0]
 counts['line_goods_vs_header_mismatch']+=r['line_goods_vs_header_mismatch'];counts['has_type4_note']+=any(l['nature']=='4' for l in r['lines']);counts['all_line_goods_zero']+=all(l['goods'] in [None,Decimal(0)] for l in r['lines']);counts['negative_unit_price']+=any(l['price'] is not None and l['price']<0 for l in r['lines'])
 matched.append(r)
result={'review_date':'2026-10-09','inventory':inventory,'unique_invoices':len(found),'errors':errors,'duplicates':duplicates,'matched_237':len(matched),'missing_237':[k for k in wanted if k not in found],'matched_57_originals':len(missing_roots&found.keys()),'counts':dict(counts),'records_237':matched,'case_12885':found.get(('1C25TUV','00012885')),'limits':['Đã đọc dữ liệu XML, chưa xác minh mật mã chữ ký hoặc tra cứu CQT trực tiếp.','Tính tổng dòng được dùng để phát hiện bất nhất, không tự kết luận mọi hóa đơn sai pháp lý.','Chưa sửa workbook hoặc import.']}
(ROOT/'scripts/data/check-review/zip-237-xml-review-20261009.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ['records_237','case_12885']},ensure_ascii=False,indent=2,default=str))
print('12885',json.dumps({k:v for k,v in result['case_12885'].items() if k not in ['lines']},ensure_ascii=False,default=str))