"""Rà dữ liệu 237 điều chỉnh đã về 0; không sửa workbook, không kết luận thay XML."""
import json,sys,runpy,contextlib,io,unicodedata,re
from pathlib import Path
from collections import Counter,defaultdict
from decimal import Decimal
from datetime import datetime
import openpyxl
ROOT=Path(__file__).resolve().parents[3]
HELPER=Path(__file__).resolve().parent
sys.stdout.reconfigure(encoding='utf-8')
# Tái đọc nguồn qua các audit hiện có; thu stdout để không in chi tiết dài.
for filename in ([] if '--cached' in sys.argv else ['audit-zero-detail.py','audit-zero-adjustments.py']):
 with contextlib.redirect_stdout(io.StringIO()):runpy.run_path(str(HELPER/filename),run_name='__main__')
a=json.loads((HELPER/'zero-adjustment-review.json').read_text(encoding='utf8'))
b=json.loads((HELPER/'zero-detail-review.json').read_text(encoding='utf8'))
original={tuple(x['key']):x for x in b['roots']}
w=openpyxl.load_workbook(ROOT/'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx',read_only=True,data_only=True)
meta={}
for n,v in enumerate(w['Dieu chinh'].iter_rows(min_row=3,values_only=True),3):
 if all(z==0 for z in v[24:27]):
  for i in range(6,24,6):
   if v[i]:meta[(v[i],v[i+1])]={'date':v[i+2].isoformat(),'root_date':v[2].isoformat(),'root_money':list(v[3:6]),'row':n}
def num(v):return None if v is None else Decimal(str(v))
def norm(v):return re.sub(r'[^\w]+',' ',unicodedata.normalize('NFC',str(v or '')).casefold()).strip()
def group(lines):
 out=defaultdict(lambda:[Decimal(0)]*3)
 for l in lines:
  for i,f in enumerate(['goods','tax','payment']):
   if l[f] is not None:out[norm(l['name'])][i]+=num(l[f])
 return {k:tuple(v) for k,v in out.items() if any(v)}
counts=Counter();records=[];examples={}
for x in a['results']:
 m=meta[tuple(x['key'])];d=x['detail'];o=original[tuple(x['root'])];row={**x,**m}
 before=m['date']<'2025-06-01';counts['before_2025_06_01' if before else 'from_2025_06_01']+=1
 exact=all(num(v)==-num(z) for v,z in zip(x['money'],m['root_money']))
 counts['header_exact_inverse_original']+=exact
 counts['header_arithmetic_ok']+=num(x['money'][0])+num(x['money'][1])==num(x['money'][2])
 counts['header_has_positive_money']+=any(num(v)>0 for v in x['money'])
 counts['detail_total_mismatch']+=x['discrepancy']
 counts['missing_detail']+=not bool(d)
 row['source_header_exact_inverse_original']=exact
 if x['category']=='Chỉ điều chỉnh khoản tổng':
  counts['lump']+=1
  counts['lump_detail_goods_all_zero']+=all(z['goods'] in [0,None] for z in d)
  counts['lump_mismatch']+=x['discrepancy']
  if x['discrepancy']:examples.setdefault('lump_mismatch',row)
 else:
  counts['named']+=1
  counts['named_mismatch']+=x['discrepancy']
  expected={k:tuple(-v for v in vals) for k,vals in group(o['detail']).items()}
  matched=group(d)==expected;counts['named_amounts_by_name_match_original_inverse']+=matched
  row['named_amounts_by_name_match_original_inverse']=matched
  qty_neg=all(z['qty'] is not None and num(z['qty'])<0 for z in d if z['goods'] not in [None,0])
  price_neg=any(z['price'] is not None and num(z['price'])<0 for z in d)
  counts['named_all_nonzero_goods_lines_negative_qty']+=qty_neg
  counts['named_has_negative_price']+=price_neg
  if not matched:examples.setdefault('named_not_matching_original',row)
  if price_neg:examples.setdefault('named_negative_price',row)
 row['assessment']='Chưa đủ căn cứ kết luận quy định: cần XML/bản hiển thị, tính chất dòng, tham chiếu đầy đủ và hồ sơ nghiệp vụ.'
 records.append(row)
assert len(records)==237 and len(meta)==237
result={'review_date':'2026-10-09','scope_roots':237,'adjustment_invoices':237,'old_invoices_including_roots':474,'counts':dict(counts),'original_detail_categories':b['counts'],'examples':examples,'records':records,'evidence_limits':['Không tìm thấy XML/PDF trong data khi rà','Không dùng mô tả dòng làm bằng chứng thiếu tham chiếu trên cả hóa đơn','Không kết luận sai pháp lý chỉ vì chi tiết khác bảng tổng','Mẫu ghi tổng và mẫu ghi chi tiết cần đối chiếu đúng nghiệp vụ thực tế'],'source_hashes':a['hashes']}
(HELPER/'zero-compliance-review.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({k:v for k,v in result.items() if k not in ['records','examples','source_hashes']},ensure_ascii=False,indent=2))
print(json.dumps({'examples':{k:{f:x[f] for f in ['root','key','row','date','money','detail_sums','detail']} for k,x in examples.items()}},ensure_ascii=False))