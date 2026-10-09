"""Đối chiếu nguồn meInvoice mới/cũ, chỉ đọc workbook và giữ phạm vi gốc đã chốt."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal
from zipfile import ZipFile
import io,json,sys,hashlib
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount
import openpyxl
sys.stdout.reconfigure(encoding='utf-8')
BASE=ROOT/'data/Du lieu Meinvoice'; NEW=BASE/'Xuat_lai_2026-10-09'
old=defaultdict(lambda: {'header':[],'detail':[],'observations':[]})
new=defaultdict(lambda: {'header':[],'detail':[],'observations':[]})
edges={'old':set(),'new':set()};inventory=[];hashes={}
def put(dst,k,family,day,ref,money=None,line=None,status=None):
 if not k:return
 dst[k]['observations'].append({'family':family,'date':day.isoformat() if day else None,'source':ref})
 if money is not None:dst[k]['header'].append({'money':money,'date':day.isoformat() if day else None,'source':ref,'status':status})
 if line is not None:dst[k]['detail'].append(line)
def norm(v):return ' '.join(str(v or '').split())
def canon(vals):return tuple(None if v is None else str(v) for v in vals)
def scan(path,data,version,kind,mtt=False):
 dst=old if version=='old' else new
 refpath=str(path);count=0;period=None
 for rn,r in rows(io.BytesIO(data)):
  if rn==3:period=r.get('A')
  if amount(r.get('A')) is None:continue
  ref=refpath+' | dòng '+str(rn)
  if kind=='total':
   k=invoice_key(r.get('B'),r.get('C'));day=invoice_day(r.get('D'))
   put(dst,k,'Bảng tổng',day,ref,money=[amount(r.get(c)) for c in ['P','Q','R']],status=r.get('U'))
  elif kind=='detail':
   if version=='new':
    k=invoice_key(r.get('B'),r.get('C'));day=invoice_day(r.get('D'));name,unit,qty,price,goods,tax,pay=[r.get(c) for c in ['L','M','N','O','P','R','S']]
   else:
    k=invoice_key(r.get('F' if mtt else 'D'),r.get('B'));day=invoice_day(r.get('D' if mtt else 'C'))
    cols=['S','T','U','V','W','AA','AB'] if mtt else ['L','M','N','O','P','T','U'];name,unit,qty,price,goods,tax,pay=[r.get(c) for c in cols]
    disc=amount(r.get('Y' if mtt else 'R'))
    if amount(goods) is not None and disc is not None:goods=amount(goods)-disc
   line={'date':day.isoformat() if day else None,'name':norm(name),'unit':norm(unit),'qty':amount(qty),'price':amount(price),'goods':amount(goods),'tax':amount(tax),'payment':amount(pay),'source':ref}
   put(dst,k,'Chi tiết',day,ref,line=line)
  else:
   a=invoice_key(r.get('B'),r.get('D'));b=invoice_key(r.get('M'),r.get('O'));daya=invoice_day(r.get('C'));dayb=invoice_day(r.get('N'))
   if a and b:edges[version].add((a,b,kind))
   put(dst,a,'Liên kết '+kind,daya,ref);put(dst,b,'Liên kết '+kind,dayb,ref)
  count+=1
 inventory.append({'version':version,'source':refpath,'kind':kind,'period':period,'rows':count})
for p in sorted((BASE/'Bang_ke_hoa_don_da_su_dung_2024_2025').glob('*.xlsx')):scan(p.relative_to(ROOT),p.read_bytes(),'old','total')
for p in sorted(BASE.glob('*.xls')):
 if 'chi_tiet_HD' in p.name:scan(p.relative_to(ROOT),p.read_bytes(),'old','detail','_MTT' in p.name)
 elif 'dieu_chinh_xuat' in p.name:scan(p.relative_to(ROOT),p.read_bytes(),'old','adjust')
 elif 'thay_the_xuat' in p.name:scan(p.relative_to(ROOT),p.read_bytes(),'old','replace')
print('Đã đọc nguồn cũ',len(old),flush=True)
for p in sorted(NEW.iterdir()):
 if p.suffix.lower() not in ['.xls','.xlsx','.zip']:continue
 data=p.read_bytes();hashes[str(p.relative_to(ROOT))]=hashlib.sha256(data).hexdigest()
 sources=[(str(p.relative_to(ROOT)),data)]
 if p.suffix.lower()=='.zip':
  z=ZipFile(io.BytesIO(data));sources=[(str(p.relative_to(ROOT))+' / '+n,z.read(n)) for n in z.namelist() if n.lower().endswith(('.xls','.xlsx'))]
 for label,data in sources:
  kind='detail' if 'chi_tiet' in label else ('adjust' if 'dieu_chinh_xuat' in label else ('replace' if 'thay_the_xuat' in label else 'total'))
  scan(label,data,'new',kind)
print('Đã đọc nguồn mới',len(new),flush=True)
# Chọn gốc theo ngày hóa đơn; đi theo liên kết, không mở rộng gốc ngoài kỳ.
roots=set()
for dataset in [old,new]:
 for k,v in dataset.items():
  for h in v['header']:
   if h['date'] and '2024-01-01'<=h['date']<='2025-06-30' and h['status'] not in ['Hoá đơn điều chỉnh','Hóa đơn điều chỉnh','Hoá đơn thay thế','Hóa đơn thay thế']:roots.add(k)
all_edges=edges['old']|edges['new'];scoped=set(roots)
while True:
 before=len(scoped)
 for a,b,t in all_edges:
  if a in scoped:scoped.add(b)
 if len(scoped)==before:break
changes=[];date_conflicts=[];missing=[];new_only=[]
fields=['name','unit','qty','price','goods','tax','payment']
for k in sorted(scoped):
 a=old.get(k);b=new.get(k)
 if b and len({o['date'] for o in b['observations'] if o['date']})>1:date_conflicts.append({'key':k,'observations':b['observations']})
 if a and b:
  diff={}
  ah=a['header'][0] if a['header'] else None;bh=b['header'][0] if b['header'] else None
  if ah and bh:
   for f in ['money','date','status']:
    if ah[f]!=bh[f]:diff['header_'+f]={'old':ah[f],'new':bh[f]}
  if a['detail'] and b['detail']:
   ca=Counter(canon([l[f] for f in fields]) for l in a['detail']);cb=Counter(canon([l[f] for f in fields]) for l in b['detail'])
   if ca!=cb:diff['detail_values']={'old_lines':len(a['detail']),'new_lines':len(b['detail']),'old_only':list((ca-cb).items()),'new_only':list((cb-ca).items())}
   da=sorted({l['date'] for l in a['detail'] if l['date']});db=sorted({l['date'] for l in b['detail'] if l['date']})
   if da!=db:diff['detail_dates']={'old':da,'new':db}
  if diff:changes.append({'key':k,'changes':diff})
 if b and b['header'] and not b['detail']:missing.append({'key':k,'date':b['header'][0]['date'],'old_detail_available':bool(a and a['detail']),'status':b['header'][0]['status']})
 if b and not a:new_only.append(k)
edge_change={v:[{'original':a,'related':b,'kind':t} for a,b,t in sorted(es) if a in scoped] for v,es in [('added',edges['new']-edges['old']),('absent_in_new',edges['old']-edges['new'])]}
# Kiểm ảnh hưởng các hóa đơn được tham chiếu trong sheet import hiện có.
import_results={}
for name in ['HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx','HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx']:
 w=openpyxl.load_workbook(ROOT/'outputs'/name,read_only=True,data_only=True); keys={(r[9],r[10]) for r in w['Import dieu chinh'].iter_rows(min_row=10,max_col=11,values_only=True) if r[0] is not None};changed_roots=keys&{tuple(r['key']) for r in changes}; affected=[]
 for c in changes:
  ck=tuple(c['key']);rs={a for a,b,t in all_edges if b==ck and a in keys}
  if ck in keys:rs.add(ck)
  if rs:affected.append({'invoice':ck,'roots':sorted(rs),'changes':list(c['changes'])})
 import_results[name]={'roots':len(keys),'direct_changed_roots':sorted(changed_roots),'affected':affected,'missing_new_detail_roots':sorted(keys&{tuple(r['key']) for r in missing})};w.close()
target=('1C25TUV','00012885')
result={'date':'2026-10-09','scope':'F0 2024-01-01..2025-06-30 và các hóa đơn liên quan','roots':len(roots),'scoped_invoices':len(scoped),'inventory':inventory,'source_hashes':hashes,'changes':changes,'date_conflicts':date_conflicts,'missing_new_detail':missing,'new_only':new_only,'edge_changes':edge_change,'import_impact':import_results,'case_12885':{'old':old.get(target),'new':new.get(target)},'limits':['Chưa có bảng liên kết điều chỉnh/thay thế năm 2026 nên chưa xác nhận không có liên kết mới năm 2026.','Không sửa nguồn hoặc workbook, không chạy builder.','Không dùng Excel thay XML để kết luận nội dung điều chỉnh.']}
# Chỉ phân loại khác biệt 0/trống để đối chiếu, không sửa nguồn.
blank_zero=0;other_details=[]
for change in changes:
 detail_change=change['changes'].get('detail_values')
 if not detail_change:continue
 def diagnostic_counter(entries):
  c=Counter()
  for vals,n in entries:c[tuple(vals[:2]+[('0' if z is None else z) for z in vals[2:]])]+=n
  return c
 if diagnostic_counter(detail_change['old_only'])==diagnostic_counter(detail_change['new_only']):blank_zero+=1
 else:other_details.append(change)
result['detail_difference_classification']={'blank_vs_zero_only_invoices':blank_zero,'other_detail_changes':other_details,'note':'Chỉ phân loại khác biệt xuất báo cáo; không chuyển ô trống nguồn mới thành 0.'}
output=ROOT/'scripts/data/check-review/source-refresh-review-20261009.json';output.write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
print(json.dumps({'roots':len(roots),'scoped':len(scoped),'changes':len(changes),'change_types':dict(Counter(f for r in changes for f in r['changes'])),'date_conflicts':[{'key':r['key'],'dates':sorted({o['date'] for o in r['observations'] if o['date']})} for r in date_conflicts],'missing_new_detail':len(missing),'missing_covered_old':sum(x['old_detail_available'] for x in missing),'new_only':len(new_only),'edge_changes':{k:len(v) for k,v in edge_change.items()},'imports':{n:{'affected':len(v['affected']),'missing_detail_roots':len(v['missing_new_detail_roots'])} for n,v in import_results.items()},'case_12885_dates':{'old':sorted({o['date'] for o in old[target]['observations'] if o['date']}),'new':sorted({o['date'] for o in new[target]['observations'] if o['date']})}},ensure_ascii=False,indent=2))