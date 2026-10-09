"""Đối chiếu ba output và hai import với bộ xuất mới, không sửa workbook."""
from pathlib import Path
from collections import defaultdict,Counter
from zipfile import ZipFile
from decimal import Decimal
import io,sys,json,hashlib,openpyxl
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts'))
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount
sys.stdout.reconfigure(encoding='utf-8');BASE=ROOT/'data/Du lieu Meinvoice';NEW=BASE/'Xuat_lai_2026-10-09'
headers={};details=defaultdict(list);edges=set();source_hashes={}
def norm(z):return ' '.join(str(z or '').split())
def sourcebooks(p):
 data=p.read_bytes();source_hashes[str(p.relative_to(ROOT))]=hashlib.sha256(data).hexdigest()
 if p.suffix=='.zip':
  z=ZipFile(io.BytesIO(data));return [(str(p.relative_to(ROOT))+' / '+n,z.read(n)) for n in z.namelist() if n.endswith(('.xlsx','.xls'))]
 return [(str(p.relative_to(ROOT)),data)]
for p in sorted(NEW.iterdir()):
 if p.suffix not in ['.xlsx','.xls','.zip'] or '1791532943641' in p.name:continue
 for label,data in sourcebooks(p):
  typ='detail' if 'chi_tiet' in label else ('link' if 'dieu_chinh_xuat' in label or 'thay_the_xuat' in label else 'total')
  for rn,r in rows(io.BytesIO(data)):
   if amount(r.get('A')) is None:continue
   if typ=='total':
    k=invoice_key(r.get('B'),r.get('C'))
    if k:headers[k]={'date':invoice_day(r.get('D')),'money':[amount(r.get(c)) for c in ['P','Q','R']],'buyer':r.get('L'),'name':r.get('H'),'tax':r.get('J'),'status':r.get('U'),'source':label+' | dòng '+str(rn)}
   elif typ=='detail':
    k=invoice_key(r.get('B'),r.get('C'))
    if k:details[k].append({'name':norm(r.get('L')),'unit':norm(r.get('M')),'qty':amount(r.get('N')),'price':amount(r.get('O')),'goods':amount(r.get('P')),'vat':r.get('Q'),'buyer':r.get('H'),'date':invoice_day(r.get('D')),'source':label+' | dòng '+str(rn)})
   else:
    a=invoice_key(r.get('B'),r.get('D'));b=invoice_key(r.get('M'),r.get('O'))
    if a and b:edges.add((a,b))
print('Đã đọc nguồn mới',len(headers),flush=True)
outputs={};known_dates={('1C25TUV','00012885'):'2025-08-18'}
blank_payment_zero={('1C24TUV','00000671'),('1C24TUV','00000744'),('1C25TUV','00000391'),('1C25TUV','00000761'),('1C25TUV','00000835'),('1C25TUV','00004748')}
for name,sheet,start in [('HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx','Chua dieu chinh',0),('HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx','Dieu chinh',0),('NGOAI_LE_F0_2024_T6_2025.xlsx','Ngoai le',2)]:
 p=ROOT/'outputs'/name;w=openpyxl.load_workbook(p,read_only=True,data_only=True);checked=0;issues=[];date_overrides=[];approved=[];rootkeys=set();linked=[]
 for rn,v in enumerate(w[sheet].iter_rows(min_row=3,values_only=True),3):
  if not v[start]:continue
  root=invoice_key(v[start],v[start+1]);rootkeys.add(root)
  root_header=headers.get(root)
  if root_header and sheet in ['Chua dieu chinh','Ngoai le']:
   buyer_pos,tax_pos=(6,7) if sheet=='Chua dieu chinh' else (0,1)
   if norm(v[buyer_pos])!=norm(root_header['name'] or root_header['buyer']):issues.append({'sheet':sheet,'row':rn,'key':root,'type':'buyer','output':v[buyer_pos],'new':root_header['name'] or root_header['buyer']})
   if norm(v[tax_pos])!=norm(root_header['tax']):issues.append({'sheet':sheet,'row':rn,'key':root,'type':'tax_id'})
  blocks=[start] if sheet=='Chua dieu chinh' else ([0,6,12,18] if sheet=='Dieu chinh' else [2,8,14])
  for i in blocks:
   if not v[i]:continue
   k=invoice_key(v[i],v[i+1]);h=headers.get(k);checked+=1;ref={'sheet':sheet,'row':rn,'key':k}
   if not h:issues.append({**ref,'type':'missing_header'});continue
   actual=[amount(v[i+j]) for j in [3,4,5]];expected=h['money']
   if actual!=expected:
    if k in blank_payment_zero and actual[:2]==expected[:2] and actual[2]==0 and expected[2] is None:approved.append({**ref,'type':'approved_blank_payment_zero'})
    else:issues.append({**ref,'type':'money','output':actual,'source':expected})
   day=invoice_day(v[i+2]);newday=h['date']
   if day!=newday:
    if k in known_dates and day.isoformat()==known_dates[k]:date_overrides.append({**ref,'output_date':day,'new_date':newday,'type':'user_confirmed_date_pending_XML'})
    else:issues.append({**ref,'type':'date','output_date':day,'new_date':newday})
   if i!=start:linked.append((root,k))
 for a,b in linked:
  if (a,b) not in edges and not any((parent,b) in edges for root,parent in linked if root==a):issues.append({'type':'missing_link_in_new','root':a,'related':b})
 # Kiểm nguồn chi tiết hiển thị trong tab phụ mà không sửa căn cứ cũ.
 if 'Chi tiet nguon' in w.sheetnames:
  for rn,v in enumerate(w['Chi tiet nguon'].iter_rows(min_row=2,values_only=True),2):
   if not v[2]:continue
   k=invoice_key(v[2],v[3]);h=headers.get(k)
   if h and [amount(z) for z in v[9:12]]!=h['money']:issues.append({'sheet':'Chi tiet nguon','row':rn,'key':k,'type':'source_sheet_money','output':[amount(z) for z in v[9:12]],'new':h['money']})
 summary={'roots':len(rootkeys),'invoice_blocks_checked':checked,'issues':issues,'approved_assumptions':approved,'date_overrides':date_overrides,'missing_new_detail_roots':[k for k in sorted(rootkeys) if not details.get(k)],'workbook_hash':hashlib.sha256(p.read_bytes()).hexdigest()}
 if 'Import dieu chinh' in w.sheetnames:
  groups=defaultdict(list)
  for rn,v in enumerate(w['Import dieu chinh'].iter_rows(min_row=10,max_col=24,values_only=True),10):
   if v[0] is not None:groups[v[0]].append((rn,v))
  impissues=[];blankdiag=[];missingdi=[];reference_dates=[];buyers=[];reduce_count=reverse_count=0
  for seq,lines in groups.items():
   rn,v=lines[0];k=invoice_key(v[9],v[10]);h=headers.get(k)
   if not h:impissues.append({'seq':seq,'type':'missing_original','key':k});continue
   if invoice_day(v[11])!=h['date']:reference_dates.append({'seq':seq,'key':k,'row':rn,'output':invoice_day(v[11]),'new':h['date']})
   for row,z in lines:
    if norm(z[2])!=norm(z[5]):buyers.append({'seq':seq,'row':row,'type':'customer_buyer_differ'})
    if norm(z[5])!=norm(h['buyer']):buyers.append({'seq':seq,'row':row,'key':k,'type':'buyer_vs_new_data','output':z[5],'new':h['buyer']})
    if norm(z[4])!=norm(h['tax']):buyers.append({'seq':seq,'row':row,'key':k,'type':'tax_id_vs_new_data'})
   totals=[amount(v[j]) for j in [19,21,22]]
   if v[16] is None and len(lines)==1:
    reverse_count+=1
    candidates=[b for a,b in edges if a==k and b in headers and [(-x if x is not None else None) for x in headers[b]['money']]==totals]
    if not candidates:impissues.append({'seq':seq,'row':rn,'key':k,'type':'reversal_total_unmatched','totals':totals})
    if amount(v[18])!=totals[0]:impissues.append({'seq':seq,'row':rn,'key':k,'type':'reversal_line_goods'})
   else:
    reduce_count+=1;expected=[-x if x is not None else None for x in h['money']]
    if k in blank_payment_zero and expected[2] is None:expected[2]=Decimal(0)
    if totals!=expected:impissues.append({'seq':seq,'row':rn,'key':k,'type':'reduce_total','output':totals,'new':expected})
    if not details.get(k):missingdi.append({'seq':seq,'key':k,'type':'use_old_detail'});continue
    def counter(a):return Counter(tuple(a) for a in a)
    actual=Counter((norm(z[14]),norm(z[15]),amount(z[16]),amount(z[17]),amount(z[18])) for row,z in lines)
    expect=Counter((l['name'],l['unit'],-l['qty'] if l['qty'] is not None else None,l['price'],-l['goods'] if l['goods'] is not None else None) for l in details[k])
    if actual!=expect:
     def diag(c):
      out=Counter()
      for vals,n in c.items():out[tuple(vals[:2])+tuple(Decimal(0) if z is None else z for z in vals[2:])]+=n
      return out
     if diag(actual)==diag(expect):blankdiag.append({'seq':seq,'key':k,'type':'new_blanks_vs_output_zero'})
     else:impissues.append({'seq':seq,'key':k,'type':'detail','output_only':list((actual-expect).items())[:3],'new_only':list((expect-actual).items())[:3]})
  summary['import']={'invoices':len(groups),'lines':sum(map(len,groups.values())),'reduce_invoices':reduce_count,'reversal_invoices':reverse_count,'issues':impissues,'new_blanks_vs_output_zero':blankdiag,'old_detail_fallback':missingdi,'original_date_mismatch':reference_dates,'customer_buyer_issues':buyers}
 outputs[name]=summary;w.close();print(name,'issues',len(issues),'import',len(summary.get('import',{}).get('issues',[])),flush=True)
result={'review_date':'2026-10-09','source_hashes':source_hashes,'outputs':outputs,'notes':['Không sửa nguồn hoặc output.','Khác biệt ô trống/0 chỉ được phân loại, không thay đổi giả định nguồn.','Ngày12885 giữ xác nhận người dùng, chờXML.']}
(ROOT/'scripts/data/check-review/outputs-vs-new-data-20261009.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
print(json.dumps({name:{'roots':v['roots'],'checked':v['invoice_blocks_checked'],'issues':v['issues'][:8],'issue_count':len(v['issues']),'date_overrides':v['date_overrides'],'missing_new_detail_roots':len(v['missing_new_detail_roots']),'approved_assumptions':len(v['approved_assumptions']),'import':{k:(len(z) if isinstance(z,list) else z) for k,z in v.get('import',{}).items()}} for name,v in outputs.items()},ensure_ascii=False,indent=2,default=str))