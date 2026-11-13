"""Lập dữ liệu HĐ phát hành 2026 cho đơn AMIS 2024/2025, giữ cả chuỗi về 0."""
from pathlib import Path
from collections import defaultdict, Counter, deque
from decimal import Decimal
from functools import lru_cache
import json, sys, hashlib, re, unicodedata

sys.stdout.reconfigure(encoding='utf-8')
R=Path('D:/UVG'); O=Path(__file__).parent
def read(n): return json.loads((O/n).read_text(encoding='utf-8'))
def norm(s): return ''.join(c for c in unicodedata.normalize('NFD',str(s or '').lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
def save(n,d): (O/n).write_text(json.dumps(d,ensure_ascii=False,separators=(',',':'),default=lambda x:sorted(x) if isinstance(x,set) else str(x)),encoding='utf-8')
def native(x): return None if x is None else int(x) if x==int(x) else float(x)
def key(s,n): return str(s).strip()+'/'+str(n).removesuffix('.0').zfill(8) if s and n else None
def split(k): return k.split('/',1) if k else ['','']
def join(v): return '\n'.join(sorted(set(v)))

ledger=read('ledger.json'); mappings=read('mapping.json'); amis=read('amis.json')
audit=read('reverse-chain-audit.json'); edges=read('edges.json')
parents=defaultdict(set); children=defaultdict(set); adj=defaultdict(set); rep=defaultdict(set)
typed={}; evidence={}; linkmoney=defaultdict(lambda:defaultdict(set))
for e in edges:
 a,b=e['old'],e['new']; typed[a,b]=e['kind']; evidence[a,b]=e['source']
 for k,vs in [(a,e['old_money']),(b,e['new_money'])]:
  for j,v in enumerate(vs):
   if v is not None:linkmoney[k][j].add(Decimal(str(v)))
for e in audit['recovered_edges']:
 typed[e['old'],e['new']]=e['kind'];evidence[e['old'],e['new']]=e['source']
for (a,b),kind in typed.items():
 parents[b].add(a);children[a].add(b)
 (rep if kind=='replacement' else adj)[a].add(b)
assert len(typed)==18821

# Tính đồ thị từ toàn bộ nguồn, không bỏ các chuỗi có số dư bằng 0.
ci={}; components=[]; resolved={}
for c in audit['components']:
 for r in c['invoices']:resolved[r['key']]=r['resolved']
for seed in sorted(ledger):
 if seed in ci:continue
 nodes=set();todo=[seed]
 while todo:
  k=todo.pop()
  if k in nodes:continue
  nodes.add(k);todo.extend((parents[k]|children[k])-nodes)
 i=len(components)
 for k in nodes:ci[k]=i
 components.append({'keys':sorted(nodes),'roots':sorted(k for k in nodes if not parents[k])})
print('GRAPH',len(components),len(typed),flush=True)

current=defaultdict(lambda:defaultdict(set)); history=defaultdict(lambda:defaultdict(set))
order_index=defaultdict(set); doc_index=defaultdict(set); map_byid={m['id']:m for m in mappings}
for m in mappings:
 for k in set(m['keys']+([m['symbol']+'/'+m['invoice']] if m['invoice'] else [])):
  if k in ci:current[ci[k]][m['id']].add('HĐ đang map AMIS')
 for k in m['history']:
  if k in ci:history[ci[k]][m['id']].add('Liên kết BHH lịch sử đã lưu')
 for order in m['order']:
  if order and not order.startswith('#'):order_index[order].add(m['id'])
 doc_index[m['doc']].add(m['id'])
order_links=defaultdict(lambda:defaultdict(set)); kt_orders=defaultdict(set)
for k,r in ledger.items():
 candidates=set(order_index.get(r['order'],set()))|set(doc_index.get(r['order'],set()))
 for d in candidates:order_links[ci[k]][d].add('Mã đơn/chứng từ khớp nguyên vẹn bảng kê 13/11')
for p in sorted((R/'outputs/import-2025-20261109/source').glob('kt_*.json')):
 d=json.loads(p.read_text(encoding='utf-8'))
 for row in d['values']:
  if len(row)<11:continue
  k=key(row[9],row[10]);order=str(row[3]) if len(row)>3 and row[3] else ''
  if k not in ci or not order or order.startswith('#'):continue
  kt_orders[k].add(order)
  for doc in order_index.get(order,set()):order_links[ci[k]][doc].add('Mã đơn nguyên vẹn trong bảng kế toán và AMIS')

selected=[]; excluded_2026=0; conflicts=[]
for i,c in enumerate(components):
 ks=[k for k in c['keys'] if k in ledger and ledger[k]['day'].startswith('2026-')]
 if not ks:continue
 root_years={ledger[k]['day'][:4] for k in c['roots'] if k in ledger}
 links=defaultdict(set)
 for source in [current[i],history[i]]:
  for d,whys in source.items():links[d].update(whys)
 if links:
  extras=set(order_links[i])-set(links)
  if extras:conflicts.append({'root':c['roots'],'linked':sorted(links),'order_extra_not_applied':sorted(extras)})
 else:
  for d,whys in order_links[i].items():links[d].update(whys)
 if not links and not root_years&{'2024','2025'}:
  excluded_2026+=len(ks);continue
 c.update({'id':i,'issued':ks,'links':dict(links),'root_years':sorted(root_years)})
 selected.append(c)
print('SELECTED',len(selected),'invoices',sum(len(c['issued']) for c in selected),'unmapped',sum(not c['links'] for c in selected),flush=True)

# Chỉ dùng số thật. Ô thanh toán trống được giữ trống nếu không có nguồn bổ sung.
needed={k for c in selected for k in c['keys'] if k in ledger}
missing_raw=[]
for k in needed:
 if k in resolved:continue
 vals=ledger[k]['money'][:]
 for j,v in enumerate(vals):
  if v is None and len(linkmoney[k][j])==1:vals[j]=native(next(iter(linkmoney[k][j])))
 resolved[k]=vals
 if any(v is None for v in vals):missing_raw.append(k)
if missing_raw:
 detail=read('detail.json')
 for k in missing_raw:
  if k not in detail:continue
  for j,v in enumerate(resolved[k]):
   if v is None and detail[k]['known'][j]>0:resolved[k][j]=detail[k]['sums'][j]
 del detail

@lru_cache(None)
def depth(k):
 if not parents[k]:return 0
 return min(depth(p)+1 for p in parents[k])
@lru_cache(None)
def subtree(k):
 seen=set();todo=[k]
 while todo:
  x=todo.pop()
  if x in seen:continue
  seen.add(x);todo.extend(children[x]-seen)
 return seen
def balance(nodes):
 active=[k for k in nodes if k in ledger and not rep[k] and 'bi thay the' not in norm(ledger[k]['status']) and 'huy' not in norm(ledger[k]['status'])]
 result=[]
 for j in range(3):
  v=[resolved[k][j] for k in active]
  result.append(None if not v or any(x is None for x in v) else native(sum(Decimal(str(x)) for x in v)))
 return result
def zero_status(bal):
 if bal==[0,0,0]:return 'Chuỗi đã về 0'
 if bal[:2]==[0,0] and bal[2] is None:return 'Tiền hàng/VAT 0, thanh toán trống'
 if any(x is None for x in bal):return 'Thiếu số tiền trong nguồn'
 return 'Chuỗi còn giá trị'
def checked_join(xs):
 v=join(xs);assert len(v)<=32767,('Excel cell length',len(v));return v

# Một HĐ 2026 chỉ xuất hiện một lần trong bảng chính, kể cả HĐ gắn nhiều BHH.
headers=['STT','Năm đơn','BHH liên quan','CT trả/giảm liên quan','Mã đơn hàng','Loại phát hành 2026','Ký hiệu HĐ 2026','Số HĐ 2026','Ngày HĐ 2026','Tiền hàng HĐ 2026 (đ)','VAT HĐ 2026 (đ)','Thanh toán HĐ 2026 (đ)','Ký hiệu HĐ gốc F0','Số HĐ gốc F0','Ngày HĐ gốc F0','Ký hiệu HĐ bị xử lý','Số HĐ bị xử lý','Bậc trong chuỗi','Kết quả chuỗi hiện tại','Tiền hàng còn của chuỗi (đ)','VAT còn của chuỗi (đ)','Thanh toán còn của chuỗi (đ)','HĐ hiện đang map AMIS','Căn cứ nối BHH','Trạng thái HĐ 2026','Trạng thái CQT','Tên người mua trên HĐ','Lưu ý','Dòng bảng kê meInvoice','HĐ gốc 2024/2025 cùng BHH']
old_roots_by_doc=defaultdict(set)
issued_by_doc=defaultdict(set)
for i,c in enumerate(components):
 for d in set(current[i])|set(history[i]):
  old_roots_by_doc[d].update(k for k in c['roots'] if k in ledger and ledger[k]['day'][:4] in ['2024','2025'])
for c in selected:
 for d in c['links']:
  old_roots_by_doc[d].update(k for k in c['roots'] if k in ledger and ledger[k]['day'][:4] in ['2024','2025'])
  issued_by_doc[d].update(c['issued'])
out=[];unmapped=[];bymonth=Counter();purposes=Counter();bhh_set=set();related_docs=set();issued_keys=set();source_ids={};source_list=[]
for c in selected:
 linked=c['links'];sales=sorted(d for d in linked if amis[d]['doc'].startswith(('BHH','BHDV')));returns=sorted(set(linked)-set(sales))
 bhh_set.update(sales);related_docs.update(returns)
 years=sorted({str(amis[d]['year']) for d in sales or list(linked)}) or sorted(set(c['root_years'])&{'2024','2025'})
 year=','.join(years);bal=balance(c['keys']);state=zero_status(bal)
 roots=c['roots'];root_symbols=checked_join(split(k)[0] for k in roots);root_numbers=checked_join(split(k)[1] for k in roots)
 dates=sorted({ledger[k]['day'] for k in roots if k in ledger});root_date=dates[0] if len(dates)==1 else None
 current_keys=sorted({(map_byid[d]['symbol']+'/'+map_byid[d]['invoice']) for d in linked if map_byid[d]['invoice']})
 basis=checked_join(why for whys in linked.values() for why in whys) if linked else 'F0 năm 2024/2025, chưa có liên kết chứng từ'
 for k in c['issued']:
  assert k not in issued_keys;issued_keys.add(k);r=ledger[k];pp=sorted(parents[k]);incoming={typed[p,k] for p in pp};money=resolved[k]
  if 'adjustment' in incoming:
   purpose='Điều chỉnh trong chuỗi về 0' if bal[:2]==[0,0] else 'Điều chỉnh khác'
  elif 'replacement' in incoming:
   purpose='Thay thế bằng HĐ 0' if money[:2]==[0,0] else 'Thay thế khác'
  elif pp:purpose='Xử lý khác'
  else:purpose='Xuất mới / xuất lại'
  notes=[]
  if not linked:notes.append('Chưa tìm thấy BHH trong nguồn; không ghép theo ngày/tiền gần giống')
  elif not current[c['id']] and not history[c['id']]:notes.append('Nối theo mã đơn nguyên vẹn; chưa có liên kết HĐ được chốt trực tiếp')
  if len(sales)>1:notes.append(f'HĐ/chuỗi liên quan {len(sales)} BHH; tiền HĐ chỉ tính một lần')
  if len(roots)!=1:notes.append('Chuỗi có nhiều F0, cần kiểm tra')
  if len(pp)>1:notes.append('HĐ có nhiều cha, cần kiểm tra')
  if any(x is None for x in r['money']):notes.append('Giữ nguyên ô thanh toán trống của bảng tổng')
  if r['money'][2] is None and money[2]==0:notes.append('Số dư chuỗi sử dụng bằng chứng đối chiếu 0 đã có')
  if any(c['roots']==x['root'] for x in conflicts):notes.append('Mã đơn chỉ sang BHH khác; giữ mapping chứng từ đã chốt')
  if 'khong hop le' in norm(r['cqt_status']):notes.append('CQT không hợp lệ')
  if k in {'1C25TUV/00011927','1C25TUV/00011929'}:notes.append('Nguồn giữ lại và AMIS có số HĐ khác nhau')
  ps=checked_join(split(p)[0] for p in pp);pn=checked_join(split(p)[1] for p in pp)
  orders={ledger[x]['order'] for x in c['keys'] if x in ledger and ledger[x]['order']}
  orders.update(o for x in c['keys'] for o in kt_orders[x])
  orders.update(o for d in linked for o in map_byid[d]['order'])
  source=r['source'];sid=source.rsplit('!',1)[0]
  if sid not in source_ids:source_ids[sid]=len(source_ids)+1;source_list.append(sid)
  old_roots=checked_join(k for d in linked for k in old_roots_by_doc[d])
  row=[0,year,checked_join(amis[d]['doc'] for d in sales),checked_join(amis[d]['doc'] for d in returns),checked_join(orders),purpose,*split(k),r['day'],*r['money'],root_symbols,root_numbers,root_date,ps,pn,'F'+str(depth(k)),state,*bal,checked_join(current_keys),basis,r['status'],r['cqt_status'],r['buyer'],'. '.join(notes),source,old_roots]
  assert len(row)==len(headers)
  out.append((r['day'],k,row,c['id']))
  purposes[purpose]+=1;bymonth[r['day'][:7]]+=1
 if not linked:
  unmapped.append([0,year,checked_join(roots),root_date,checked_join(c['issued']),len(c['issued']),state,*bal,checked_join(ledger[k]['buyer'] for k in c['keys'] if k in ledger),'HĐ tổng cũ/chuỗi về 0 chưa có mã đơn hoặc liên kết BHH trong nguồn'])
out.sort(key=lambda v:(v[0],v[1]));rows=[]
for j,(_,k,row,_) in enumerate(out,1):row[0]=j;rows.append(row)
for j,row in enumerate(unmapped,1):row[0]=j

# Nguồn BHH chỉ xuất hiện một lần mỗi chứng từ, để tra lại vị trí AMIS không lặp tiền HĐ.
bhh_headers=['Năm AMIS','BHH / chứng từ','Ngày chứng từ','Khách hàng AMIS','HĐ hiện đang map AMIS','Tiền hàng AMIS (đ)','VAT AMIS (đ)','Thanh toán AMIS (đ)','Mã đơn','Vị trí AMIS','Ghi chú mapping','HĐ gốc 2024/2025 liên quan','HĐ phát hành 2026 liên quan']
bhh_rows=[]
for d in sorted(bhh_set|related_docs):
 m=map_byid[d];ad=amis[d]
 bhh_rows.append([m['year'],m['doc'],m['day'],m['customer'],(m['symbol']+'/'+m['invoice']) if m['invoice'] else '',*m['amis_money'],checked_join(m['order']),m['source_amis'],'. '.join(m['issues']),checked_join(old_roots_by_doc[d]),checked_join(issued_by_doc[d])])

summary={'as_of':'2026-11-13','source_invoice_count':len(ledger),'source_2026_count':sum(r['day'].startswith('2026') for r in ledger.values()),'issued_2026':len(rows),'chain_count':len(selected),'bhh_count':len(bhh_set),'other_doc_count':len(related_docs),'unmapped_chains':len(unmapped),'unmapped_2026_invoices':sum(x[5] for x in unmapped),'year_counts':dict(Counter(r[1] for r in rows)),'purpose_counts':dict(purposes),'month_counts':dict(sorted(bymonth.items())),'money_sums':[native(sum(Decimal(str(r[j])) for r in rows if r[j] is not None)) for j in [9,10,11]],'missing_money_counts':[sum(r[j] is None for r in rows) for j in [9,10,11]],'mapping_basis':dict(Counter(r[23] for r in rows)),'ignored_order_conflicts':len(conflicts),'excluded_2026_unrelated':excluded_2026,'source_amis_sha256':audit['summary']['source_amis_sha256']}
data={'summary':summary,'headers':headers,'rows':rows,'bhh_headers':bhh_headers,'bhh_rows':bhh_rows,'unmapped_headers':['STT','Năm F0','HĐ gốc F0','Ngày F0','HĐ phát hành 2026 trong chuỗi','Số HĐ 2026','Kết quả chuỗi','Tiền hàng còn (đ)','VAT còn (đ)','Thanh toán còn (đ)','Tên người mua','Phần thiếu trong nguồn'],'unmapped_rows':unmapped,'invoice_sources':source_list}
for p,expected in summary['source_amis_sha256'].items():assert hashlib.sha256(Path(p).read_bytes()).hexdigest()==expected
save('issued-2026-report-data.json',data)
save('issued-2026-selection-audit.json',{'summary':summary,'invoice_keys':sorted(issued_keys),'components':[{'root':c['roots'],'keys':c['keys'],'issued':c['issued'],'docs':c['links']} for c in selected],'order_conflicts_not_applied':conflicts})
print(json.dumps(summary,ensure_ascii=False,indent=2),flush=True)
