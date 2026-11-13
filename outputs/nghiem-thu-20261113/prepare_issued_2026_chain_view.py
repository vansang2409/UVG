"""Tạo bảng chuỗi F0...Fn theo chứng từ, phân biệt xuất lại qua BHH với quan hệ pháp lý."""
from pathlib import Path
from collections import defaultdict,Counter
from functools import lru_cache
import json,sys,hashlib,unicodedata
sys.stdout.reconfigure(encoding='utf-8');O=Path(__file__).parent
def read(n):return json.loads((O/n).read_text(encoding='utf-8'))
def save(n,x):(O/n).write_text(json.dumps(x,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
def lines(s):return [x for x in str(s or '').split('\n') if x]
def norm(s):return ''.join(c for c in unicodedata.normalize('NFD',str(s or '').lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
original=O/'TONG_HOP_HD_2026_CHO_DON_2024_2025_RUT_GON.xlsx'
assert hashlib.sha256(original.read_bytes()).hexdigest()=='403bc88ad1d180de06caa56a6077d2cff6f1c2b4b12906838e7c4589583bc88d'
d=read('issued-2026-report-data.json');selection=read('issued-2026-selection-audit.json');ledger=read('ledger.json')
edges=read('edges.json');old=read('reverse-chain-audit.json');edges+=old['recovered_edges']
resolved={i['key']:i['resolved'] for c in old['components'] for i in c['invoices']}
del old
replaced={e['old'] for e in edges if e['kind']=='replacement'}
parent={};children=defaultdict(set);kind={}
for e in edges:
 a,b=e['old'],e['new']
 assert b not in parent or parent[b]==a
 parent[b]=a;children[a].add(b);kind[a,b]=e['kind']
@lru_cache(None)
def root(k):return root(parent[k]) if k in parent else k
@lru_cache(None)
def depth(k):return depth(parent[k])+1 if k in parent else 0
@lru_cache(None)
def members(k):
 assert root(k)==k
 out=set();todo=[k]
 while todo:
  x=todo.pop()
  if x in out:continue
  out.add(x);todo.extend(children[x]-out)
 assert all(x in ledger for x in out)
 return tuple(sorted(out,key=lambda x:(depth(x),ledger[x]['day'],x)))

target=set(selection['invoice_keys']);assert len(target)==82207
records=[];covered=set();depth_counts=Counter();branch_rows=0
def add_row(year,doc,r,extra,current,source_note='',source_new=None):
 global branch_rows
 ns=members(r);legal=defaultdict(list)
 for k in ns:legal[depth(k)].append(k)
 phases={0:[r],1:[],2:[],3:list(extra)}
 for k in ns:
  if k==r:continue
  phases[1 if ledger[k]['day'][:4]<'2026' else 2].append(k)
 for keys in phases.values():keys.sort(key=lambda k:(ledger[k]['day'],k))
 notes=[]
 if not doc:notes.append('Chưa xác định chứng từ AMIS')
 if source_note:notes.append(source_note)
 if any(len(ks)>1 for ks in legal.values()):branch_rows+=1
 if current:
  newkeys=source_new or []
  if len(newkeys)>1 or current not in newkeys:notes.append('Đang map AMIS: '+current)
 else:
  if doc:notes.append('Chưa có HĐ đang map AMIS')
 if 'không hợp lệ' in ledger[r]['cqt_status'].lower():notes.append('F0: CQT không hợp lệ')
 covered.update(k for k in ns if ledger[k]['day'].startswith('2026-'))
 covered.update(extra)
 for dep,ks in legal.items():depth_counts[dep]+=len(ks)
 records.append({'year':str(year),'doc':doc,'root':r,'nodes':list(ns),'legal_stages':{str(k):v for k,v in legal.items()},'phases':{str(k):v for k,v in phases.items()},'reissued':extra,'current':current,'notes':'; '.join(dict.fromkeys(notes))})

for ref in d['bhh_rows']:
 year,doc=ref[0],ref[1];newkeys=lines(ref[12]);roots={root(k) for k in lines(ref[11])+newkeys}
 assert roots
 processed=sorted(r for r in roots if len(members(r))>1 or ledger[r]['day'][:4]<'2026')
 standalone=sorted(roots-set(processed))
 if processed:
  assert all(ledger[r]['day'].startswith('2026-') and len(members(r))==1 for r in standalone)
  for r in processed:
   note=ref[10]
   if len(processed)>1:note='; '.join(filter(None,[note,f'Cùng chứng từ có {len(processed)} chuỗi gốc']))
   add_row(year,doc,r,standalone,ref[4],note,newkeys)
 else:
  for r in standalone:add_row(year,doc,r,[],ref[4],ref[10],newkeys)

# Đưa các chuỗi chưa có BHH vào cùng bảng, thay cho tab theo dõi riêng đã bỏ.
for comp in selection['components']:
 if comp['docs']:continue
 assert len(comp['root'])==1
 r=comp['root'][0];year=ledger[r]['day'][:4]
 add_row(year,'',r,[],'')

assert covered==target,{'missing':sorted(target-covered)[:10],'extra':sorted(covered-target)[:10]}
seen=set()
for r in records:
 ident=(r['year'],r['doc'],r['root']);assert ident not in seen;seen.add(ident)
max_depth=max(int(k) for r in records for k in r['legal_stages'])
phase_titles=['F0 - HĐ gốc','F1 - ĐC/TT trước 2026','F2 - Xử lý năm 2026','F3 - Xuất mới / xuất lại 2026']
headers=['Nhóm chuỗi','Năm đơn','BHH / chứng từ']
for i in range(4):headers += [f'HĐ F{i}',f'Ngày F{i}',f'Thanh toán F{i} (đ)']
headers+=['Giá trị sau cùng (đ)','Ghi chú cần rà']
records.sort(key=lambda r:(r['year'],not bool(r['doc']),r['doc'],ledger[r['root']]['day'],r['root']))
rows=[]
for i,r in enumerate(records,1):
 # Chỉ cộng HĐ còn hiệu lực; HĐ thay thế được tính thay cho HĐ cũ.
 active=[];excluded=[];fallback=[];amounts={};mismatches=[]
 for k in dict.fromkeys(r['nodes']+r['reissued']):
  x=ledger[k]
  if k in replaced or 'bi thay the' in norm(x['status']) or 'huy' in norm(x['status']) or 'khong hop le' in norm(x['cqt_status']):
   excluded.append(k);continue
  active.append(k);g,v,pay=resolved.get(k,x['money'])
  if pay is None and g is not None and v is not None:
   pay=g+v;fallback.append(k)
  elif x['money'][2] is None and pay is not None:fallback.append(k)
  if x['money'][2] is not None and g is not None and v is not None and abs(g+v-x['money'][2])>1:mismatches.append(k)
  amounts[k]=pay
 final=None if any(v is None for v in amounts.values()) else sum(amounts.values())
 # Các nguồn trống hiện tại đều được xác nhận giá trị 0; không đổi ô nguồn thành 0.
 assert all(amounts[k]==0 for k in active if ledger[k]['money'][2] is None)
 refs=[];start=len(rows)
 for phase in range(4):
  for j,k in enumerate(r['phases'][str(phase)]):
   if k in active:refs.append(['F','I','L','O'][phase]+str(start+j+7))
 formula='=SUM('+','.join(refs)+')' if refs else '=0'
 extra=[]
 if fallback:extra.append(f'{len(fallback)} HĐ trống Thanh toán: tính 0 theo Tiền hàng + VAT nguồn')
 if any('khong hop le' in norm(ledger[k]['cqt_status']) for k in excluded):extra.append('Không cộng HĐ CQT không hợp lệ')
 if mismatches:extra.append('Nguồn lệch Thanh toán với Tiền hàng + VAT: giữ số Thanh toán')
 if final is None:extra.append('Chưa đủ số tiền để tính giá trị cuối')
 r.update(final_value=final,final_active=active,final_excluded=excluded,final_zero_fallback=fallback,final_formula=formula)
 r['notes']='; '.join(filter(None,[r['notes'],*extra]))
 for j in range(max(len(v) for v in r['phases'].values())):
  cells=[]
  for phase in range(4):
   keys=r['phases'][str(phase)]
   if j>=len(keys):cells += ['',None,None];continue
   k=keys[j];suffix=''
   if phase in [1,2]:suffix=' (ĐC)' if kind[parent[k],k]=='adjustment' else ' (TT)'
   if 'bi thay the' in norm(ledger[k]['status']):suffix+=' (đã bị TT)'
   if 'huy' in norm(ledger[k]['status']):suffix+=' (đã hủy)'
   cells += [k+suffix,ledger[k]['day'],ledger[k]['money'][2]]
  row=[i,r['year'],r['doc'],*cells,final if j==0 else None,r['notes'] if j==0 else '']
  assert all(not isinstance(v,str) or len(v)<=32767 for v in row)
  rows.append(row)
summary={'rows':len(rows),'groups':len(records),'max_depth':max_depth,'phase_count':4,'presentation':'business_phases_with_dates_and_payments','columns':len(headers),'source_2026_invoices_covered':len(covered),'documents':len(d['bhh_rows']),'unmapped_chain_groups':sum(not r['doc'] for r in records),'rows_with_parallel_branches':branch_rows,'rows_with_reissues':sum(bool(r['reissued']) for r in records),'legal_stage_memberships':dict(depth_counts),'final_value_counts':dict(Counter('unknown' if r['final_value'] is None else 'zero' if r['final_value']==0 else 'positive' if r['final_value']>0 else 'negative' for r in records))}
save('issued-2026-chain-view.json',{'summary':summary,'phase_titles':phase_titles,'headers':headers,'rows':rows})
save('issued-2026-chain-view-audit.json',{'summary':summary,'records':records,'invoice_keys':sorted(covered)})
print(json.dumps(summary,ensure_ascii=False,indent=2),flush=True)
