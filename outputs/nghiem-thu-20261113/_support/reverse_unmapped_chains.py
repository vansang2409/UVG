"""Đối chiếu ngược chuỗi meInvoice còn giá trị với AMIS 2024–2025."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal
from functools import lru_cache
import json,re,hashlib,unicodedata,sys
sys.stdout.reconfigure(encoding='utf-8')
R=Path('D:/UVG');O=Path(__file__).parent
sys.path.insert(0,str(R/'scripts'))
from prepare_invoice_reports import rows as source_rows,invoice_key
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def norm(v):return ''.join(c for c in unicodedata.normalize('NFD',str(v or '').lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
def key(v):return '/'.join(v) if isinstance(v,(list,tuple)) else v
def dec(v):return None if v is None or v=='' else Decimal(str(v))
def native(v):
    if isinstance(v,Decimal):return int(v) if v==int(v) else float(v)
    if isinstance(v,set):return sorted(v)
    raise TypeError(type(v).__name__)
def dump(n,v):(O/n).write_text(json.dumps(v,ensure_ascii=False,separators=(',',':'),default=native),encoding='utf-8')
ledger=read(O/'ledger.json');detail=read(O/'detail.json');edges=read(O/'edges.json');mapping=read(O/'mapping.json');amis=read(O/'amis.json')
manifest=read(O/'manifest.json')
for m in manifest:
    if 'member' not in m and '/meinvoice/' in m['file'].replace('\\','/'):
        assert hashlib.sha256((R/m['file']).read_bytes()).hexdigest()==m['sha256']
header_hashes={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in (R/'data/1311/amis').glob('*.xlsx')}
parents=defaultdict(set);children=defaultdict(set);replacement=defaultdict(set);typed={};evidence=defaultdict(list);linkmoney=defaultdict(lambda:defaultdict(set));recovered=[];edge_conflicts=[]
def add(a,b,kind,src,historical=False):
    if historical:
        if a not in ledger or b not in ledger:return
        if kind=='replacement' and 'bi thay the' not in norm(ledger[a]['status']) and (a,b) not in typed:
            edge_conflicts.append({'old':a,'new':b,'kind':kind,'issue':'Quan hệ lịch sử khác trạng thái bảng kê mới','source':src});return
        if kind=='replacement' and replacement[a] and b not in replacement[a]:
            edge_conflicts.append({'old':a,'new':b,'kind':kind,'existing':sorted(replacement[a]),'source':src});return
        if parents[b] and a not in parents[b]:
            edge_conflicts.append({'old':a,'new':b,'kind':kind,'existing_parents':sorted(parents[b]),'source':src});return
    if (a,b) not in typed:
        typed[a,b]=kind;parents[b].add(a);children[a].add(b)
        if kind=='replacement':replacement[a].add(b)
        if historical:recovered.append({'old':a,'new':b,'kind':kind,'source':src})
    elif typed[a,b]!=kind:edge_conflicts.append({'old':a,'new':b,'kind':kind,'source':src,'issue':'Loại quan hệ khác nhau'})
    evidence[a,b].append(src)
for e in edges:
    add(e['old'],e['new'],e['kind'],e['source'])
    for k,vs in [(e['old'],e['old_money']),(e['new'],e['new_money'])]:
        for i,v in enumerate(vs):
            if v is not None:linkmoney[k][i].add(dec(v))
historical_sources=[]
p=R/'scripts/data/post-issue-20261029/audit.json';old=read(p);historical_sources.append({'file':str(p.relative_to(R)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
for e in old['edges']:add(key(e['old']),key(e['new']),e['kind'],{'file':str(p.relative_to(R)),'evidence':e['sources']},True)
del old
p=R/'outputs/import-2025-20261109/source-index.json';old=read(p);historical_sources.append({'file':str(p.relative_to(R)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
for a,bs in old['children'].items():
    for b in bs:add(a,b,'replacement' if b in old['replacement'].get(a,[]) else 'adjustment',{'file':str(p.relative_to(R))},True)
del old
raw_history=[R/'data/Du lieu Meinvoice/Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_2024_2025.xls']
raw_history+=sorted((R/'data/Du lieu Meinvoice/Xuat_lai_2026-10-09').glob('*thay_the*.xls'))
raw_history+=sorted((R/'data/Du lieu Meinvoice/Xuat_lai_2026-10-22').glob('*thay_the*.xls'))
for p in raw_history:
    if not p.exists():continue
    historical_sources.append({'file':str(p.relative_to(R)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
    for rn,r in source_rows(p):
        if rn<7 or not str(r.get('A') or '').isdigit():continue
        a,b=invoice_key(r.get('B'),r.get('D')),invoice_key(r.get('M'),r.get('O'))
        if a and b:add(key(a),key(b),'replacement',{'file':str(p.relative_to(R)),'row':rn},True)
p=R/'scripts/data/revenue-chain-audit-20261028/prepared.json'
if p.exists():
    old=read(p);historical_sources.append({'file':str(p.relative_to(R)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
    for e in old['edges']:add(key(e['parent']),key(e['child']),e['kind'],{'file':str(p.relative_to(R)),'evidence':e.get('evidence')},True)
    del old
p=O/'reverse-live-chain-31054.json'
if p.exists():
    live=read(p);historical_sources.append({'file':str(p.relative_to(R)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
    for a,b,kind in live['edges']:add(a,b,kind,{'file':str(p.relative_to(R)),'range':live['range']},True)
print('GRAPH',len(ledger),'invoices',len(typed),'edges',len(recovered),'recovered',flush=True)
# Chỉ lấy số tiền từ bộ mới; số thanh toán trống không mặc định 0.
fallbacks=[];amount_issues=defaultdict(list)
for k,r in ledger.items():
    r['resolved']=[dec(v) for v in r['money']];r['fallback']=[]
    for i,v in enumerate(r['resolved']):
        vs=linkmoney[k][i]
        if v is None and len(vs)==1:r['resolved'][i]=next(iter(vs));r['fallback'].append('Bảng liên kết 13/11')
        elif v is not None and vs and (len(vs)>1 or v not in vs):amount_issues[k].append('Khác tiền giữa bảng tổng và liên kết')
        if r['resolved'][i] is None and k in detail and detail[k]['known'][i]>0:
            # Bỏ qua ô dòng hàng trống; chỉ cộng các ô có số thật trong bảng chi tiết.
            r['resolved'][i]=dec(detail[k]['sums'][i]);r['fallback'].append('Bảng chi tiết 13/11')
    if r['fallback']:fallbacks.append(k)
policy=R/'scripts/data/post-issue-20261029/confirmed-blank-payment-zero.json'
if policy.exists():
    pol=read(policy);historical_sources.append({'file':str(policy.relative_to(R)),'sha256':hashlib.sha256(policy.read_bytes()).hexdigest()})
    for kk in pol['invoice_keys']:
        k=key(kk)
        if k in ledger and ledger[k]['resolved']==[Decimal(0),Decimal(0),None]:ledger[k]['resolved'][2]=Decimal(0);ledger[k]['fallback'].append('Xác nhận thanh toán 0 đã có')
if (O/'reverse-live-chain-31054.json').exists() and ledger['1C26MUV/00002639']['resolved']==[Decimal(0),Decimal(0),None]:
    ledger['1C26MUV/00002639']['resolved'][2]=Decimal(0);ledger['1C26MUV/00002639']['fallback'].append('Ktoan live AD30482:AF30482 xác nhận chuỗi 0')
known=defaultdict(set);orders=defaultdict(set);candidate_hints=defaultdict(set)
for m in mapping:
    for k in set(m['keys']+([m['symbol']+'/'+m['invoice']] if m['invoice'] else [])):known[k].add(m['id'])
    for v in m['order']:
        if v and not v.startswith('#'):orders[v].add(m['id'])
    for k in m['candidates']+m['history']:candidate_hints[k].add(m['id'])
    previous=m.get('previous_mapping') or {}
    for k in previous.get('candidates',[])+([previous['key']] if previous.get('key') else []):candidate_hints[k].add(m['id'])
# Giữ nguyên toàn bộ mã chứng từ, kể cả .1 và mã không theo mẫu năm.
doc_lookup=defaultdict(set)
for dk,d in amis.items():doc_lookup[d['doc']].add(dk)
canonical=read(O/'final-links-audit.json')
for doc,ks in canonical['final_links'].items():
    for k in ks:known[k].update(doc_lookup[doc])
graph=defaultdict(set)
for a,b in typed:graph[a].add(b);graph[b].add(a)
allkeys=set(ledger)|set(graph);seen=set();results=[];out_counts=Counter();scope_counts=Counter();root_counts=Counter()
for seed in sorted(allkeys):
    if seed in seen:continue
    todo=[seed];nodes=set()
    while todo:
        k=todo.pop()
        if k in nodes:continue
        nodes.add(k);seen.add(k);todo.extend(graph[k]-nodes)
    roots=sorted(k for k in nodes if not parents[k]);linked=set().union(*(known[k] for k in nodes))
    exact_orders={ledger[k]['order'] for k in nodes if k in ledger and ledger[k]['order'] in orders}
    source_doc_tokens=set()
    for k in nodes:
        if k not in ledger:continue
        order=ledger[k]['order']
        if order in doc_lookup:source_doc_tokens.add(order)
        source_doc_tokens.update(c for c in re.findall(r'(?<![A-Z0-9])(?:BHH[A-Z]*|BHDV)\d+(?:\.\d+)?(?![A-Z0-9.])',order) if c in doc_lookup)
    root_years={ledger[k]['day'][:4] for k in roots if k in ledger and ledger[k]['day']}
    declared_root_years={ledger[k]['day'][:4] if k in ledger and ledger[k]['day'] else '20'+k[2:4] if re.match(r'1C\d{2}',k) else '?' for k in roots}
    if root_years & {'2024','2025'}:scope='F0 năm '+','.join(sorted(root_years))
    elif linked:scope='HĐ phát hành sau cho AMIS2024/2025'
    elif exact_orders or source_doc_tokens:scope='Ứng viên liên quan AMIS2024/2025 theo mã đơn/CT'
    elif declared_root_years & {'2024','2025'}:scope='F0 ký hiệu 2024/2025, thiếu ngày gốc'
    else:out_counts['Ngoài phạm vi']+=1;continue
    scope_counts[scope]+=1
    issues=[];active=[];excluded=[]
    if len(roots)!=1:issues.append('Không có duy nhất F0: '+str(len(roots)))
    for k in sorted(nodes):
        if k not in ledger:issues.append('Thiếu HĐ trong bảng tổng: '+k);continue
        r=ledger[k];st=norm(r['status'])
        if replacement[k] or 'bi thay the' in st or 'huy' in st:
            excluded.append(k)
            if 'bi thay the' in st and not replacement[k]:issues.append('Thiếu liên kết thay thế: '+k)
            if replacement[k] and len(replacement[k])>1:issues.append('Có nhiều HĐ thay thế: '+k)
            continue
        active.append(k);issues+=amount_issues[k]
        if len(parents[k])>1:issues.append('HĐ có nhiều cha: '+k)
        if st in ['hoa don dieu chinh','hoa don thay the'] and not parents[k]:issues.append('Thiếu HĐ cha: '+k)
        if 'da bi dieu chinh' in st and not any(typed[k,c]=='adjustment' for c in children[k]):issues.append('Thiếu liên kết điều chỉnh: '+k)
        if 'tu choi' in norm(r['cqt_status']) or 'khong hop le' in norm(r['cqt_status']):issues.append('Trạng thái CQT cần kiểm: '+k)
    for k in nodes:
        if replacement[k] and any(typed[k,c]=='adjustment' and c in active for c in children[k]):issues.append('Có nhánh điều chỉnh gắn HĐ đã bị thay thế: '+k)
    balance=[];known_sums=[];missing=[]
    for i in range(3):
        vals=[ledger[k]['resolved'][i] for k in active];known_sums.append(sum(v for v in vals if v is not None))
        bad=[k for k in active if ledger[k]['resolved'][i] is None]
        missing.append(bad);balance.append(None if bad else sum(vals))
    nonzero=any(v is not None and v!=0 for v in balance)
    zero=all(v==0 for v in balance)
    goods_vat_zero=balance[:2]==[Decimal(0),Decimal(0)] and balance[2] is None
    if linked:status='Đã map AMIS'
    elif zero and not issues:status='Đã về 0'
    elif nonzero and not issues:status='Chưa map - còn giá trị'
    elif goods_vat_zero and not issues:status='Tiền hàng/VAT 0, thanh toán trống'
    else:status='Chưa map - cần xác minh chuỗi/tiền'
    hints=set().union(*(candidate_hints[k] for k in nodes))
    order_docs=set().union(*(orders[v] for v in exact_orders)) if exact_orders else set()
    token_docs=set().union(*(doc_lookup[v] for v in source_doc_tokens)) if source_doc_tokens else set()
    if len(roots)==1:root_counts[ledger.get(roots[0],{}).get('day','?')[:4]]+=1
    rec={'root':roots[0] if len(roots)==1 else '; '.join(roots),'root_day':ledger.get(roots[0],{}).get('day') if len(roots)==1 else None,'scope':scope,'status':status,'keys':sorted(nodes),'active':active,'excluded':excluded,'balance':balance,'known_sums':known_sums,'missing':missing,'issues':sorted(set(issues)),'mapped_docs':sorted(linked),'order_candidates':sorted(order_docs|token_docs),'historical_candidates':sorted(hints),'orders':sorted({ledger[k]['order'] for k in nodes if k in ledger and ledger[k]['order']}),'buyers':sorted({ledger[k]['buyer'] for k in active if ledger[k]['buyer']}),'invoices':[dict(ledger[k],resolved=ledger[k]['resolved']) for k in sorted(nodes) if k in ledger]}
    results.append(rec)
counts=Counter(r['status'] for r in results);unmapped=[r for r in results if r['status']=='Chưa map - còn giá trị'];review=[r for r in results if r['status'].startswith('Chưa map - cần')]
summary={'as_of':'2026-11-13','ledger':len(ledger),'edges_current':len(edges),'edges_after_recovery':len(typed),'recovered_edges':len(recovered),'edge_conflicts':len(edge_conflicts),'in_scope_components':len(results),'scope_counts':dict(scope_counts),'statuses':dict(counts),'unmapped_nonzero':len(unmapped),'unmapped_nonzero_by_root_year':dict(Counter((r['root_day'] or '?')[:4] for r in unmapped)),'unmapped_nonzero_balances':[sum(r['balance'][i] for r in unmapped if r['balance'][i] is not None) for i in range(3)],'review_unknown':len(review),'source_amis_sha256':header_hashes,'historical_sources':historical_sources}
dump('reverse-chain-audit.json',{'summary':summary,'components':results,'recovered_edges':recovered,'edge_conflicts':edge_conflicts});dump('reverse-chain-unmapped.json',{'summary':summary,'unmapped_nonzero':unmapped,'review':review,'zero_payment_blank':[r for r in results if r['status']=='Tiền hàng/VAT 0, thanh toán trống']})
print(json.dumps(summary,ensure_ascii=False,default=native),flush=True)
print('UNMAPPED_SAMPLE',json.dumps([{k:r[k] for k in ['root','root_day','balance','buyers','orders','order_candidates','historical_candidates','issues']} for r in unmapped[:6]],ensure_ascii=False,default=native),flush=True)
print('REVIEW_SAMPLE',json.dumps([{k:r[k] for k in ['root','balance','issues']} for r in review[:12]],ensure_ascii=False,default=native),flush=True)
