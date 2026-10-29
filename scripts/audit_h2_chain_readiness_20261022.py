"""Ghép nguồn chỉ đọc để kiểm tra khả năng lập chuỗi F0 nửa cuối 2025."""
import sys, json, io, hashlib, unicodedata
from pathlib import Path
from zipfile import ZipFile
from collections import defaultdict, Counter
from datetime import date
from prepare_invoice_reports import rows, invoice_key, invoice_day, amount

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / 'data/Du lieu Meinvoice'
OUT = ROOT / 'scripts/data/h2-chain-readiness-20261022'
headers, edges, sources, problems = {}, {}, [], []
observations = defaultdict(list)

def record(k, day, money, source, row, status=None, name=None):
    observations[k].append(dict(date=str(day), money=money, source=source, row=row))
    return dict(date=str(day), money=money, source=source, row=row, status=status, name=name)

def read_header(stream, label):
    count = 0; period = None
    for rn, r in rows(stream):
        if rn == 3: period = r.get('A')
        if rn < 6 or amount(r.get('A')) is None: continue
        k = invoice_key(r.get('B'), r.get('C'))
        if not k: continue
        v = record(k, invoice_day(r.get('D')), [r.get(c) for c in ('N','Q','R')], label, rn, r.get('U'), r.get('H') or r.get('L'))
        v.update(buyer=r.get('L') or r.get('H'), tax_id=r.get('J'), order=r.get('F'), cqt=r.get('E'), cqt_status=r.get('S'))
        if k in headers and headers[k]['money'] != v['money']: problems.append(dict(type='duplicate_header_conflict', key=k, old=headers[k], new=v))
        headers[k] = v; count += 1
    sources.append(dict(source=label, period=period, rows=count))

fresh = BASE / 'Xuat_lai_2026-10-09'
for p in sorted(fresh.glob('Bang_ke_hoa_don_da_su_dung*')):
    label = str(p.relative_to(ROOT))
    if p.suffix == '.xlsx': read_header(p, label)
    elif p.suffix == '.zip':
        with ZipFile(p) as z:
            for name in z.namelist():
                if name.endswith('.xlsx'): read_header(io.BytesIO(z.read(name)), label+'!'+name)

relfiles = [
    fresh/'Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_1791533363253.xls',
    fresh/'Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_1791533378419.xls',
    fresh/'Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_1791533430588.xls',
    BASE/'Xuat_lai_2026-10-22/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_1792636231350.xls',
    BASE/'Xuat_lai_2026-10-22/Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_1792637478021.xls',
]
for p in relfiles:
    label = str(p.relative_to(ROOT)); count=0; period=None
    kind = 'replacement' if 'thay_the' in p.name else 'adjustment'
    for rn,r in rows(p):
        if rn==3: period=r.get('A')
        if rn<7 or amount(r.get('A')) is None: continue
        parent=invoice_key(r.get('B'),r.get('D')); child=invoice_key(r.get('M'),r.get('O'))
        if not parent or not child: continue
        for k,dc,mc in [(parent,'C',('H','K','L')),(child,'N',('S','V','W'))]:
            record(k,invoice_day(r.get(dc)),[r.get(c) for c in mc],label,rn)
        edge=dict(parent=parent,child=child,kind=kind,source=label,row=rn)
        if child in edges and edges[child]['parent']!=parent: problems.append(dict(type='multiple_parents',old=edges[child],new=edge))
        edges[child]=edge;count+=1
    sources.append(dict(source=label,period=period,rows=count,sha256=hashlib.sha256(p.read_bytes()).hexdigest()))

# Bổ sung liên kết thiếu trong báo cáo bằng ảnh người dùng gửi, giữ rõ nguồn chứng cứ.
evidence_path = OUT / 'ui-evidence.json'
if evidence_path.exists():
    evidence = json.loads(evidence_path.read_text(encoding='utf-8'))
    for item in evidence['confirmed_edges']:
        parent, child = tuple(item['parent']), tuple(item['child'])
        if child in edges:
            assert tuple(edges[child]['parent']) == parent
        else:
            edges[child] = dict(parent=parent,child=child,kind='replacement',source=str(evidence_path.relative_to(ROOT)),row=None)

def norm(s): return ''.join(c for c in unicodedata.normalize('NFD',s or '') if not unicodedata.combining(c)).lower().replace('đ','d')
def root(k):
    seen=set()
    while k in edges:
        if k in seen: raise ValueError('Cycle '+str(k))
        seen.add(k);k=tuple(edges[k]['parent'])
    return k

roots={k for k,v in headers.items() if '2025-07-01'<=v['date']<='2025-12-31' and k not in edges}
chains=defaultdict(set)
for k in set(headers)|set(observations):
    rt=root(k)
    if rt in roots: chains[rt].add(k)
scope=set().union(*chains.values()) if chains else set()
conflicts=[]; missing=[]
for k in sorted(scope):
    obs=observations[k]
    dates={v['date'] for v in obs}; monies={tuple(amount(x) for x in v['money']) for v in obs}
    if len(dates)>1 or len(monies)>1: conflicts.append(dict(key=k, observations=obs))
    if any(x is None for x in next(iter(monies))): missing.append(dict(key=k, observations=obs))
unlinked=[dict(key=k,**headers[k]) for k in sorted(roots) if any(s in norm(headers[k]['status']) for s in ('hoa don dieu chinh','hoa don thay the'))]
related=[edges[k] for k in sorted(scope) if k in edges]
status_without_child=[]
parents={tuple(e['parent']) for e in edges.values()}
for k in sorted(scope & set(headers)):
    if any(s in norm(headers[k]['status']) for s in ('da bi dieu chinh','da bi thay the')) and k not in parents:
        status_without_child.append(dict(key=k,**headers[k]))
result=dict(sources=sources,header_count=len(headers),header_years=dict(Counter(v['date'][:4] for v in headers.values())),
    h2_headers=sum('2025-07-01'<=v['date']<='2025-12-31' for v in headers.values()),
    root_count=len(roots),root_statuses=dict(Counter(headers[k]['status'] for k in roots)),
    chains_with_related=sum(len(v)>1 for v in chains.values()),scope_invoices=len(scope),
    related_count=len(related),related_by_year_kind=dict(Counter(observations[tuple(e['child'])][0]['date'][:4]+' '+e['kind'] for e in related)),
    missing_header_count=len(scope-set(headers)),missing_header=[dict(key=k, observations=observations[k]) for k in sorted(scope-set(headers))],
    conflicting_values=conflicts,missing_money=missing,unlinked_adjustment_or_replacement=unlinked,
    status_without_child=status_without_child,other_problems=problems,
    max_chain_size=max(map(len,chains.values()),default=0),
    chains=[dict(root=k,invoices=sorted(v)) for k,v in sorted(chains.items())],related_edges=related)
OUT.mkdir(parents=True,exist_ok=True)
(OUT/'audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
summary={k:v for k,v in result.items() if k not in ('chains','related_edges','missing_header','conflicting_values','missing_money','unlinked_adjustment_or_replacement','status_without_child','other_problems')}
summary.update({k:len(result[k]) for k in ('conflicting_values','missing_money','unlinked_adjustment_or_replacement','status_without_child','other_problems')})
(OUT/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
# Dữ liệu dành cho trình tạo workbook, giữ bộ tiền và nguồn dòng của từng hóa đơn.
payload=[]
for c in result['chains']:
    rt=tuple(c['root']); pending=set(map(tuple,c['invoices']))-{rt}; ordered=[rt]
    while pending:
        ready=sorted((k for k in pending if tuple(edges[k]['parent']) in ordered),key=lambda k:(headers[k]['date'],k))
        assert ready, (rt,pending)
        ordered.extend(ready);pending.difference_update(ready)
    invoices=[]
    for k in ordered:
        v=dict(headers[k]);v['key']=k
        v['money']=[float(amount(x)) if amount(x) is not None else None for x in v['money']]
        v['relationship']=edges.get(k)
        invoices.append(v)
    payload.append(dict(root=rt,invoices=invoices))
(OUT/'workbook-input.json').write_text(json.dumps(dict(chains=payload,sources=sources),ensure_ascii=False),encoding='utf-8')
sys.stdout.reconfigure(encoding='utf-8');print(json.dumps({k:v for k,v in summary.items() if k!='sources'},ensure_ascii=False,indent=2))
