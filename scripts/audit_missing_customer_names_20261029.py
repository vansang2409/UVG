"""Đối chiếu tên nguồn meInvoice cho các hóa đơn import đang thiếu tên; chỉ đọc."""
import sys, json, re, io, hashlib
import posixpath
import argparse
from lxml import etree
from pathlib import Path
from zipfile import ZipFile
from collections import defaultdict, Counter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from prepare_invoice_reports import rows, invoice_key, amount

def blank(v):
    return not str(v or '').strip()

parser = argparse.ArgumentParser()
parser.add_argument('--book', default=str(ROOT / 'outputs/HOA_DON_DA_DIEU_CHINH_IMPORT_2_SHEET.xlsx'))
parser.add_argument('--expected', type=int, default=478)
parser.add_argument('--out', default=str(ROOT / 'outputs/missing-customer-names-20261029'))
args = parser.parse_args()
groups = {}
book = Path(args.book)
with ZipFile(book) as z:
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    wb = etree.fromstring(z.read('xl/workbook.xml'))
    sh = next(s for s in wb.findall('s:sheets/s:sheet', ns) if s.get('name') == 'Import dieu chinh')
    rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    rel = next(r for r in etree.fromstring(z.read('xl/_rels/workbook.xml.rels')) if r.get('Id') == rid)
    target = rel.get('Target')
    target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
    temp = io.BytesIO()
    with ZipFile(temp, 'w') as tmp:
        tmp.writestr('xl/worksheets/sheet1.xml', z.read(target))
        if 'xl/sharedStrings.xml' in z.namelist():
            tmp.writestr('xl/sharedStrings.xml', z.read('xl/sharedStrings.xml'))
temp.seek(0)
for rn, r in rows(temp):
    if rn < 10 or blank(r.get('A')) or not blank(r.get('C')):
        continue
    seq = str(r['A'])
    root = invoice_key(r.get('J'), r.get('K'))
    old = str(r.get('O') or '')
    matches = re.findall(r'HĐ\s+(1C\d{2}\w+)/(\d+)', old)
    source_key = invoice_key(*matches[0]) if matches and ('đảo khoản' in old.lower()) else root
    item = groups.setdefault(seq, dict(seq=seq, root=root, source_key=source_key, rows=[], buyer_blank=True))
    assert item['root'] == root
    item['rows'].append(rn)
    item['buyer_blank'] &= blank(r.get('F'))
roots = {tuple(x['root']) for x in groups.values()}
if args.expected >= 0:
    assert len(groups) == args.expected, len(groups)
wanted = roots | {tuple(x['source_key']) for x in groups.values()}
print(json.dumps(dict(stage='scope', import_invoices=len(groups), import_rows=sum(len(x['rows']) for x in groups.values()), roots=len(roots), source_keys=len(wanted))), flush=True)

observations = defaultdict(list)
details = defaultdict(list)
hashes = {}
base = ROOT / 'data/Du lieu Meinvoice/Xuat_lai_2026-10-09'
for p in sorted(base.iterdir()):
    is_total = p.name.startswith('Bang_ke_hoa_don_da_su_dung_')
    is_detail = p.name.startswith('Bang_ke_chi_tiet_hoa_don_da_su_dung_')
    if not (is_total or is_detail) or p.suffix not in ('.zip', '.xlsx', '.xls'):
        continue
    # Bộ này là xuất ngoài kỳ; giữ cùng phạm vi nguồn đã dùng tạo import.
    if '1791532943641' in p.name:
        continue
    raw = p.read_bytes()
    hashes[str(p.relative_to(ROOT))] = hashlib.sha256(raw).hexdigest()
    if p.suffix == '.zip':
        with ZipFile(io.BytesIO(raw)) as z:
            books = [(p.name + '/' + n, z.read(n)) for n in z.namelist() if n.endswith(('.xlsx', '.xls'))]
    else:
        books = [(p.name, raw)]
    for label, data in books:
        for rn, r in rows(io.BytesIO(data)):
            if amount(r.get('A')) is None:
                continue
            key = invoice_key(r.get('B'), r.get('C'))
            if key not in wanted:
                continue
            if is_total:
                observations[key].append(dict(file=label, row=rn, customer=r.get('H'), buyer=r.get('L'), date=r.get('D')))
            else:
                details[key].append(dict(file=label, row=rn, buyer=r.get('H')))
    print(json.dumps(dict(stage='read', file=p.name, found_headers=len(observations), found_details=len(details))), flush=True)

def classify(key):
    hits = observations.get(tuple(key), [])
    if not hits:
        return 'not_found'
    signatures = {(str(h['customer'] or '').strip(), str(h['buyer'] or '').strip()) for h in hits}
    if len(signatures) != 1:
        return 'conflicting_sources'
    customer, buyer = next(iter(signatures))
    return 'both_blank' if not customer and not buyer else 'customer_only' if customer and not buyer else 'buyer_only' if buyer and not customer else 'both_present'

result = []
for item in groups.values():
    rt = tuple(item['root']); sk = tuple(item['source_key'])
    result.append(dict(**item, root_class=classify(rt), root_evidence=observations.get(rt, []), source_class=classify(sk), source_evidence=observations.get(sk, []), root_detail_rows=len(details.get(rt, [])), root_detail_nonblank=[x for x in details.get(rt, []) if not blank(x['buyer'])]))
summary = dict(import_invoices=len(result), import_rows=sum(len(x['rows']) for x in result), f0_roots=len(roots), import_by_f0_class=dict(Counter(x['root_class'] for x in result)), roots_by_class=dict(Counter(classify(k) for k in roots)), import_by_source_class=dict(Counter(x['source_class'] for x in result)), roots_without_detail=sum(not details.get(k) for k in roots), roots_with_detail_name=sum(any(not blank(x['buyer']) for x in details.get(k, [])) for k in roots), all_import_buyers_blank=all(x['buyer_blank'] for x in result))
report = dict(summary=summary, source_hashes=hashes, records=result)
out = Path(args.out)
out.mkdir(parents=True, exist_ok=True)
(out / 'audit.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(dict(summary=summary, exceptions=[x for x in result if x['root_class'] != 'both_blank'][:5], output=str(out / 'audit.json')), ensure_ascii=False), flush=True)
