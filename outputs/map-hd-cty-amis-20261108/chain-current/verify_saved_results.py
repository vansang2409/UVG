"""Verify saved chain results and the readback of the updated Google Sheet."""
from pathlib import Path
from collections import Counter, defaultdict
import hashlib
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
OUTPUTS = ROOT / 'outputs'

def read(path):
    return json.loads(path.read_text(encoding='utf-8'))

def at(row, col):
    return row[col] if col < len(row) else None

saved = read(HERE/'comparison-by-final-chain.json')
values = read(HERE/'final-sheet-values.json')['values']
before = read(HERE/'hd-cty-before-chain.json')['values']
cells = read(HERE/'final-sheet-cells.json')['sheets'][0]['data'][0]['rowData']
record_by_row = {r['row']: r for r in saved['records']}
assert len(record_by_row) == 408 and len(cells) == 409
for rn in range(413):
    for col in range(20):
        assert at(values[rn], col) == at(before[rn], col), ('A:T changed', rn+1, col+1)

counts = Counter()
records = []
money_bad = vat_bad = any_bad = 0
formula_count = 0
vat_only = []
doc_roots = defaultdict(set)
for r in saved['records']:
    doc_roots[r['doc']].add(r['root'])
for rn in range(5, 413):
    row = values[rn-1]
    src = record_by_row[rn]
    assert row[16] == src['doc']
    balance = row[23:26]
    assert balance == src['balance'] and not src['chain']['issues']
    assert at(row,18) == src['amis_vat'] and at(row,19) == src['amis_money']
    assert at(row,20) in (src['deltaVat'], '' if src['deltaVat'] is None else src['deltaVat'])
    assert at(row,21) in (src['deltaMoney'], '' if src['deltaMoney'] is None else src['deltaMoney'])
    assert at(row,26) in (src['deltaRevenue'], '' if src['deltaRevenue'] is None else src['deltaRevenue'])
    if not src['found']:
        status = 'Không tìm thấy CT AMIS 08/11'
    elif len(doc_roots[src['doc']]) > 1:
        status = 'Cần kiểm tra liên kết CT–chuỗi'
    elif src['deltaRevenue'] == src['deltaVat'] == src['deltaMoney'] == 0:
        status = 'Khớp số dư cuối chuỗi'
    elif max(abs(src[k]) for k in ('deltaRevenue','deltaVat','deltaMoney')) <= 2:
        status = 'Lệch cuối chuỗi 1–2 đ'
    else:
        status = 'Lệch số dư cuối chuỗi'
    assert row[22] == status, (rn, row[22], status)
    counts[status] += 1
    if src['found']:
        assert src['deltaVat'] == row[18]-balance[1]
        assert src['deltaMoney'] == row[19]-balance[2]
        assert src['deltaRevenue'] == row[19]-row[18]-balance[0]
        money_bad += src['deltaMoney'] != 0
        vat_bad += src['deltaVat'] != 0
        any_bad += any(src[k] != 0 for k in ('deltaRevenue','deltaVat','deltaMoney'))
        if src['deltaMoney'] == 0 and src['deltaVat'] != 0:
            vat_only.append(rn)
    for cell in cells[rn-4]['values']:
        assert 'errorValue' not in cell.get('effectiveValue', {}), (rn, cell)
        formula_count += 'formulaValue' in cell.get('userEnteredValue', {})
    for ci in (2,3,8):
        assert 'formulaValue' in cells[rn-4]['values'][ci]['userEnteredValue'], (rn, ci)
    records.append({
        'sheet_row': rn, 'invoice': src['invoice'], 'doc': src['doc'],
        'amis_vat': src['amis_vat'], 'amis_payment': src['amis_money'],
        'chain_root': src['root'], 'chain_keys': src['chain']['keys'],
        'active_invoices': src['chain']['active'], 'excluded_invoices': src['chain']['excluded'],
        'chain_revenue': balance[0], 'chain_vat': balance[1], 'chain_payment': balance[2],
        'difference_revenue': src['deltaRevenue'], 'difference_vat': src['deltaVat'],
        'difference_payment': src['deltaMoney'], 'result': status,
        'ktoan_row': src['ktRow'], 'source_display': row[27],
    })

summary = {
    'rows': len(records), 'unique_chains': len({r['chain_root'] for r in records}),
    'payment_difference_rows': money_bad, 'vat_difference_rows': vat_bad,
    'any_difference_rows': any_bad, 'payment_matched_but_vat_differs_rows': vat_only,
    'results': dict(counts), 'formula_count': formula_count,
}
expected = read(HERE/'final-verification.json')
assert (money_bad, vat_bad, any_bad, len(vat_only)) == (36,47,49,13)
assert summary['unique_chains'] == expected['uniqueChains'] == 391
assert dict(counts) == expected['counts'] and expected['issues'] == []
detail_vat = read(HERE.parent/'vat-only-verified.json')
assert [r['row'] for r in detail_vat] == vat_only
for r in detail_vat:
    assert r['deltaBeforeTax'] == -r['deltaVat']

comparison = OUTPUTS/'doi-chieu-amis-0811-20261108'
manifest = read(comparison/'manifest.json')
for entry in manifest:
    path = ROOT/'data/0811'/Path(entry['path']).name
    assert path.stat().st_size == entry['bytes']
    assert hashlib.sha256(path.read_bytes()).hexdigest() == entry['sha256']
source_summary = read(comparison/'comparison-summary.json')
assert source_summary['common_docs'] == 35636
assert len(source_summary['not_in_new_export']) == 10
assert sum(m['change_money'] for m in source_summary['monthly'].values()) == -17221701
assert sum(m['change_vat'] for m in source_summary['monthly'].values()) == -1275682

result = {'spreadsheet_id': '1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw',
          'sheet': 'HD CTY 2025', 'range': 'S4:AB412',
          'invoice_sources': 'meInvoice đến 29/10/2026; bổ sung phát hành 01/11/2026; Ktoan đọc 08/11/2026',
          'amis_source_date': '2026-11-08', 'summary': summary, 'records': records}
(HERE/'final-chain-results.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'verified':True, 'source_sha256_verified':len(manifest), **summary},ensure_ascii=False))
