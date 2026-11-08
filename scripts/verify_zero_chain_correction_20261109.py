"""Verify the correction against the preserved, previously delivered draft."""
from pathlib import Path
from collections import Counter
from decimal import Decimal
import csv, json

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/import-2025-20261109'
BEFORE = OUT / 'archive/before-zero-chain-fix-20261109'

def read(path):
    return json.loads(path.read_text(encoding='utf-8'))

def decisions(base):
    with (base / 'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig', newline='') as f:
        return {r['Chứng từ AMIS']: r for r in csv.DictReader(f)}

def main():
    old, new = read(BEFORE / 'authoring.json'), read(OUT / 'authoring.json')
    prepared = read(OUT / 'prepared.json')
    od, nd = decisions(BEFORE), decisions(OUT)
    assert od.keys() == nd.keys()
    changed = [c for c in nd if od[c]['Phân loại'] != nd[c]['Phân loại']]
    assert len(changed) == 419
    assert all(od[c]['Phân loại'] == 'Chờ xác minh' and nd[c]['Phân loại'] == 'Tạo import mới' for c in changed)
    raw = {r['doc']: r for r in prepared['decisions']}
    roots = Counter()
    for code in changed:
        d = raw[code]
        assert d['date'] < '2025-07-01'
        assert d['roots']
        for key in d['roots']:
            ch = prepared['chains'][key]
            assert not ch['issues'] and len(ch['balance']) == 3
            assert all(v is not None and Decimal(str(v)) == 0 for v in ch['balance'])
            roots[key] += 1
    assert new['zero_audit'] == old['zero_audit'], 'No further old-invoice mutation is needed'
    for kind in ['adjustment', 'replacement']:
        a = next(o for o in old['outputs'] if o['kind'] == kind)
        b = next(o for o in new['outputs'] if o['kind'] == kind)
        assert a == b
        assert (BEFORE / a['file']).read_bytes() == (OUT / 'deliverables' / b['file']).read_bytes()
    a = next(o for o in old['outputs'] if o['kind'] == 'new')
    b = next(o for o in new['outputs'] if o['kind'] == 'new')
    assert a['headers'] == b['headers']
    assert b['rows'][:len(a['rows'])] == a['rows'], 'Preserve every previous item and invoice sequence'
    added = b['rows'][len(a['rows']):]
    assert len(added) == 469 and {r[18] for r in added} == set(changed)
    assert sorted({r[0] for r in added}) == list(range(29794, 30213))
    old_return = next(o for o in old['outputs'] if o['kind'] == 'return_adjustment')
    new_return = next(o for o in new['outputs'] if o['kind'] == 'return_adjustment')
    targets = {'BTL00356': 'BHHT2025014212', 'BTL00378': 'BHHT2025014037', 'BTL00407': 'BHHT2025014069'}
    row_changes = []
    seqs = {r[18]: r[0] for r in b['rows']}
    assert len(old_return['rows']) == len(new_return['rows']) == 410
    for i, (before, after) in enumerate(zip(old_return['rows'], new_return['rows']), 10):
        if before != after:
            assert before[:28] == after[:28]
            assert after[24] in targets and after[25] == targets[after[24]]
            assert after[28:] == ['IMPORT_HD_MOI_2025.xlsx', seqs[after[25]], 'Chờ số HĐ bán mới sau phát hành']
            assert all(after[j] is None for j in [9, 10, 11, 12])
            row_changes.append({'row': i, 'return': after[24], 'sale': after[25], 'invoice_sequence': after[29]})
    assert len(row_changes) == 5
    blocked = [r for r in prepared['decisions'] if r['action'] == 'review' and r['roots'] and all(k in roots for k in r['roots'])]
    assert len(blocked) == 1
    sums = {name: sum(Decimal(str(r[idx] or 0)) for r in added) for name, idx in [('goods', 16), ('vat', 11)]}
    sums['pay'] = sums['goods'] + sums['vat']
    result = {'passed': True, 'added_invoices': len(changed), 'added_item_rows': len(added), 'added_totals': {k: str(v) for k, v in sums.items()}, 'roots': dict(roots), 'all_existing_item_rows_unchanged': len(a['rows']), 'old_zeroing_files_binary_unchanged': True, 'return_row_changes': row_changes, 'still_blocked': [{'sale': r['doc'], 'decision': nd[r['doc']]} for r in blocked], 'new_summary': new['summary']}
    (OUT / 'zero-chain-correction-verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: result[k] for k in ['passed', 'added_invoices', 'added_item_rows', 'added_totals', 'roots', 'return_row_changes', 'still_blocked']}, ensure_ascii=False))

if __name__ == '__main__':
    main()
