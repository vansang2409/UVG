"""Independent openpyxl readback of invoice coverage, duplicate pairs and output."""
import csv
import hashlib
import json
import sys
from collections import Counter, defaultdict
from decimal import Decimal
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / 'scripts/data/new-sales-issued-20261101'


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    summary = json.loads((WORK / 'summary.json').read_text(encoding='utf-8'))
    resolved = json.loads((WORK / 'resolved-coverage.json').read_text(encoding='utf-8'))
    records = {}
    for entry in summary['source_manifest']:
        source = ROOT / entry['file']
        assert hashlib.sha256(source.read_bytes()).hexdigest() == entry['sha256']
        wb = load_workbook(BytesIO(source.read_bytes()), read_only=True, data_only=True)
        sh = wb.active
        header = next(sh.iter_rows(min_row=5, max_row=5, values_only=True))
        ix = {v: i for i, v in enumerate(header) if v}
        local = []
        for rn, r in enumerate(sh.iter_rows(min_row=6, values_only=True), 6):
            if not isinstance(r[0], (int, float)):
                continue
            key = f"{r[ix['Ký hiệu']]}/{str(r[ix['Số hóa đơn']]).zfill(8)}"
            assert key not in records
            a = {'key': key, 'order': str(r[ix['Mã đơn hàng']]).strip(), 'date': r[ix['Ngày hóa đơn']],
                 'money': tuple(None if r[ix[c]] in (None, '') else Decimal(str(r[ix[c]])) for c in ['Doanh số bán chưa thuế', 'Tiền thuế GTGT', 'Tổng tiền']),
                 'cqt_status': r[ix['Trạng thái gửi CQT']], 'cqt_code': r[ix['Mã của CQT']], 'row': rn}
            records[key] = a
            local.append(a)
        control = next(c for c in summary['export']['footer_controls'] if c['file'] == entry['file'])['money']
        assert [sum((a['money'][i] for a in local if a['money'][i] is not None), Decimal(0)) for i in range(3)] == control
        wb.close()
    mapping = defaultdict(set)
    for path in WORK.glob('google-import-map-*.json'):
        for r in json.loads(path.read_text(encoding='utf-8'))['values']:
            mapping[r[1]].add(str(r[0]).strip())
    plan = defaultdict(lambda: [Decimal(0), Decimal(0)])
    wb = load_workbook(ROOT / summary['plan']['file'], read_only=True, data_only=True)
    for r in wb.active.iter_rows(min_row=10, values_only=True):
        if r[0] is None:
            continue
        plan[r[18]][0] += Decimal(str(r[16]))
        if r[11] is not None:
            plan[r[18]][1] += Decimal(str(r[11]))
    wb.close()
    expected = Counter()
    for doc, amounts in plan.items():
        assert len(mapping[doc]) == 1
        expected[(next(iter(mapping[doc])), (*amounts, sum(amounts)))] += 1
    assert max(expected.values()) == 1
    found = Counter((a['order'], a['money']) for a in records.values())
    zero_blank = [(o, money) for o, money in expected if money == (0, 0, 0) and not found[(o, money)] and found[(o, (Decimal(0), Decimal(0), None))] == 1]
    missing = sum(max(0, n - found[k]) for k, n in expected.items())
    assert missing == len(zero_blank) == 2
    extras = {k: found[k] - n for k, n in expected.items() if found[k] > n}
    assert len(extras) == sum(extras.values()) == 2000
    assert len(records) == 37012 and len(plan) == 34739
    assert len(resolved['duplicate_pairs']) == 2000
    for p in resolved['duplicate_pairs']:
        a, b = (records[r['key']] for r in p['invoices'])
        assert a['order'] == b['order'] == p['order']
        assert a['money'] == b['money'] == tuple(p['expected'])
        assert a['date'] == b['date'] == '01/11/2026'
        assert int(b['key'].split('/')[1]) - int(a['key'].split('/')[1]) == 2000
        assert a['cqt_status'] == b['cqt_status'] == 'HĐ hợp lệ' and a['cqt_code'] and b['cqt_code']
    with (ROOT / 'outputs/kiem-tra-xuat-moi-20261101/CAC_CAP_TRUNG_MA_DON_VA_TIEN.csv').open(encoding='utf-8-sig', newline='') as f:
        data = list(csv.reader(f))
    assert len(data) == 2001
    for r, p in zip(data[1:], resolved['duplicate_pairs']):
        assert r[0] == p['doc'] and r[2] == p['order'] and r[3] == p['invoices'][0]['key'] and r[5] == p['invoices'][1]['key']
    output = ROOT / 'outputs/kiem-tra-xuat-moi-20261101/KIEM_TRA_XUAT_MOI_2024.xlsx'
    if output.exists():
        wb = load_workbook(output, read_only=True, data_only=True)
        sh = wb['Kiem tra trung']
        assert sh['C15'].data_type == 's' and sh['C15'].number_format == '@'
        for r, p in zip(sh.iter_rows(min_row=15, max_row=2014, max_col=13, values_only=True), resolved['duplicate_pairs']):
            assert r[1] == p['doc'] and r[2] == p['order'] and r[3] == p['invoices'][0]['key'] and r[4] == p['invoices'][1]['key']
            assert list(r[6:9]) == p['expected']
        assert sh['D4'].value == 34739 and sh['D5'].value == 2000
        assert [sh[f'D{r}'].value for r in [6, 7, 8]] == [2883691447, 230695306, 3114386753]
        wb.close()
    result = {'passed': True, 'source_reader': 'openpyxl independently of XML audit parser', 'export_invoices': len(records), 'plan_documents': len(plan), 'exact_money_coverage': len(plan) - missing,
              'unknown_zero_payment': len(zero_blank), 'duplicate_groups_and_extra_invoices': 2000, 'source_footer_controls_match': True, 'csv_rows': len(data) - 1, 'workbook_readback': output.exists()}
    (WORK / 'verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()
