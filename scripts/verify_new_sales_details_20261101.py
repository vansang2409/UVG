"""Independent openpyxl verification of the line-level findings and report."""
import hashlib
import json
import sys
from collections import Counter, defaultdict
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / 'scripts/data/new-sales-issued-20261101'


def txt(v):
    return '' if v is None else str(v).strip()


def number(v):
    return None if v in (None, '') else Decimal(str(v))


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    summary = json.loads((WORK / 'detail-summary.json').read_text(encoding='utf-8'))
    previous = json.loads((WORK / 'summary.json').read_text(encoding='utf-8'))
    resolved = json.loads((WORK / 'resolved-coverage.json').read_text(encoding='utf-8'))
    pairs = json.loads((WORK / 'detail-pairs.json').read_text(encoding='utf-8'))
    detail = defaultdict(list)
    for entry in summary['source_manifest']:
        path = ROOT / entry['file']
        assert hashlib.sha256(path.read_bytes()).hexdigest() == entry['sha256']
        wb = load_workbook(BytesIO(path.read_bytes()), read_only=True, data_only=True)
        sh = wb.active
        ix = {v: i for i, v in enumerate(next(sh.iter_rows(min_row=5, max_row=5, values_only=True))) if v}
        for rn, r in enumerate(sh.iter_rows(min_row=6, values_only=True), 6):
            if not isinstance(r[0], (int, float)):
                continue
            get = lambda label: r[ix[label]]
            key = f"{get('Ký hiệu')}/{str(get('Số hóa đơn')).zfill(8)}"
            line = (txt(get('Tên hàng')), txt(get('ĐVT')), number(get('Số lượng')), number(get('Đơn giá')), number(get('Thành tiền')))
            full = (*line, txt(get('Thuế suất (%)')), number(get('Tiền thuế GTGT')), number(get('Tổng tiền TT')), txt(get('Tính chất HHDV')), txt(get('Hàng KM')))
            buyer = tuple(txt(get(c)) for c in ['Tên khách hàng', 'Địa chỉ', 'MST/CCCD chủ hộ', 'Người mua hàng', 'Hình thức TT', 'Ngày hóa đơn'])
            detail[key].append({'line': line, 'full': full, 'buyer': buyer, 'row': rn, 'issue_status': get('Trạng thái phát hành')})
        wb.close()
    assert len(detail) == 37012 and sum(map(len, detail.values())) == 47187
    assert all(x['issue_status'] == 'Đã cấp mã' for d in detail.values() for x in d)
    header_keys = set()
    for entry in previous['source_manifest']:
        wb = load_workbook(BytesIO((ROOT / entry['file']).read_bytes()), read_only=True, data_only=True)
        for r in wb.active.iter_rows(min_row=6, values_only=True):
            if isinstance(r[0], (int, float)):
                header_keys.add(f'{r[1]}/{str(r[2]).zfill(8)}')
        wb.close()
    assert header_keys == set(detail)
    block_line_count = 0
    for p in resolved['duplicate_pairs']:
        a, b = (detail[x['key']] for x in p['invoices'])
        assert Counter(x['full'] for x in a) == Counter(x['full'] for x in b)
        assert Counter(x['buyer'] for x in a) == Counter(x['buyer'] for x in b)
        block_line_count += len(a)
    assert block_line_count == 2720
    plan = defaultdict(list)
    wb = load_workbook(ROOT / previous['plan']['file'], read_only=True, data_only=True)
    for r in wb.active.iter_rows(min_row=10, values_only=True):
        if r[0] is not None:
            plan[r[18]].append((txt(r[12]), txt(r[13]), number(r[14]), number(r[15]), number(r[16])))
    wb.close()
    matches = rounding = promotions = 0
    rounded_docs = set()
    for p in resolved['coverage'] + resolved['blank_payments']:
        expected = Counter(plan[p['doc']])
        for invoice in p['invoices']:
            actual = Counter(x['line'] for x in detail[invoice['key']])
            if expected == actual:
                matches += 1
                continue
            old, new = list((expected - actual).elements()), list((actual - expected).elements())
            if len(old) == len(new) == 1:
                a, b = old[0], new[0]
                if all(a[i] == b[i] for i in [0, 1, 2, 4]) and a[3] is not None and b[3] == a[3].quantize(Decimal(1), rounding=ROUND_HALF_UP):
                    rounding += 1
                    rounded_docs.add(p['doc'])
                    continue
                lines = detail[invoice['key']]
                if a[3] == a[4] == 0 and b[0] == a[0] + ' (Hàng khuyến mại không thu tiền)' and a[1:3] == b[1:3] and len(lines) == 1 and lines[0]['full'][7] == 0 and lines[0]['full'][8] == 'Khuyến mại':
                    promotions += 1
                    continue
            raise AssertionError((p['doc'], invoice['key'], old, new))
    assert (matches, rounding, promotions, len(rounded_docs)) == (36701, 36, 2, 33)
    output = ROOT / 'outputs/kiem-tra-xuat-moi-20261101/KIEM_TRA_XUAT_MOI_2024.xlsx'
    wb = load_workbook(output, read_only=True, data_only=True)
    sh = wb['Kiem tra trung']
    for r, p in zip(sh.iter_rows(min_row=15, max_row=2014, max_col=19, values_only=True), pairs):
        assert r[1] == p['doc'] and r[3] == p['first_invoice'] and r[4] == p['second_invoice']
        assert r[13] == p['first_lines'] and r[14] == 'Trùng hoàn toàn' and r[15] == p['first_issue_times'][0] and r[16] == p['second_issue_times'][0] and r[18] == 'Trùng'
    other = wb['Don gia va hang KM']
    rows_out = list(other.iter_rows(min_row=5, max_row=42, max_col=10, values_only=True))
    assert Counter(r[3] for r in rows_out) == {'Đơn giá làm tròn': 36, 'Hàng khuyến mại': 2}
    for r, p in zip(rows_out[:36], summary['price_rounding']['items']):
        assert r[1] == p['doc'] and r[2] == p['invoice'] and r[4] == p['item'] and r[7] == p['issued_price'] and r[8] == p['goods']
    assert all(r[7] is None and r[8] is None for r in rows_out[-2:])
    assert sh['D5'].value == 2000 and sh['D8'].value == 3114386753
    wb.close()
    result = {'passed': True, 'independent_reader': 'openpyxl', 'detail_invoices': len(detail), 'detail_rows': 47187, 'duplicate_pairs_identical': 2000,
              'lines_each_block': block_line_count, 'plan_candidates_exact': matches, 'rounded_price_invoices': rounding, 'rounded_price_documents': len(rounded_docs),
              'promotion_zero_payment_evidence': promotions, 'report_readback': True}
    (WORK / 'detail-verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()
