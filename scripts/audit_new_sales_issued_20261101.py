"""Read-only reconciliation of the user's two-day invoice export and 2024 sales plan."""
from __future__ import annotations

import csv
import hashlib
import json
import sys
from collections import Counter, defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from openpyxl import load_workbook
from prepare_invoice_reports import amount, invoice_day, invoice_key, rows

ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = Path(r'C:\Users\thuong ngo\Downloads\Bang_ke_hoa_don_da_su_dung_1793518718337.zip')
SOURCE = ROOT / 'data/Meinv_SAU_XUAT_MOI_20261101'
WORK = ROOT / 'scripts/data/new-sales-issued-20261101'
PLAN = ROOT / 'outputs/amis-2024-theo-btl-bhh-20261026/IMPORT_HOA_DON_MOI_AMIS_2024.xlsx'
PERIOD = {date(2026, 10, 31), date(2026, 11, 1)}
MONEY = ['Doanh số bán chưa thuế', 'Tiền thuế GTGT', 'Tổng tiền']


def native(v):
    if isinstance(v, Decimal):
        return int(v) if v == int(v) else str(v)
    if isinstance(v, (date, Path)):
        return str(v)
    if isinstance(v, set):
        return sorted(v)
    raise TypeError(type(v))


def dump(name, value):
    (WORK / name).write_text(json.dumps(value, ensure_ascii=False, indent=2, default=native) + '\n', encoding='utf-8')


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def txt(v):
    return '' if v is None else str(v).strip()


def totals(records):
    return [sum((r['money'][i] for r in records if r['money'][i] is not None), Decimal(0)) for i in range(3)]


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    WORK.mkdir(parents=True, exist_ok=True)
    SOURCE.mkdir(parents=True, exist_ok=True)
    manifest = []
    with ZipFile(ARCHIVE) as z:
        assert z.testzip() is None, 'ZIP CRC failed'
        for entry in z.infolist():
            if entry.is_dir():
                continue
            name = PurePosixPath(entry.filename.replace('\\', '/'))
            assert len(name.parts) == 1 and name.suffix.lower() in {'.xls', '.xlsx'}, entry.filename
            target = (SOURCE / name.name).resolve()
            assert target.parent == SOURCE.resolve()
            content = z.read(entry)
            if target.exists():
                assert target.read_bytes() == content, f'Existing source differs: {target}'
            else:
                target.write_bytes(content)
            manifest.append({'file': target.relative_to(ROOT).as_posix(), 'bytes': len(content), 'sha256': sha(target)})

    actual, file_stats, footer_controls = [], [], []
    for entry in manifest:
        path = ROOT / entry['file']
        headers, indices, local = {}, [], []
        period_label = None
        for rn, r in rows(path):
            if rn == 3:
                period_label = r.get('A')
            if rn == 5:
                headers = {v: k for k, v in r.items() if v}
                assert set(['STT', 'Ký hiệu', 'Số hóa đơn', 'Ngày hóa đơn', 'Mã đơn hàng', *MONEY]).issubset(headers)
            if rn < 6:
                continue
            if amount(r.get('A')) is None:
                if any(amount(r.get(headers.get(c))) is not None for c in MONEY):
                    footer_controls.append({'file': entry['file'], 'row': rn, 'label': r.get('A'), 'money': [amount(r.get(headers[c])) for c in MONEY]})
                continue
            get = lambda k: r.get(headers.get(k))
            key = invoice_key(get('Ký hiệu'), get('Số hóa đơn'))
            assert key, (entry['file'], rn)
            record = {
                'key': '/'.join(key), 'date': invoice_day(get('Ngày hóa đơn')),
                'order': txt(get('Mã đơn hàng')), 'money': [amount(get(c)) for c in MONEY],
                'goods_before_discount': amount(get('Tổng tiền hàng')), 'discount': amount(get('Tiền chiết khấu')),
                'invoice_status': txt(get('Trạng thái hóa đơn')), 'cqt_status': txt(get('Trạng thái gửi CQT')),
                'cqt_code_present': bool(txt(get('Mã của CQT'))),
                'source': {'file': entry['file'], 'row': rn},
            }
            indices.append(int(amount(r['A'])))
            local.append(record)
        assert indices == list(range(1, len(indices) + 1)), (path.name, 'STT is not complete')
        actual.extend(local)
        file_stats.append({'file': entry['file'], 'period': period_label, 'invoices': len(local), 'money': totals(local), 'dates': dict(Counter(str(r['date']) for r in local))})

    # Read precisely the approved workbook. No source workbook is modified.
    wb = load_workbook(PLAN, read_only=True, data_only=True)
    sheet = wb['Hóa đơn GTGT']
    plan_by_seq = {}
    row_docs = {}
    lines = 0
    for rn, r in enumerate(sheet.iter_rows(min_row=10, values_only=True), 10):
        if r[0] is None:
            continue
        seq, doc = int(r[0]), txt(r[18])
        assert doc and amount(r[16]) is not None, ('invalid plan row', rn)
        if seq not in plan_by_seq:
            plan_by_seq[seq] = {'seq': seq, 'doc': doc, 'original_date': invoice_day(r[1]), 'money': [Decimal(0), Decimal(0), Decimal(0)], 'first_row': rn, 'lines': 0}
        p = plan_by_seq[seq]
        assert p['doc'] == doc, (seq, doc)
        p['money'][0] += amount(r[16])
        if r[11] is not None:
            assert amount(r[11]) is not None
            p['money'][1] += amount(r[11])
        p['lines'] += 1
        row_docs[rn] = doc
        lines += 1
    wb.close()
    plan = list(plan_by_seq.values())
    assert sorted(plan_by_seq) == list(range(1, len(plan) + 1))
    assert len({p['doc'] for p in plan}) == len(plan), 'Repeated plan document'
    for p in plan:
        p['money'][2] = p['money'][0] + p['money'][1]

    # The live import's T/U columns are order/document, verified against its headers.
    mapping = defaultdict(set)
    mapping_rows, mapping_errors = {}, []
    snapshot_manifest = []
    for path in sorted(WORK.glob('google-import-map-*.json')):
        data = json.loads(path.read_text(encoding='utf-8'))
        snapshot_manifest.append({'file': path.relative_to(ROOT).as_posix(), 'sha256': sha(path), 'range': data['range']})
        for rn, r in enumerate(data['values'], data['first_row']):
            order, doc = (txt(r[0]) if len(r) > 0 else ''), (txt(r[1]) if len(r) > 1 else '')
            assert rn not in mapping_rows, ('overlapping snapshot', rn)
            mapping_rows[rn] = doc
            if doc != row_docs.get(rn):
                mapping_errors.append({'row': rn, 'live_doc': doc, 'local_doc': row_docs.get(rn)})
            mapping[doc].add(order)
    live_doc_counts, local_doc_counts = Counter(mapping_rows.values()), Counter(row_docs.values())
    assert live_doc_counts == local_doc_counts, ('live/local import documents or line counts differ', list((live_doc_counts - local_doc_counts).items())[:10], list((local_doc_counts - live_doc_counts).items())[:10])
    for p in plan:
        codes = mapping[p['doc']]
        assert len(codes) == 1 and next(iter(codes)) and not next(iter(codes)).startswith('#'), (p['doc'], codes)
        p['order'] = next(iter(codes))

    pa, aa = defaultdict(list), defaultdict(list)
    for p in plan:
        pa[p['order']].append(p)
    for a in actual:
        aa[a['order']].append(a)
    matches, ambiguous, missing, mismatches = [], [], [], []
    for order, expected in pa.items():
        found = aa.get(order, [])
        if not found:
            missing.extend(expected)
            continue
        if len(expected) == len(found) == 1:
            p, a = expected[0], found[0]
            result = {'doc': p['doc'], 'seq': p['seq'], 'order': order, 'expected': p['money'], 'actual': a['money'], 'delta': [None if a['money'][i] is None else a['money'][i] - p['money'][i] for i in range(3)], 'invoice': a}
            (matches if a['money'] == p['money'] else mismatches).append(result)
        else:
            ambiguous.append({'order': order, 'expected': expected, 'actual': found, 'expected_count': len(expected), 'actual_count': len(found)})

    unrelated = [a for a in actual if a['order'] not in pa]
    # Resolve repeated order codes by exact amount tuples, preserving all equally
    # matching invoices instead of selecting one arbitrarily for return references.
    coverage, duplicate_pairs, blank_payments, true_differences = [], [], [], []
    for order, expected in pa.items():
        expected_amounts = Counter(tuple(p['money']) for p in expected)
        assert max(expected_amounts.values()) == 1, ('Ambiguous plan fingerprint', order)
        for p in expected:
            exact = [a for a in aa.get(order, []) if a['money'] == p['money']]
            blank = [a for a in aa.get(order, []) if a['money'][:2] == p['money'][:2] and a['money'][2] is None]
            base = {'doc': p['doc'], 'seq': p['seq'], 'order': order, 'expected': p['money']}
            if exact:
                coverage.append({**base, 'invoices': exact})
                if len(exact) > 1:
                    duplicate_pairs.append({**base, 'invoices': sorted(exact, key=lambda a: a['key']), 'additional_invoices_vs_plan': len(exact) - 1})
            elif blank:
                blank_payments.append({**base, 'invoices': blank})
            else:
                true_differences.append({**base, 'invoices': aa.get(order, [])})
    duplicate_pairs.sort(key=lambda p: p['invoices'][0]['key'])
    duplicate_money = [sum((p['expected'][i] * p['additional_invoices_vs_plan'] for p in duplicate_pairs), Decimal(0)) for i in range(3)]
    duplication_blocks = []
    if duplicate_pairs and all(len(p['invoices']) == 2 for p in duplicate_pairs):
        first = [int(p['invoices'][0]['key'].split('/')[1]) for p in duplicate_pairs]
        second = [int(p['invoices'][1]['key'].split('/')[1]) for p in duplicate_pairs]
        duplication_blocks = [{'first_min': min(first), 'first_max': max(first), 'second_min': min(second), 'second_max': max(second),
                               'first_contiguous': first == list(range(min(first), max(first) + 1)), 'second_contiguous': second == list(range(min(second), max(second) + 1)),
                               'number_offsets': dict(Counter(str(b - a) for a, b in zip(first, second)))}]
    matched_actual = [m['invoice'] for m in matches]
    key_counts = Counter(r['key'] for r in actual)
    status_groups = Counter((r['invoice_status'], r['cqt_status'], r['cqt_code_present']) for r in actual)
    numeric_order_types = Counter('digits' if r['order'].isdigit() else 'other' if r['order'] else 'blank' for r in actual)
    expected_summary = json.loads((ROOT / 'scripts/data/amis-2024-linked-20261026/summary.json').read_text(encoding='utf-8'))['summary']['Đưa vào import']
    assert len(plan) == expected_summary['documents'] and lines == expected_summary['import_lines']
    assert totals(plan) == [Decimal(expected_summary[k]) for k in ['goods', 'vat', 'payment']]
    summary = {
        'archive': {'path': str(ARCHIVE), 'sha256': sha(ARCHIVE)},
        'source_manifest': manifest, 'plan': {'file': PLAN.relative_to(ROOT).as_posix(), 'sha256': sha(PLAN), 'invoices': len(plan), 'lines': lines, 'money': totals(plan)},
        'live_mapping': {'url': 'https://docs.google.com/spreadsheets/d/1d4fLQRALDLXOFOQQlZuXMCTAg4lnRQ7NbKH-UPJN0XE/edit', 'sheet': 'Hóa đơn GTGT', 'range': 'T10:U44198', 'rows_checked_against_local_plan_by_document_and_line_count': len(mapping_rows), 'different_row_positions': len(mapping_errors), 'snapshots': snapshot_manifest},
        'export': {'invoices': len(actual), 'money': totals(actual), 'files': file_stats, 'footer_controls': footer_controls, 'dates': dict(Counter(str(r['date']) for r in actual)),
                   'statuses': [{'invoice_status': s[0], 'cqt_status': s[1], 'cqt_code_present': s[2], 'count': n} for s, n in status_groups.items()], 'order_types': dict(numeric_order_types)},
        'reconciliation': {'exact_unique_matches': len(matches), 'matched_money': totals(matched_actual), 'money_mismatches': len(mismatches), 'missing_documents': len(missing), 'ambiguous_orders': len(ambiguous), 'ambiguous_plan_documents': sum(len(x['expected']) for x in ambiguous), 'outside_plan_invoices': len(unrelated), 'outside_plan_money': totals(unrelated),
                           'matched_dates': dict(Counter(str(r['date']) for r in matched_actual)), 'matched_cqt_statuses': dict(Counter(r['cqt_status'] for r in matched_actual)), 'matched_invoice_statuses': dict(Counter(r['invoice_status'] for r in matched_actual)), 'matched_missing_cqt_codes': sum(not r['cqt_code_present'] for r in matched_actual),
                           'duplicate_invoice_keys': {k: n for k, n in key_counts.items() if n > 1}, 'outside_declared_date_period': [r['key'] for r in actual if r['date'] not in PERIOD], 'plan_duplicate_orders': {k: len(v) for k, v in pa.items() if len(v) > 1}},
        'resolved_by_order_and_exact_amount': {'plan_documents_with_exact_money': len(coverage), 'plan_documents_with_unknown_payment_only': len(blank_payments), 'plan_documents_with_other_money_or_missing_issues': len(true_differences),
            'duplicate_groups': len(duplicate_pairs), 'additional_invoices_vs_plan': sum(p['additional_invoices_vs_plan'] for p in duplicate_pairs), 'additional_money_vs_plan': duplicate_money, 'duplication_blocks': duplication_blocks,
            'covered_plan_dates': dict(Counter(str(p['invoices'][0]['date']) for p in coverage + blank_payments)),
            'all_plan_candidate_statuses': dict(Counter(a['cqt_status'] for p in coverage + blank_payments for a in p['invoices'])),
            'all_plan_candidates_have_cqt_code': all(a['cqt_code_present'] for p in coverage + blank_payments for a in p['invoices']),
            'blank_payment_invoices': blank_payments},
        'limits': ['Bảng kê tổng kiểm tra số lượng, mã đơn và tiền ở cấp hóa đơn; chưa có bảng kê chi tiết mặt hàng.', 'Trạng thái và mã CQT theo bảng kê nguồn; chưa tra cứu CQT trực tiếp.', 'Không sửa nguồn, bộ import hoặc Google Sheets; chưa xử lý hoặc xác nhận phát hành phần trả hàng.'],
    }
    dump('summary.json', summary)
    dump('matched.json', matches)
    dump('exceptions.json', {'missing': missing, 'money_mismatches': mismatches, 'ambiguous': ambiguous, 'outside_plan': unrelated})
    dump('resolved-coverage.json', {'coverage': coverage, 'duplicate_pairs': duplicate_pairs, 'blank_payments': blank_payments, 'other_money_issues': true_differences})
    output = ROOT / 'outputs/kiem-tra-xuat-moi-20261101'
    output.mkdir(parents=True, exist_ok=True)
    with (output / 'CAC_CAP_TRUNG_MA_DON_VA_TIEN.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Số chứng từ AMIS', 'STT import', 'Mã đơn hàng', 'HĐ số nhỏ', 'Ngày HĐ số nhỏ', 'HĐ số lớn', 'Ngày HĐ số lớn', 'Tiền hàng mỗi HĐ', 'VAT mỗi HĐ', 'Thanh toán mỗi HĐ', 'Trạng thái CQT HĐ số nhỏ', 'Trạng thái CQT HĐ số lớn', 'File bảng kê', 'Dòng HĐ số nhỏ', 'Dòng HĐ số lớn'])
        for p in duplicate_pairs:
            assert len(p['invoices']) == 2
            a, b = p['invoices']
            writer.writerow([p['doc'], p['seq'], p['order'], a['key'], a['date'], b['key'], b['date'], *p['expected'], a['cqt_status'], b['cqt_status'], a['source']['file'], a['source']['row'], b['source']['row']])
    with (WORK / 'matched-invoices.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Số chứng từ AMIS', 'STT import', 'Mã đơn hàng', 'Ký hiệu/Số hóa đơn', 'Ngày hóa đơn', 'Tiền hàng', 'VAT', 'Thanh toán', 'Trạng thái gửi CQT', 'Có mã CQT', 'File nguồn', 'Dòng nguồn'])
        for m in matches:
            a = m['invoice']
            writer.writerow([m['doc'], m['seq'], m['order'], a['key'], a['date'], *a['money'], a['cqt_status'], a['cqt_code_present'], a['source']['file'], a['source']['row']])
    assert all(sha(ROOT / e['file']) == e['sha256'] for e in manifest)
    assert sha(PLAN) == summary['plan']['sha256']
    for entry in manifest:
        controls = [c for c in footer_controls if c['file'] == entry['file']]
        assert len(controls) == 1 and controls[0]['money'] == next(s['money'] for s in file_stats if s['file'] == entry['file']), ('footer differs', entry)
    print(json.dumps({k: v for k, v in summary.items() if k in {'resolved_by_order_and_exact_amount'}}, ensure_ascii=False, indent=2, default=native))


if __name__ == '__main__':
    main()
