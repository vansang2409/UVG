"""Compare fresh issued invoice lines with the approved plan and duplicate pairs."""
from __future__ import annotations

import hashlib
import json
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from prepare_invoice_reports import amount, invoice_day, invoice_key, rows

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / 'scripts/data/new-sales-issued-20261101'
ARCHIVE = Path(r'C:\Users\thuong ngo\Downloads\Bang_ke_chi_tiet_hoa_don_da_su_dung_20261101_143931.zip')
SOURCE = ROOT / 'data/Meinv_SAU_XUAT_MOI_20261101/CHI_TIET'


def native(v):
    if isinstance(v, Decimal):
        return int(v) if v == int(v) else str(v)
    if isinstance(v, date):
        return v.isoformat()
    raise TypeError(type(v))


def text(v):
    return '' if v is None else str(v).strip()


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def issue_time(v):
    raw = text(v)
    try:
        return datetime.fromisoformat(raw).isoformat(sep=' ', timespec='seconds')
    except ValueError:
        try:
            return (datetime(1899, 12, 30) + timedelta(seconds=round(float(raw) * 86400))).isoformat(sep=' ', timespec='seconds')
        except ValueError:
            return raw


def dump(name, value):
    (WORK / name).write_text(json.dumps(value, ensure_ascii=False, indent=2, default=native) + '\n', encoding='utf-8')


def tuple_counter(values, fields):
    return Counter(tuple(r[k] for k in fields) for r in values)


def delta_counter(a, b, fields):
    ca, cb = tuple_counter(a, fields), tuple_counter(b, fields)
    return {'expected_only': [list(k) + [n] for k, n in (ca - cb).items()], 'actual_only': [list(k) + [n] for k, n in (cb - ca).items()]}


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    SOURCE.mkdir(parents=True, exist_ok=True)
    previous = json.loads((WORK / 'summary.json').read_text(encoding='utf-8'))
    resolved = json.loads((WORK / 'resolved-coverage.json').read_text(encoding='utf-8'))
    manifest = []
    with ZipFile(ARCHIVE) as z:
        assert z.testzip() is None
        for entry in z.infolist():
            if entry.is_dir():
                continue
            name = PurePosixPath(entry.filename.replace('\\', '/'))
            assert len(name.parts) == 1 and name.suffix.lower() in {'.xls', '.xlsx'}
            target = (SOURCE / name.name).resolve()
            assert target.parent == SOURCE.resolve()
            content = z.read(entry)
            if target.exists():
                assert target.read_bytes() == content
            else:
                target.write_bytes(content)
            manifest.append({'file': target.relative_to(ROOT).as_posix(), 'bytes': len(content), 'sha256': sha(target)})

    details = defaultdict(list)
    periods, footers, stt = {}, [], []
    for entry in manifest:
        headers = {}
        path = ROOT / entry['file']
        for rn, r in rows(path):
            if rn == 3:
                periods[entry['file']] = r.get('A')
            if rn == 5:
                headers = {v: k for k, v in r.items() if v}
                assert set(['Ký hiệu', 'Số hóa đơn', 'Tên hàng', 'Số lượng', 'Đơn giá', 'Thành tiền']).issubset(headers)
            if rn < 6:
                continue
            get = lambda k: r.get(headers.get(k))
            if amount(r.get('A')) is None:
                if any(amount(get(k)) is not None for k in ['Thành tiền', 'Tiền thuế GTGT', 'Tổng tiền TT']):
                    footers.append({'file': entry['file'], 'row': rn, 'label': r.get('A'), 'money': [amount(get(k)) for k in ['Thành tiền', 'Tiền thuế GTGT', 'Tổng tiền TT']]})
                continue
            stt.append(int(amount(r['A'])))
            key = invoice_key(get('Ký hiệu'), get('Số hóa đơn'))
            assert key
            details['/'.join(key)].append({
                'name': text(get('Tên hàng')), 'unit': text(get('ĐVT')), 'qty': amount(get('Số lượng')),
                'price': amount(get('Đơn giá')), 'goods': amount(get('Thành tiền')), 'rate': text(get('Thuế suất (%)')),
                'vat': amount(get('Tiền thuế GTGT')), 'payment': amount(get('Tổng tiền TT')),
                'kind': text(get('Tính chất HHDV')), 'promotion': text(get('Hàng KM')),
                'date': invoice_day(get('Ngày hóa đơn')), 'buyer': text(get('Người mua hàng')),
                'customer': text(get('Tên khách hàng')), 'address': text(get('Địa chỉ')), 'tax': text(get('MST/CCCD chủ hộ')),
                'payment_method': text(get('Hình thức TT')), 'cqt_status': text(get('Trạng thái gửi CQT')),
                'invoice_status': text(get('Trạng thái HĐ')), 'issue_status': text(get('Trạng thái phát hành')),
                'issue_time': issue_time(get('Thời gian phát hành')), 'issue_time_raw': text(get('Thời gian phát hành')),
                'source': {'file': entry['file'], 'row': rn},
            })
    assert stt == list(range(1, len(stt) + 1)), 'Detail STT incomplete'

    # The approved plan has S=document, one original row for each product line.
    plan = defaultdict(list)
    plan_path = ROOT / previous['plan']['file']
    for rn, r in rows(plan_path):
        if rn < 10 or amount(r.get('A')) is None:
            continue
        plan[text(r.get('S'))].append({'name': text(r.get('M')), 'unit': text(r.get('N')), 'qty': amount(r.get('O')),
                                      'price': amount(r.get('P')), 'goods': amount(r.get('Q')), 'promotion': text(r.get('R')), 'row': rn})
    assert len(plan) == previous['plan']['invoices'] and sum(map(len, plan.values())) == previous['plan']['lines']

    basic_fields = ['name', 'unit', 'qty', 'price', 'goods']
    full_fields = [*basic_fields, 'rate', 'vat', 'payment', 'kind', 'promotion']
    buyer_fields = ['buyer', 'customer', 'address', 'tax', 'payment_method', 'date']
    pair_results = []
    for p in resolved['duplicate_pairs']:
        a, b = p['invoices']
        ar, br = details.get(a['key'], []), details.get(b['key'], [])
        result = {'doc': p['doc'], 'order': p['order'], 'first_invoice': a['key'], 'second_invoice': b['key'],
                  'first_lines': len(ar), 'second_lines': len(br), 'both_present': bool(ar and br),
                  'identical_product_lines': bool(ar and br) and tuple_counter(ar, full_fields) == tuple_counter(br, full_fields),
                  'identical_buyer_information': bool(ar and br) and tuple_counter(ar, buyer_fields) == tuple_counter(br, buyer_fields),
                  'first_matches_plan': bool(ar) and tuple_counter(plan[p['doc']], basic_fields) == tuple_counter(ar, basic_fields),
                  'second_matches_plan': bool(br) and tuple_counter(plan[p['doc']], basic_fields) == tuple_counter(br, basic_fields),
                  'first_issue_times': sorted({r['issue_time'] for r in ar}), 'second_issue_times': sorted({r['issue_time'] for r in br}),
                  'first_source_rows': [r['source']['row'] for r in ar], 'second_source_rows': [r['source']['row'] for r in br],
                  'money': p['expected']}
        if not result['identical_product_lines']:
            result['line_differences'] = delta_counter(ar, br, full_fields)
        pair_results.append(result)

    coverage_results, plan_differences = [], []
    for p in resolved['coverage'] + resolved['blank_payments']:
        for a in p['invoices']:
            ar = details.get(a['key'], [])
            same = bool(ar) and tuple_counter(plan[p['doc']], basic_fields) == tuple_counter(ar, basic_fields)
            result = {'doc': p['doc'], 'invoice': a['key'], 'expected_lines': len(plan[p['doc']]), 'actual_lines': len(ar), 'matches': same}
            coverage_results.append(result)
            if not same:
                result['differences'] = delta_counter(plan[p['doc']], ar, basic_fields)
                plan_differences.append(result)

    # Independently compare all detail invoice totals with the supplied summary.
    header_invoices = {}
    for entry in previous['source_manifest']:
        headers = {}
        for rn, r in rows(ROOT / entry['file']):
            if rn == 5:
                headers = {v: k for k, v in r.items() if v}
            if rn < 6 or amount(r.get('A')) is None:
                continue
            get = lambda k: r.get(headers.get(k))
            key = '/'.join(invoice_key(get('Ký hiệu'), get('Số hóa đơn')))
            header_invoices[key] = {'date': invoice_day(get('Ngày hóa đơn')), 'money': [amount(get(c)) for c in ['Doanh số bán chưa thuế', 'Tiền thuế GTGT', 'Tổng tiền']]}
    total_differences, date_differences = [], []
    for key, h in header_invoices.items():
        ar = details.get(key, [])
        if not ar:
            continue
        observed = []
        for field in ['goods', 'vat', 'payment']:
            values = [r[field] for r in ar]
            observed.append(sum(values, Decimal(0)) if all(v is not None for v in values) else None)
        if observed != h['money']:
            total_differences.append({'invoice': key, 'summary': h['money'], 'detail': observed, 'lines': len(ar)})
        if {r['date'] for r in ar} != {h['date']}:
            date_differences.append(key)

    blank_evidence = []
    for p in resolved['blank_payments']:
        for a in p['invoices']:
            blank_evidence.append({'doc': p['doc'], 'invoice': a['key'], 'summary_money': a['money'], 'detail_rows': details.get(a['key'], [])})
    price_rounding, promotion_changes, other_changes = [], [], []
    zero_invoice_keys = {a['key'] for p in resolved['blank_payments'] for a in p['invoices']}
    for result in plan_differences:
        old, new = result['differences']['expected_only'], result['differences']['actual_only']
        if len(old) == len(new) == 1:
            a, b = old[0], new[0]
            unchanged = all(a[i] == b[i] for i in [0, 1, 2, 4, 5])
            if unchanged and amount(a[3]) is not None and amount(b[3]) == amount(a[3]).quantize(Decimal(1), rounding=ROUND_HALF_UP):
                price_rounding.append({**result, 'item': a[0], 'qty': a[2], 'plan_price': a[3], 'issued_price': b[3], 'goods': a[4]})
                continue
        if result['invoice'] in zero_invoice_keys:
            ar = details[result['invoice']]
            if len(ar) == 1 and ar[0]['kind'] == 'Khuyến mại' and ar[0]['payment'] == 0:
                p = plan[result['doc']]
                if len(p) == 1 and p[0]['price'] == p[0]['goods'] == 0 and p[0]['qty'] == ar[0]['qty'] and p[0]['unit'] == ar[0]['unit'] and ar[0]['name'] == p[0]['name'] + ' (Hàng khuyến mại không thu tiền)':
                    promotion_changes.append({**result, 'item': ar[0]['name'], 'qty': ar[0]['qty'], 'payment_evidence': 0})
                    continue
        other_changes.append(result)
    outside = json.loads((WORK / 'exceptions.json').read_text(encoding='utf-8'))['outside_plan']
    summary = {
        'archive': {'path': str(ARCHIVE), 'sha256': sha(ARCHIVE)}, 'source_manifest': manifest, 'periods': periods,
        'detail_rows': len(stt), 'detail_invoices': len(details), 'summary_invoices': len(header_invoices), 'footer_controls': footers,
        'detail_missing_vs_summary': sorted(set(header_invoices) - set(details)), 'detail_extra_vs_summary': sorted(set(details) - set(header_invoices)),
        'date_differences': date_differences, 'invoice_total_differences': total_differences,
        'duplicate_pair_count': len(pair_results), 'pairs_with_both_invoices': sum(p['both_present'] for p in pair_results),
        'pairs_with_identical_lines': sum(p['identical_product_lines'] for p in pair_results), 'pairs_with_identical_buyer_information': sum(p['identical_buyer_information'] for p in pair_results),
        'pairs_where_both_match_plan': sum(p['first_matches_plan'] and p['second_matches_plan'] for p in pair_results),
        'duplicate_block_line_counts': {'first': sum(p['first_lines'] for p in pair_results), 'second': sum(p['second_lines'] for p in pair_results)},
        'plan_invoice_candidates_checked': len(coverage_results), 'plan_invoice_candidates_matching_lines': sum(p['matches'] for p in coverage_results), 'plan_line_differences': plan_differences,
        'detail_issue_statuses': dict(Counter(r['issue_status'] for ar in details.values() for r in ar)),
        'price_rounding': {'invoices': len(price_rounding), 'documents': len({x['doc'] for x in price_rounding}), 'items': price_rounding},
        'promotion_representation_changes': promotion_changes, 'other_plan_changes': other_changes,
        'outside_plan_dates': dict(Counter(a['date'] for a in outside)),
        'team_2026_user_confirmation': 'Người dùng xác nhận ngày31/10/2026 có team2026 xuất và phát hành đơn2026. Giữ273HĐ ngoài kế hoạch2024 riêng; chưa đối chiếu từng HĐ đó với sổ bán2026.',
        'duplicate_issue_time_ranges': {label: {'min': min(t for p in pair_results for t in p[label]), 'max': max(t for p in pair_results for t in p[label])} for label in ['first_issue_times', 'second_issue_times']},
        'blank_payment_evidence': blank_evidence,
        'amount_tuple_order': ['Tiền hàng', 'VAT', 'Thanh toán'],
        'limits': ['Đối chiếu từ hai bảng kê người dùng xuất; chưa tra cứu trực tiếp CQT hoặc kiểm chữ ký XML.', 'Chưa tự chọn hóa đơn giữ lại trong các cặp trùng hoặc thực hiện xử lý hóa đơn.'],
    }
    dump('detail-summary.json', summary)
    dump('detail-pairs.json', pair_results)
    dump('detail-plan-comparison.json', coverage_results)
    dump('detail-lines.json', details)
    assert sha(plan_path) == previous['plan']['sha256']
    assert all(sha(ROOT / e['file']) == e['sha256'] for e in manifest)
    display = {k: v for k, v in summary.items() if k not in {'source_manifest', 'archive', 'periods', 'blank_payment_evidence', 'plan_line_differences', 'invoice_total_differences', 'price_rounding', 'promotion_representation_changes'}}
    display['price_rounding'] = {k: v for k, v in summary['price_rounding'].items() if k != 'items'}
    display['plan_line_differences_count'] = len(plan_differences)
    display['invoice_total_differences_count'] = len(total_differences)
    display['plan_line_difference_samples'] = plan_differences[:5]
    display['invoice_total_difference_samples'] = total_differences[:5]
    display['blank_payment_evidence'] = blank_evidence
    print(json.dumps(display, ensure_ascii=False, indent=2, default=native))


if __name__ == '__main__':
    main()
