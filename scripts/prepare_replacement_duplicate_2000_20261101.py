"""Prepare replacement-to-zero imports for the second verified duplicate batch."""
from collections import Counter, defaultdict
from datetime import date
from io import BytesIO
from pathlib import Path
import hashlib
import json

from openpyxl import load_workbook
from prepare_invoice_reports import rows, invoice_key, invoice_day, amount

ROOT = Path(__file__).resolve().parents[1]
AUDIT = ROOT / 'scripts/data/new-sales-issued-20261101'
WORK = ROOT / 'scripts/data/replacement-duplicate-2000-20261101'
TEMPLATE = ROOT / 'data/Mau_hoa_don_GTGT_MTT_1_thue_suat_Thay_the.xls'
FIRST = {f'1C26MUV/{n:08}' for n in range(131274, 133274)}
TARGET = {f'1C26MUV/{n:08}' for n in range(133274, 135274)}


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def txt(v):
    return '' if v is None else str(v).strip()


def main():
    WORK.mkdir(parents=True, exist_ok=True)
    summary = json.loads((AUDIT / 'summary.json').read_text(encoding='utf-8'))
    detail_audit = json.loads((AUDIT / 'detail-summary.json').read_text(encoding='utf-8'))
    pairs = json.loads((AUDIT / 'resolved-coverage.json').read_text(encoding='utf-8'))['duplicate_pairs']
    verification = json.loads((AUDIT / 'detail-verification.json').read_text(encoding='utf-8'))
    assert verification['passed'] and verification['duplicate_pairs_identical'] == 2000
    assert {p['invoices'][0]['key'] for p in pairs} == FIRST
    assert {p['invoices'][1]['key'] for p in pairs} == TARGET
    sources = summary['source_manifest'] + detail_audit['source_manifest']
    assert all(sha(ROOT / p['file']) == p['sha256'] for p in sources)
    heads, details = {}, defaultdict(list)
    for entry in summary['source_manifest']:
        headers = {}
        for rn, r in rows(ROOT / entry['file']):
            if rn == 5:
                headers = {v: k for k, v in r.items() if v}
            if rn < 6 or amount(r.get('A')) is None:
                continue
            get = lambda name: r.get(headers.get(name))
            key = '/'.join(invoice_key(get('Ký hiệu'), get('Số hóa đơn')))
            if key not in TARGET:
                continue
            assert key not in heads
            heads[key] = {'date': invoice_day(get('Ngày hóa đơn')), 'cqt': txt(get('Mã của CQT')),
                          'cqt_status': get('Trạng thái gửi CQT'), 'status': get('Trạng thái hóa đơn'),
                          'money': [amount(get(c)) for c in ['Doanh số bán chưa thuế', 'Tiền thuế GTGT', 'Tổng tiền']],
                          'source_file': entry['file'], 'source_row': rn}
    for entry in detail_audit['source_manifest']:
        headers = {}
        for rn, r in rows(ROOT / entry['file']):
            if rn == 5:
                headers = {v: k for k, v in r.items() if v}
            if rn < 6 or amount(r.get('A')) is None:
                continue
            get = lambda name: r.get(headers.get(name))
            key = '/'.join(invoice_key(get('Ký hiệu'), get('Số hóa đơn')))
            if key not in TARGET | FIRST:
                continue
            details[key].append({'name': get('Tên hàng'), 'unit': get('ĐVT'), 'quantity': amount(get('Số lượng')),
                                 'price': amount(get('Đơn giá')), 'goods': amount(get('Thành tiền')), 'vat': amount(get('Tiền thuế GTGT')),
                                 'payment': amount(get('Tổng tiền TT')), 'rate': get('Thuế suất (%)'), 'promotion': get('Hàng KM'), 'kind': get('Tính chất HHDV'),
                                 'customer': get('Tên khách hàng'), 'buyer': get('Người mua hàng'), 'tax': get('MST/CCCD chủ hộ'),
                                 'address': get('Địa chỉ'), 'identity': get('Căn cước công dân'), 'payment_method': get('Hình thức TT'),
                                 'date': invoice_day(get('Ngày hóa đơn')), 'issue_status': get('Trạng thái phát hành'),
                                 'source_file': entry['file'], 'source_row': rn})
    assert set(heads) == TARGET and set(details) == TARGET | FIRST
    assert len({h['cqt'] for h in heads.values()}) == 2000 and all(h['cqt'] for h in heads.values())
    wb = load_workbook(BytesIO(TEMPLATE.read_bytes()), read_only=True, data_only=True)
    header = list(next(wb.active.iter_rows(min_row=9, max_row=9, max_col=22, values_only=True)))
    notes = [next(wb.active.iter_rows(min_row=i, max_row=i, max_col=1, values_only=True))[0] for i in range(1, 8)]
    wb.close()
    header += ['Lý do thay thế', 'Là dòng ghi chú/diễn giải']
    new_date = date(2026, 11, 1)
    notes[6] = '- Ngày hóa đơn 01/11/2026. Mỗi hóa đơn một dòng ghi chú/diễn giải, lý do thay thế giống Tên hàng hóa/dịch vụ; số lượng và các khoản tiền bằng 0, thuế suất 8%.'
    output_rows, audit = [], []
    comparison_fields = ['name', 'unit', 'quantity', 'price', 'goods', 'vat', 'payment', 'rate', 'promotion', 'kind']
    buyer_fields = ['customer', 'buyer', 'tax', 'address', 'identity', 'payment_method']
    for seq, key in enumerate(sorted(TARGET), 1):
        retained = f'1C26MUV/{int(key.split("/")[1])-2000:08}'
        h, ls = heads[key], details[key]
        assert h['date'] == date(2026, 11, 1) and h['cqt_status'] == 'HĐ hợp lệ' and h['status'] == 'Hoá đơn mới'
        assert Counter(tuple(r[f] for f in comparison_fields) for r in ls) == Counter(tuple(r[f] for f in comparison_fields) for r in details[retained])
        assert len({tuple(r[f] for f in buyer_fields) for r in ls}) == 1
        assert [sum(r[f] for r in ls) for f in ['goods', 'vat', 'payment']] == h['money']
        start = len(output_rows) + 10
        for r in ls:
            assert r['date'] == date(2026, 11, 1) and r['rate'] == '8%' and r['name'] and r['unit']
            assert r['kind'] == 'Hàng hóa, dịch vụ' and not r['promotion'] and r['issue_status'] == 'Đã cấp mã'
        r = ls[0]
        reason = f"Thay thế hóa đơn số {key.split('/')[1]} ngày {h['date'].strftime('%d/%m/%Y')} do xuất trùng đơn hàng"
        row = [None] * 24
        row[0], row[1] = seq, (new_date - date(1899, 12, 30)).days
        row[2], row[3], row[4], row[5] = r['customer'] or None, r['tax'] or None, r['address'] or None, r['buyer'] or None
        row[8], row[9] = r['identity'] or None, r['payment_method'] or None
        row[11], row[12] = '1C26MUV', key.split('/')[1]
        row[13], row[14] = (h['date'] - date(1899, 12, 30)).days, h['cqt']
        row[15] = reason
        row[17], row[18], row[19], row[20], row[21] = 0, 0, 0, 8, 0
        row[22], row[23] = reason, 1
        output_rows.append(row)
        audit.append({'sequence': seq, 'replaced': key, 'retained': retained, 'replaced_date': '2026-11-01',
                      'cqt': h['cqt'], 'lines': 1, 'source_lines': len(ls), 'first_output_row': start, 'last_output_row': len(output_rows)+9,
                      'detail_source_file': ls[0]['source_file'], 'detail_source_rows': [r['source_row'] for r in ls],
                      'before': [int(v) for v in h['money']], 'replacement': [0, 0, 0]})
    totals = [sum(a['before'][i] for a in audit) for i in range(3)]
    assert len(audit) == len(output_rows) == 2000 and sum(a['source_lines'] for a in audit) == 2720
    assert totals == [2883691447, 230695306, 3114386753]
    prepared = {'headers': header, 'notes': notes, 'rows': output_rows, 'audit': audit, 'template': str(TEMPLATE), 'template_sha256': sha(TEMPLATE),
                'source_manifest': sources, 'invoices': 2000, 'lines': 2000, 'source_lines': 2720, 'before': totals, 'replacement_totals': [0, 0, 0],
                'new_dates_blank': False, 'new_date': new_date.isoformat(), 'one_note_line_per_invoice': True, 'note_flag': 1,
                'reference_date': '2026-11-01', 'reference_series': '1C26MUV', 'first_reference': '00133274', 'last_reference': '00135273',
                'reason': 'Thay thế về 0 lô hóa đơn phát hành lặp lần sau; giữ lô 00131274–00133273.',
                'import_help': 'https://helpv4.meinvoice.vn/kb/nhap-khau-hoa-don-thay-the/'}
    (WORK / 'prepared.json').write_text(json.dumps(prepared, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({k: prepared[k] for k in ['invoices', 'lines', 'before', 'replacement_totals', 'first_reference', 'last_reference', 'new_dates_blank']}))


if __name__ == '__main__':
    main()
