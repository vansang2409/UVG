"""Fill only dates and original-invoice references in the live return import."""
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path
import hashlib
import json

from prepare_invoice_reports import rows, invoice_key, invoice_day, amount

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / 'scripts/data/return-adjustment-issued-20261101'
AUDIT = ROOT / 'scripts/data/new-sales-issued-20261101'
SHEET_ID = 700000001


def value(cell):
    data = cell.get('effectiveValue', cell.get('userEnteredValue', {}))
    assert 'errorValue' not in data, data
    return next(iter(data.values()), None)


def encode(v):
    if v is None:
        return {}
    return {'userEnteredValue': {'numberValue': v} if isinstance(v, (int, float)) else {'stringValue': v}}


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def main():
    snapshot = json.loads((WORK / 'before-cells.json').read_text(encoding='utf-8'))
    sheet = next(s for s in snapshot['sheets'] if s['properties']['sheetId'] == SHEET_ID)
    assert sheet['properties']['title'] == 'Import dieu chinh'
    original = sheet['data'][0]['rowData']
    data = [[value(c) for c in r.get('values', [])] for r in original]
    data = [r + [None] * (26 - len(r)) for r in data]
    assert len(data) == 228
    header, body = data[8], data[9:]
    assert header[1] == 'Ngày hóa đơn' and header[9:13] == ['Ký hiệu HĐ bị điều chỉnh (*)', 'Số hóa đơn bị điều chỉnh (*)', 'Ngày hóa đơn bị điều chỉnh', 'Mã của CQT trên HĐ bị điều chỉnh']
    baseline = json.loads((ROOT / 'scripts/data/amis-2024-linked-20261026/return-adjustment-import.json').read_text(encoding='utf-8'))
    # The user changed buyer fields and sequence order after the older local file.
    # All source return amounts, documents and product lines must still match.
    for live, local in zip(body, baseline['rows']):
        assert all(live[i] == local[i] for i in range(26) if i not in [0, 2, 5]), ('return plan differs', live[25])
    assert len(body) == 219 and len({r[0] for r in body}) == len({r[25] for r in body}) == 163
    assert [sum(r[i] or 0 for r in body) for i in [18, 21, 22]] == [-225865926, -18069275, -243935201]
    resolved = json.loads((AUDIT / 'resolved-coverage.json').read_text(encoding='utf-8'))
    candidates = {p['doc']: p['invoices'] for p in resolved['coverage'] + resolved['blank_payments']}
    needed = {r[25] for r in body}
    assert needed <= candidates.keys() and all(len(candidates[doc]) == 1 for doc in needed)
    map_doc = {doc: candidates[doc][0] for doc in needed}
    assert all(not (133274 <= int(a['key'].split('/')[1]) <= 135273) for a in map_doc.values())
    summary = json.loads((AUDIT / 'summary.json').read_text(encoding='utf-8'))
    target_keys = {a['key'] for a in map_doc.values()}
    sources, headers_by_key = summary['source_manifest'], {}
    for entry in sources:
        source = ROOT / entry['file']
        assert sha(source) == entry['sha256']
        labels = {}
        for rn, r in rows(source):
            if rn == 5:
                labels = {v: k for k, v in r.items() if v}
            if rn < 6 or amount(r.get('A')) is None:
                continue
            get = lambda name: r.get(labels.get(name))
            key = '/'.join(invoice_key(get('Ký hiệu'), get('Số hóa đơn')))
            if key in target_keys:
                assert key not in headers_by_key and get('Trạng thái gửi CQT') == 'HĐ hợp lệ' and get('Mã của CQT')
                headers_by_key[key] = {'date': invoice_day(get('Ngày hóa đơn')), 'cqt': get('Mã của CQT'), 'source_file': entry['file'], 'source_row': rn}
    assert set(headers_by_key) == target_keys
    note = '- Đã ghép HĐ gốc đã phát hành; ngày điều chỉnh 01/11/2026. BTL00023, BTL00024 giữ đúng Giá ghi trả AMIS theo xác nhận 28/10/2026.'
    requests, audit, changes = [], [], []
    date_cells, reference_cells = [], []
    seen_seq, seen_doc = set(), set()
    date_value = (date(2026, 11, 1) - date(1899, 12, 30)).days
    for index, r in enumerate(body, 9):
        seq, doc = r[0], r[25]
        first = seq not in seen_seq
        if first:
            assert doc not in seen_doc
            seen_seq.add(seq)
            seen_doc.add(doc)
        else:
            assert body[index - 10][0] == seq and body[index - 10][25] == doc, 'Non-contiguous invoice group'
        a = map_doc[doc]
        h = headers_by_key[a['key']]
        assert h['date'] == date(2026, 11, 1) and str(r[24]) == a['order']
        if not first:
            date_cells.append({'values': [encode(None)]})
            reference_cells.append({'values': [encode(None) for _ in range(4)]})
            continue
        series, number = a['key'].split('/')
        refs = [series, number, (h['date'] - date(1899, 12, 30)).days, h['cqt']]
        for col, v in [(1, date_value), *[(9 + i, v) for i, v in enumerate(refs)]]:
            cell = original[index]['values'][col]
            assert not cell.get('dataValidation') and not cell.get('userEnteredValue'), ('Target no longer empty', index + 1, col)
            changes.append({'row': index + 1, 'column': col + 1, 'value': v})
        date_cells.append({'values': [encode(date_value)]})
        reference_cells.append({'values': [encode(v) for v in refs]})
        audit.append({'row': index + 1, 'sequence': seq, 'doc': doc, 'order': a['order'], 'invoice': a['key'], 'date': h['date'].isoformat(), 'cqt': h['cqt'], 'source_file': h['source_file'], 'source_row': h['source_row']})
    assert len(audit) == 163 and len(changes) == 815
    assert len(date_cells) == len(reference_cells) == 219
    requests.extend([
        {'updateCells': {'range': {'sheetId': SHEET_ID, 'startRowIndex': 9, 'endRowIndex': 228, 'startColumnIndex': 1, 'endColumnIndex': 2}, 'rows': date_cells, 'fields': 'userEnteredValue'}},
        {'updateCells': {'range': {'sheetId': SHEET_ID, 'startRowIndex': 9, 'endRowIndex': 228, 'startColumnIndex': 9, 'endColumnIndex': 13}, 'rows': reference_cells, 'fields': 'userEnteredValue'}},
    ])
    requests.append({'updateCells': {'range': {'sheetId': SHEET_ID, 'startRowIndex': 6, 'endRowIndex': 7, 'startColumnIndex': 0, 'endColumnIndex': 1}, 'rows': [{'values': [encode(note)]}], 'fields': 'userEnteredValue'}})
    changes.append({'row': 7, 'column': 1, 'value': note})
    for request in requests:
        assert len(request) == 1
        update = request['updateCells']
        region = update['range']
        assert len(update['rows']) == region['endRowIndex'] - region['startRowIndex']
        assert all(len(r['values']) == region['endColumnIndex'] - region['startColumnIndex'] for r in update['rows'])
    prepared = {'spreadsheet_id': snapshot['spreadsheetId'], 'sheet_id': SHEET_ID, 'sheet': sheet['properties']['title'], 'rows': 219, 'invoices': 163,
                'adjustment_date': '2026-11-01', 'money': [-225865926, -18069275, -243935201], 'audit': audit, 'changes': changes,
                'source_manifest': sources, 'before_sha256': sha(WORK / 'before-cells.json'), 'requests': requests}
    (WORK / 'prepared.json').write_text(json.dumps(prepared, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'invoices': 163, 'lines': 219, 'money': prepared['money'], 'changed_cells': len(changes), 'requests': len(requests), 'all_original_dates': '2026-11-01'}))


if __name__ == '__main__':
    main()
