"""Read-only comparison of 08/11 AMIS exports with the linked AMIS tab."""
from pathlib import Path
from collections import Counter, defaultdict
from datetime import datetime, date, timedelta
from decimal import Decimal
import csv
import hashlib
import json
import sys
import unicodedata
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding='utf-8')
OUT = Path(__file__).resolve().parent
SOURCE = Path('D:/UVG/data/0811')
FIELDS = ['Ngày hạch toán', 'Số chứng từ', 'Mã khách hàng', 'Khách hàng', 'Diễn giải', 'Tiền thuế GTGT', 'Tổng tiền thanh toán']

def text(x):
    return '' if x is None else unicodedata.normalize('NFC', str(x)).strip()

def day(x):
    if isinstance(x, (datetime, date)):
        return x.isoformat()[:10]
    if isinstance(x, (float, int)):
        return (datetime(1899, 12, 30) + timedelta(days=x)).date().isoformat()
    return text(x)[:10]

def number(x):
    return Decimal(str(x)) if isinstance(x, (int, float)) else None

def native(x):
    if isinstance(x, Decimal):
        return int(x) if x == x.to_integral_value() else float(x)
    if isinstance(x, (datetime, date)):
        return x.isoformat()
    if isinstance(x, set):
        return sorted(x)
    raise TypeError(type(x).__name__)

def save(name, value):
    (OUT / name).write_text(json.dumps(value, ensure_ascii=False, indent=2, default=native) + '\n', encoding='utf-8')

def csv_save(name, headers, records):
    with (OUT / name).open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for rec in records:
            writer.writerow([native(v) if isinstance(v, Decimal) else v for v in rec])

live = {}
for path in sorted((OUT / 'live').glob('amis_*.json'), key=lambda p: int(p.stem.split('_')[-1])):
    chunk = json.loads(path.read_text(encoding='utf-8'))
    start = int(path.stem.split('_')[-1])
    for rownum, raw in enumerate(chunk['values'], start):
        r = raw + [None] * max(0, 13 - len(raw))
        if not r[2]:
            continue
        doc = text(r[2])
        assert doc not in live, ('duplicate live doc', doc)
        live[doc] = {'row': rownum, 'values': [day(r[0]), doc, text(r[3]), text(r[4]), text(r[5]), number(r[6]), number(r[7])], 'order': r[8], 'invoice': r[9], 'symbol': r[10], 'mi_money': r[11], 'delta': r[12]}
assert len(live) == 35646

new = {}
new_dups = []
source_controls = []
wb = load_workbook(SOURCE / 'Ban_hang.xlsx', read_only=True, data_only=True)
for ws in wb:
    assert list(next(ws.iter_rows(min_row=3, max_row=3, values_only=True)))[1:8] == FIELDS
    count = 0
    footer = []
    for rownum, raw in enumerate(ws.iter_rows(min_row=4, values_only=True), 4):
        r = list(raw)
        if not isinstance(r[1], (datetime, date)) or not r[2]:
            if any(v is not None for v in r):
                footer.append([rownum, r])
            continue
        doc = text(r[2])
        rec = {'sheet': ws.title, 'row': rownum, 'values': [day(r[1]), doc, text(r[3]), text(r[4]), text(r[5]), number(r[6]), number(r[7])], 'invoice_status': r[8], 'payment_status': r[9], 'stock_status': r[10]}
        if doc in new:
            new_dups.append([doc, new[doc], rec])
        else:
            new[doc] = rec
        count += 1
    source_controls.append({'sheet': ws.title, 'docs': count, 'footer': footer})
wb.close()
assert not new_dups, new_dups[:5]
print(json.dumps({'stage': 'sales_read', 'docs': len(new), 'source_controls': source_controls}, ensure_ascii=False, default=native), flush=True)

new_h2 = {k: v for k, v in new.items() if '2025-07-01' <= v['values'][0] <= '2025-12-31'}
common = sorted(live.keys() & new.keys())
missing = sorted(live.keys() - new.keys())
added = sorted(new_h2.keys() - live.keys())
changes = []
changed_docs = set()
field_counts = Counter()
for doc in common:
    old, fresh = live[doc], new[doc]
    for idx, field in enumerate(FIELDS):
        a, b = old['values'][idx], fresh['values'][idx]
        if isinstance(a, Decimal) and isinstance(b, Decimal):
            equal = abs(a - b) < Decimal('0.000001')
        else:
            equal = a == b
        if not equal:
            changed_docs.add(doc)
            field_counts[field] += 1
            changes.append({'doc': doc, 'live_row': old['row'], 'new_sheet': fresh['sheet'], 'new_row': fresh['row'], 'field': field, 'old': a, 'new': b, 'difference': b-a if isinstance(a, Decimal) and isinstance(b, Decimal) else None})

monthly = {}
for month in range(7, 13):
    prefix = f'2025-{month:02d}'
    old_records = [r for r in live.values() if r['values'][0].startswith(prefix)]
    new_records = [r for r in new_h2.values() if r['values'][0].startswith(prefix)]
    monthly[str(month)] = {'live_docs': len(old_records), 'new_docs': len(new_records), 'changed_docs': sum(1 for d in changed_docs if live[d]['values'][0].startswith(prefix)), 'only_new': sum(1 for d in added if new[d]['values'][0].startswith(prefix)), 'not_in_new': sum(1 for d in missing if live[d]['values'][0].startswith(prefix))}
    for key, idx in [('vat', 5), ('money', 6)]:
        monthly[str(month)]['live_' + key] = sum((r['values'][idx] or Decimal(0) for r in old_records), Decimal(0))
        monthly[str(month)]['new_' + key] = sum((r['values'][idx] or Decimal(0) for r in new_records), Decimal(0))
        monthly[str(month)]['change_' + key] = monthly[str(month)]['new_' + key] - monthly[str(month)]['live_' + key]

impact = Counter()
impact_monthly = defaultdict(Counter)
impact_details = []
for doc in common:
    old = live[doc]
    fresh = new[doc]
    old_delta = number(old['delta'])
    mi = number(old['mi_money'])
    if mi is None and old['mi_money'] in (None, ''):
        mi = Decimal(0)  # Same arithmetic as =H-L in the existing sheet.
    if old_delta is None or mi is None:
        status = 'unavailable_invoice_mapping'
        new_delta = None
    else:
        new_delta = fresh['values'][6] - mi
        was_bad = abs(old_delta) > Decimal('0.000001')
        now_bad = abs(new_delta) > Decimal('0.000001')
        status = ('still_mismatched' if now_bad else 'resolved') if was_bad else ('new_mismatch' if now_bad else 'still_matched')
    impact[status] += 1
    impact_monthly[old['values'][0][5:7]][status] += 1
    if status in ('resolved', 'new_mismatch') or doc in changed_docs:
        impact_details.append({'doc': doc, 'live_row': old['row'], 'status': status, 'old_money': old['values'][6], 'new_money': fresh['values'][6], 'mi_money': old['mi_money'], 'old_delta': old['delta'], 'new_delta': new_delta})

summary = {'live_docs': len(live), 'new_full_year_docs': len(new), 'new_h2_docs': len(new_h2), 'common_docs': len(common), 'unchanged_common_docs': len(common)-len(changed_docs), 'changed_docs': len(changed_docs), 'changed_fields': dict(field_counts), 'change_cells': len(changes), 'not_in_new_export': missing, 'only_in_new_h2': added, 'monthly': monthly, 'impact_holding_existing_meinvoice_mapping': dict(impact), 'impact_monthly': dict(impact_monthly), 'source_controls': source_controls}
save('comparison-summary.json', summary)
save('changed-fields.json', changes)
save('mismatch-impact.json', impact_details)
save('membership-differences.json', {'not_in_new': {d: live[d] for d in missing}, 'only_in_new_h2': {d: new[d] for d in added}})
csv_save('chi-tiet-thay-doi.csv', ['Số chứng từ', 'Dòng trên link', 'Tab file mới', 'Dòng file mới', 'Trường thay đổi', 'Giá trị trên link', 'Giá trị file mới', 'Chênh lệch mới - cũ'], [[r[k] for k in ['doc', 'live_row', 'new_sheet', 'new_row', 'field', 'old', 'new', 'difference']] for r in changes])
print(json.dumps({'stage': 'comparison', **summary, 'change_samples': changes[:20]}, ensure_ascii=False, default=native), flush=True)

# Aggregate detailed sales to independently check the new summary export.
detail = defaultdict(lambda: {'rows': 0, 'money': Decimal(0), 'vat': Decimal(0), 'dates': set(), 'invoices': set(), 'sheets': set(), 'first_row': None})
detail_controls = []
wb = load_workbook(SOURCE / 'So_chi_tiet_ban_hang.xlsx', read_only=True, data_only=True)
for ws in wb:
    headers = list(next(ws.iter_rows(min_row=4, max_row=4, values_only=True)))
    assert headers[2] == 'Số chứng từ' and headers[41] == 'Thuế GTGT' and headers[43] == 'Tổng thanh toán'
    count = 0
    footer = []
    for rownum, raw in enumerate(ws.iter_rows(min_row=5, values_only=True), 5):
        r = list(raw)
        if not isinstance(r[0], (datetime, date)) or not r[2]:
            if any(v is not None for v in r):
                footer.append([rownum, r])
            continue
        doc = text(r[2])
        x = detail[doc]
        x['rows'] += 1
        x['money'] += number(r[43]) or Decimal(0)
        x['vat'] += number(r[41]) or Decimal(0)
        x['dates'].add(day(r[0]))
        x['invoices'].add(text(r[3]) + '/' + text(r[5]))
        x['sheets'].add(ws.title)
        if x['first_row'] is None:
            x['first_row'] = rownum
        count += 1
    detail_controls.append({'sheet': ws.title, 'lines': count, 'footer': footer})
    print(json.dumps({'stage': 'detail_sheet_read', 'sheet': ws.title, 'lines': count}, ensure_ascii=False), flush=True)
wb.close()
detail_mismatches = []
for doc in sorted(new.keys() & detail.keys()):
    s = new[doc]
    d = detail[doc]
    if abs(s['values'][5] - d['vat']) > Decimal('0.000001') or abs(s['values'][6] - d['money']) > Decimal('0.000001') or d['dates'] != {s['values'][0]}:
        detail_mismatches.append({'doc': doc, 'summary_vat': s['values'][5], 'detail_vat': d['vat'], 'summary_money': s['values'][6], 'detail_money': d['money'], 'summary_date': s['values'][0], 'detail_dates': sorted(d['dates'])})
extra_detail = sorted(detail.keys() - new.keys())
check = {'detail_docs': len(detail), 'detail_lines': sum(v['rows'] for v in detail.values()), 'summary_docs_missing_detail': sorted(new.keys() - detail.keys()), 'detail_docs_not_in_summary': extra_detail, 'summary_detail_mismatch_count': len(detail_mismatches), 'summary_detail_mismatches': detail_mismatches, 'detail_controls': detail_controls, 'detail_not_in_summary_prefixes': dict(Counter(d[:3] for d in extra_detail))}
save('detail-verification.json', check)
save('detail-aggregates.json', dict(detail))
save('manifest.json', [{'path': str(p), 'bytes': p.stat().st_size, 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()} for p in SOURCE.glob('*.xlsx')])
print(json.dumps({'stage': 'detail_verification', **check}, ensure_ascii=False, default=native), flush=True)
