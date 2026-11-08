from pathlib import Path
from collections import Counter, defaultdict
from datetime import datetime
import json
import sys
from decimal import Decimal

sys.stdout.reconfigure(encoding='utf-8')
OUT = Path(__file__).resolve().parent
source_path = OUT.parent / 'doi-chieu-amis-0811-20261108/detail-aggregates.json'
source = json.loads(source_path.read_text(encoding='utf-8'))
sheet = json.loads((OUT / 'hd-cty-before-values.json').read_text(encoding='utf-8'))['values']
records = []
groups = defaultdict(list)
for rn, row in enumerate(sheet, 1):
    if rn < 5 or len(row) < 4 or not row[3]:
        continue
    r = row + [None] * max(0, 23-len(row))
    doc = str(r[16] or '').strip()
    rec = {'row': rn, 'invoice': f'{r[1]}{r[2]}/{r[3]}', 'invoice_status': r[11], 'invoice_vat': r[9], 'invoice_money': r[10], 'doc': doc, 'order': r[12], 'buyer': r[6], 'date': r[4], 'noteR': r[17]}
    x = source.get(doc)
    if x is None:
        rec.update(found=False, amis_vat=None, amis_money=None, delta_vat=None, delta_money=None)
    else:
        dv = x['vat'] - r[9] if isinstance(r[9], (int, float)) else None
        dm = x['money'] - r[10] if isinstance(r[10], (int, float)) else None
        rec.update(found=True, amis_vat=x['vat'], amis_money=x['money'], delta_vat=dv, delta_money=dm)
    records.append(rec)
    groups[doc].append(rec)
counts = Counter()
for x in records:
    if not x['found']:
        status = 'Không tìm thấy CT trong AMIS 08/11'
    elif x['invoice_status'] == 'Hóa đơn đã bị thay thế':
        status = 'HĐ đã bị thay thế; kiểm tra chuỗi'
    elif len(groups[x['doc']]) > 1:
        status = f'Cùng CT có {len(groups[x["doc"]])} HĐ; kiểm tra theo chuỗi'
    elif x['delta_vat'] is None or x['delta_money'] is None:
        status = 'Thiếu số tiền trên HĐ để so sánh'
    elif x['delta_vat'] == 0 and x['delta_money'] == 0:
        status = 'Khớp tiền và VAT'
    elif abs(x['delta_vat']) <= 2 and abs(x['delta_money']) <= 2:
        status = 'Lệch 1–2 đ'
    else:
        status = 'Lệch tiền hoặc VAT'
    x['result'] = status
    counts[status] += 1

summary = {'invoice_rows': len(records), 'distinct_doc_cells': len(groups), 'mapped_rows': sum(r['found'] for r in records), 'not_found_rows': sum(not r['found'] for r in records), 'results': dict(counts), 'row_level_money_nonzero': sum(r['delta_money'] is not None and r['delta_money'] != 0 for r in records), 'row_level_vat_nonzero': sum(r['delta_vat'] is not None and r['delta_vat'] != 0 for r in records), 'either_nonzero': sum((r['delta_money'] or 0) != 0 or (r['delta_vat'] or 0) != 0 for r in records)}
(OUT / 'mapping-analysis.json').write_text(json.dumps({'summary': summary, 'records': records}, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'summary': summary, 'exceptions': [r for r in records if r['result'] not in ('Khớp tiền và VAT', 'Lệch 1–2 đ')]}, ensure_ascii=False), flush=True)
