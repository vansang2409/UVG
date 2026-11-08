from pathlib import Path
from collections import defaultdict
from datetime import datetime, date
from decimal import Decimal
import json
import sys
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding='utf-8')
OUT = Path(__file__).resolve().parent
cases = json.loads((OUT/'vat-only-live-candidates.json').read_text(encoding='utf-8'))
docs = {r['doc'] for r in cases}
totals = defaultdict(lambda: {'rows': 0, 'sales': Decimal(0), 'discount': Decimal(0), 'returns': Decimal(0), 'reduction': Decimal(0), 'export_tax': Decimal(0), 'vat': Decimal(0), 'money': Decimal(0)})
wb = load_workbook('D:/UVG/data/0811/So_chi_tiet_ban_hang.xlsx', read_only=True, data_only=True)
for ws in wb:
    for row in ws.iter_rows(min_row=5, values_only=True):
        doc = str(row[2] or '').strip()
        if doc not in docs or not isinstance(row[0], (datetime, date)):
            continue
        rec = totals[doc]
        rec['rows'] += 1
        for field, idx in [('sales',24),('discount',29),('returns',36),('reduction',37),('export_tax',39),('vat',41),('money',43)]:
            val = row[idx]
            assert val is None or isinstance(val, (int, float)), (doc, field, val)
            rec[field] += Decimal(str(val)) if val is not None else Decimal(0)
wb.close()
results = []
for case in cases:
    t = totals[case['doc']]
    assert t['rows'] > 0
    assert t['export_tax'] == 0
    before = t['sales']-t['discount']-t['returns']-t['reduction']
    assert t['vat'] == case['amisVat'] and t['money'] == case['amisMoney']
    assert before+t['vat'] == t['money'], (case['doc'], before, t)
    results.append({**case, 'amisBeforeTax': before, 'deltaBeforeTax': before-Decimal(str(case['invoiceBefore'])), 'detailRows': t['rows']})
def encode(x):
    if isinstance(x, Decimal):
        return int(x) if x == x.to_integral_value() else float(x)
    raise TypeError(type(x).__name__)
(OUT/'vat-only-verified.json').write_text(json.dumps(results,ensure_ascii=False,indent=2,default=encode),encoding='utf-8')
print(json.dumps({'cases':len(results),'results':results},ensure_ascii=False,default=encode),flush=True)
