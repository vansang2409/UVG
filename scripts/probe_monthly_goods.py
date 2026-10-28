"""Khảo sát mã hàng, số lượng, giá vốn và độ phủ chi tiết hóa đơn giữ lại."""
from pathlib import Path
from collections import Counter,defaultdict
import json
from decimal import Decimal
from prepare_invoice_reports import rows,invoice_key,amount
ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/monthly-goods-20261028'
WORK.mkdir(parents=True,exist_ok=True)
f=json.loads((ROOT/'scripts/data/revenue-forecast-20261028/prepared.json').read_text(encoding='utf8'))
p=json.loads((ROOT/'scripts/data/amis-2024-linked-20261026/prepared.json').read_text(encoding='utf8'))
gap=json.loads((ROOT/'scripts/data/revenue-gap-20261028/prepared.json').read_text(encoding='utf8'))
manual={tuple(r['root']) for r in f['manual_records']}
kept={tuple(r['root']):r for r in f['retained_roots'] if tuple(r['root']) not in manual}
needed={tuple(r['key']):r for r in f['old_invoices'] if tuple(r['root']) in kept}
details=defaultdict(list)
for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025.xls']:
    src=ROOT/'data/Du lieu Meinvoice'/name
    for rn,r in rows(src):
        k=invoice_key(r.get('D'),r.get('B'))
        if rn<7 or k not in needed:continue
        details[k].append({'row':rn,'source':str(src.relative_to(ROOT)),'code':r.get('K'),'name':r.get('L'),'unit':r.get('M'),'qty':r.get('N'),'price':r.get('O'),'goods':r.get('P'),'discount':r.get('R'),'vat':r.get('T'),'payment':r.get('U')})
native=lambda x:int(x) if isinstance(x,Decimal) and x==int(x) else float(x) if isinstance(x,Decimal) else x
out={'kept_roots':list(kept.values()),'needed':list(needed.values()),'details':{'/'.join(k):v for k,v in details.items()},'groups':gap['groups']}
(WORK/'probe.json').write_text(json.dumps(out,ensure_ascii=False,default=native),encoding='utf8')
mismatches=[]
for k,h in needed.items():
    s=sum((amount(r['goods']) or 0)-(amount(r['discount']) or 0) for r in details[k])
    if s!=h['money'][0]:mismatches.append({'key':k,'header':h['money'][0],'detail':native(s),'lines':len(details[k])})
docs={d['doc']:d for d in p['docs']}
print(json.dumps({'roots':len(kept),'invoices':len(needed),'detail_lines':sum(map(len,details.values())),'missing_invoices':['/'.join(k) for k in needed if not details[k]],'lines_without_code':sum(not r['code'] for rs in details.values() for r in rs),'mismatches':mismatches,'sale_sample':docs['BHHS2024000055']['original_lines'][0]['values'],'return_sample':docs['BTL00024']['original_lines'][0]['values'],'kept_samples':list(out['details'].items())[:2]},ensure_ascii=False,default=native))
