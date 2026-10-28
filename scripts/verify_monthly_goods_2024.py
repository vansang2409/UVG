"""Kiểm bảng tháng/mã hàng trực tiếp với sổ AMIS và tổng phương án đã kiểm."""
from pathlib import Path
from collections import defaultdict
from decimal import Decimal
import json,hashlib
from prepare_invoice_reports import rows,invoice_day,amount
ROOT=Path(__file__).resolve().parents[1]
P=ROOT/'scripts/data/monthly-goods-20261028'
d=json.loads((P/'prepared.json').read_text(encoding='utf8'))
f=json.loads((ROOT/'scripts/data/revenue-forecast-20261028/prepared.json').read_text(encoding='utf8'))
z=lambda x:amount(x) or Decimal(0)
raw=defaultdict(lambda:[Decimal(0)]*3);count=0
src=ROOT/'data/Du lieu Amis/So_chi_tiet_ban_hang_2024.xlsx'
for rn,r in rows(src):
    a=invoice_day(r.get('A'));day=invoice_day(r.get('B'))
    if rn<5 or not a or a.year!=2024 or not r.get('C'):continue
    assert day and day.year==2024
    key=(day.month,r.get('M'),r.get('Q'));v=raw[key]
    for i,x in enumerate([z(r.get('R'))-z(r.get('AA')),z(r.get('U'))-z(r.get('AB')),z(r.get('AN'))]):v[i]+=x
    count+=1
assert count==44578
stats={(r['month'],r['code'],r['unit']):r for r in d['stats']}
for k,money in raw.items():
    r=stats[k]
    assert money==[z(r['qty_amis']),z(r['rev_amis']),z(r['cost_amis'])],(k,money,r)
for k,r in stats.items():
    if k not in raw:assert [r['qty_amis'],r['rev_amis'],r['cost_amis']]==[0,0,0]
assert sum(z(r['rev_invoice']) for r in d['stats'])==f['totals']['forecast_after_manual'][0]==49930540760
assert sum(z(r['cost_invoice']) for r in d['stats'])==24542992457
assert sum(z(r['revenue']) for r in d['pending'] if r['kind']=='Phân bổ tiền')==4739490
assert len(d['cost_pending'])==2
assert {r['invoice'] for r in d['cost_pending']}=={'1C24TUV/00000063','1C24TUV/00000111'}
for m in d['monthly']:
    rs=[r for r in d['stats'] if r['month']==m['month']]
    for a,b in [('rev_amis','rev_amis'),('rev_invoice','rev_invoice'),('cost_amis','cost_amis'),('cost_invoice_known','cost_invoice')]:assert z(m[a])==sum(z(r[b]) for r in rs)
for source in d['sources']:assert hashlib.sha256((ROOT/source['path']).read_bytes()).hexdigest()==source['sha256']
result={'status':'passed','raw_amis_lines':count,'month_sku_rows':len(stats),'forecast_revenue':49930540760,'amis_revenue':49928171800,'known_cost_only':24542992457,'missing_cost_invoices':['1C24TUV/00000063','1C24TUV/00000111'],'unmapped_item_rows':d['counts']['unmapped_item_rows'],'unallocated_money_rows':d['counts']['unallocated_money_rows'],'unallocated_money':4739490,'source_hashes_verified':len(d['sources'])}
(P/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(result,ensure_ascii=False))
