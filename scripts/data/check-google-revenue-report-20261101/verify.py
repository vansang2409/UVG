"""Check the live read-only revenue tab against previously verified issue evidence."""
from pathlib import Path
from collections import Counter
import json,sys
sys.stdout.reconfigure(encoding='utf-8')

ROOT=Path(__file__).resolve().parents[3]
WORK=Path(__file__).resolve().parent
def load(p):return json.loads(p.read_text(encoding='utf-8'))
live=load(WORK/'live.json')
old=load(ROOT/'scripts/data/report-post-issue-20261030/prepared.json')
final=load(ROOT/'scripts/data/post-final-detail-20261101/summary.json')
verified=load(ROOT/'scripts/data/post-final-detail-20261101/verification.json')
assert verified['passed'] and final['issues']==[]
rows=live['data']['values'];assert len(rows)==len(old['evidence'])==778
assert len({r[0] for r in rows})==778
for row,e in zip(rows,old['evidence']):
    assert row[0]=='/'.join(e['root'])
    assert [row[c] for c in [6,9,10,11,13]]==[e['before_goods'],e['adjustment_goods'],e['replacement_reduction'],e['after_goods'],e['status']]
    assert row[6]+row[9]+row[10]==row[11]
    assert row[12]==0 and (len(row)<=14 or row[14] in [0,None,''])
top=live['top']['sheets'][0]['data'][0]['rowData']
def b(row):return top[row-1]['values'][1]['effectiveValue']['numberValue']
assert [b(r) for r in range(4,14)]==[47402157888,-46407850480,-371344277,622963131,49533443555,-225865926,49930540760,49928171800,2368960,2368960]
for rn,c in [(4,6),(5,9),(6,10)]:assert b(rn)==sum(row[c] for row in rows)
assert b(7)==sum(b(i) for i in [4,5,6])==sum(row[11] for row in rows)
assert b(8)==final['base_sales_money'][0]
assert b(9)==final['adjustments']['money'][0]
assert b(8)+b(9)==final['net_new_sales_goods']==49307577629
assert b(10)==b(7)+final['net_new_sales_goods']==49930540760
assert b(12)==b(10)-b(11)==2368960
assert final['replacements']['checks_passed']==2000
assert final['replacements']['latest_cqt_statuses']=={'HĐ hợp lệ':2000}
assert final['adjustments']['latest_cqt_statuses']=={'HĐ hợp lệ':163}
d=live['difference']['values']
sixteen=d[14:30]
assert len(sixteen)==16
assert all(r[3]-r[4]==r[5] for r in sixteen)
assert sum(r[5] for r in sixteen)==d[4][1]==-1519929
assert d[9][5]==d[3][1]==3888889
assert d[3][1]+d[4][1]==b(12)==d[5][1]==d[31][5]
result={'passed':True,'read_only':True,'tab':'Ra soat chuoi 2024','sheet_id':700000010,
        'historical_roots_checked':778,'historical_status_counts':dict(Counter(r[13] for r in rows)),
        'old_retained_goods':b(7),'new_sales_goods_after_returns':final['net_new_sales_goods'],
        'combined_report_goods':b(10),'amis_goods':b(11),'gap':b(12),
        'gap_components':{'invoice_111_different_year':3888889,'sixteen_groups':-1519929},
        'money_matches_verified_issued_scope':True,'not_fully_matched_to_amis':True,
        'outdated_text':['A1:A2 update date / evidence cutoff','A8:C10 still forecast / not yet checked'],
        'google_sheet_edited':False}
(WORK/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False,indent=2))
