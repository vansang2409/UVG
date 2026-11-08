"""Check exactly eight confirmed free promotions and preserve the previous import."""
from pathlib import Path
import json,csv

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/import-2025-20261109'
BEFORE=OUT/'archive/before-promotion-fix-20261109'
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def decisions(base):
    with (base/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig',newline='') as f:return {r['Chứng từ AMIS']:r for r in csv.DictReader(f)}
def main():
    old,new=read(BEFORE/'authoring.json'),read(OUT/'authoring.json')
    confirmed=read(OUT/'promotion-confirmations.json');wanted={r['doc'] for r in confirmed['records']}
    assert len(wanted)==8 and confirmed['amis_live_checked']
    od,nd=decisions(BEFORE),decisions(OUT)
    assert od.keys()==nd.keys() and {c for c in nd if od[c]['Phân loại']!=nd[c]['Phân loại']}==wanted
    assert all(od[c]['Phân loại']=='Chờ xác minh' and nd[c]['Phân loại']=='Tạo import mới' for c in wanted)
    a=next(o for o in old['outputs'] if o['kind']=='new');b=next(o for o in new['outputs'] if o['kind']=='new')
    assert b['headers']==a['headers']+['Hàng khuyến mại']
    assert [r[:22] for r in b['rows'][:len(a['rows'])]]==a['rows']
    assert all(r[22] is None for r in b['rows'][:len(a['rows'])])
    added=b['rows'][len(a['rows']):]
    assert len(added)==8 and {r[18] for r in added}==wanted
    assert sorted({r[0] for r in added})==list(range(30225,30233))
    assert all(r[10]==8 and r[11]==0 and r[14]==1 and r[15]==r[16]==0 and r[22]==1 for r in added)
    assert {r[18] for r in b['rows'] if r[22]==1}==wanted
    for key in ['new_goods','new_vat','new_pay']:assert old['summary'][key]==new['summary'][key]
    assert new['zero_audit']==old['zero_audit']
    for kind in ['adjustment','replacement','return_adjustment']:
        a=next(o for o in old['outputs'] if o['kind']==kind);b=next(o for o in new['outputs'] if o['kind']==kind)
        assert a==b and (BEFORE/a['file']).read_bytes()==(OUT/'deliverables'/b['file']).read_bytes()
    p=read(OUT/'prepared.json')
    for c in wanted:
        assert p['docs'][c]['lines'][0]['v'][18]=='✓'
    result={'passed':True,'promotion_docs':sorted(wanted),'promotion_rows':8,'rate_read_on_amis':8,'promotion_marker_column':'W - Hàng khuyến mại','marker':1,'old_34994_rows_A_to_V_unchanged':True,'all_monetary_totals_unchanged':True,'other_three_imports_binary_unchanged':True,'sequences':{r[18]:r[0] for r in added},'new_invoices':new['summary']['new_invoices'],'new_rows':new['summary']['new_lines'],'h1_new':sum(r['action']=='new' and r['date']<'2025-07-01' for r in p['decisions']),'h1_pending':sum(r['action']=='review' and r['date']<'2025-07-01' for r in p['decisions'])}
    (OUT/'promotion-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
