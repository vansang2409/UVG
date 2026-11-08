"""Verify this twelve-document correction against the delivered draft before it."""
from pathlib import Path
from decimal import Decimal
from openpyxl import load_workbook
import json,csv,hashlib

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/import-2025-20261109'
BEFORE=OUT/'archive/before-h1-12-link-review-20261109'
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def decisions(base):
    with (base/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig',newline='') as f:return {r['Chứng từ AMIS']:r for r in csv.DictReader(f)}
def main():
    old,new=read(BEFORE/'authoring.json'),read(OUT/'authoring.json')
    review=read(OUT/'h1-12-invoice-link-review.json');wanted={r['doc'] for r in review['records']}
    assert len(wanted)==12 and all(r['confirmed_invoice'] is None for r in review['records'])
    od,nd=decisions(BEFORE),decisions(OUT)
    assert od.keys()==nd.keys() and {c for c in nd if od[c]['Phân loại']!=nd[c]['Phân loại']}==wanted
    assert all(od[c]['Phân loại']=='Chờ xác minh' and nd[c]['Phân loại']=='Tạo import mới' for c in wanted)
    a=next(o for o in old['outputs'] if o['kind']=='new');b=next(o for o in new['outputs'] if o['kind']=='new')
    assert a['headers']==b['headers'] and b['rows'][:len(a['rows'])]==a['rows']
    added=b['rows'][len(a['rows']):]
    assert len(added)==13 and {r[18] for r in added}==wanted
    assert sorted({r[0] for r in added})==list(range(30213,30225))
    assert sum(Decimal(str(r[16] or 0))+Decimal(str(r[11] or 0)) for r in added)==11699000
    assert new['zero_audit']==old['zero_audit']
    for kind in ['adjustment','replacement','return_adjustment']:
        a=next(o for o in old['outputs'] if o['kind']==kind);b=next(o for o in new['outputs'] if o['kind']==kind)
        assert a==b
        assert (BEFORE/a['file']).read_bytes()==(OUT/'deliverables'/b['file']).read_bytes()
    prepared=read(OUT/'prepared.json');current={d['doc']:d for d in prepared['decisions']}
    for r in review['records']:
        assert current[r['doc']]['roots']==[]
        assert r['snapshot_invoice'] in prepared['docs'][r['doc']]['unverified_snapshot_roots']
    scan=read(OUT/'h1-12-invoice-source-audit.json')
    assert len(scan['controls'])==591 and not scan['hits'] and not scan['unreadable_files']
    assert all(hashlib.sha256((ROOT/r['file']).read_bytes()).hexdigest()==r['sha256'] for r in scan['controls'])
    report=next(o for o in new['outputs'] if o['kind']=='report')
    table=next(s for s in report['sheets'] if s['name']=='Ra soat 12 don T6')
    w=load_workbook(OUT/'deliverables'/report['file'],read_only=True,data_only=True)
    actual=list(w[table['name']].iter_rows(min_row=1,max_row=13,max_col=11,values_only=True))
    assert actual[0]==tuple(table['headers'])
    assert actual[1:]==[tuple(r) for r in table['rows']]
    w.close()
    result={'passed':True,'added_docs':12,'added_rows':13,'added_pay':11699000,'existing_rows_unchanged':34981,'other_three_imports_binary_unchanged':True,'actual_invoice_unresolved_docs':sorted(wanted),'sales_sequences':{r[18]:r[0] for r in added},'source_files_checked':591,'actions':new['summary']['actions'],'h1_new':sum(r['action']=='new' and r['date']<'2025-07-01' for r in current.values()),'h1_pending':sum(r['action']=='review' and r['date']<'2025-07-01' for r in current.values())}
    (OUT/'h1-12-import-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
