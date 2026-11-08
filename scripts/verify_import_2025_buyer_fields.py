"""Compare the full buyer correction against the frozen pre-correction payload."""
from pathlib import Path
from collections import Counter
import json,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def main():
    before=read(OUT/'archive/before-buyer-field-fix-20261109/authoring.json');after=read(OUT/'authoring.json')
    legal={'CÔNG TY TNHH SẢN XUẤT VÀ ĐẦU TƯ AKIA','CÔNG TY TNHH ĐẦU TƯ Y TẾ TRẦN DIỆP KHANH'}
    assert before['summary']==after['summary'] and before['zero_audit']==after['zero_audit']
    stats={}
    for a,b in zip(before['outputs'],after['outputs']):
        assert {k:v for k,v in a.items() if k not in ['rows','sheets']}=={k:v for k,v in b.items() if k not in ['rows','sheets']}
        tables=[(a,b)] if a['kind']!='report' else list(zip(a['sheets'],b['sheets']))
        for x,y in tables:
            name=x.get('name') or x['kind']
            if a['kind']=='report' and name!='Tra lai cho so HD':assert x==y;continue
            assert x['headers']==y['headers'] and len(x['rows'])==len(y['rows'])
            count=0
            for old,new in zip(x['rows'],y['rows']):
                expected=old.copy()
                if old[2] and old[2] not in legal:
                    expected[5]=old[5] or old[2];expected[2]=None;count+=1
                assert new==expected,(name,old[0])
            stats[name]={'changed_rows':count,'company_rows':sum(bool(r[2]) for r in y['rows'])}
    assert stats['new']=={'changed_rows':35010,'company_rows':6}
    assert stats['adjustment']['changed_rows']==7 and stats['replacement']['changed_rows']==567
    assert stats['return_adjustment']['changed_rows']==stats['Tra lai cho so HD']['changed_rows']==410
    sales=next(o for o in after['outputs'] if o['kind']=='new')
    assert {r[2] for r in sales['rows'] if r[2]}==legal
    assert all(r[5] for r in sales['rows'] if r[2] is None)
    result={'passed':True,'tables':stats,'all_other_cells_unchanged':True,'financial_totals_unchanged':True,'existing_person_names_retained':True}
    (OUT/'buyer-field-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
