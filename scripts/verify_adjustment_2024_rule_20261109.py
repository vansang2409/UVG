"""Re-read original HĐ fields and verify the exported 2024-style adjustment workflow."""
from pathlib import Path
from collections import defaultdict
from decimal import Decimal
import json,sys,hashlib
from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel
from prepare_invoice_reports import rows,invoice_day
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109';BASE=OUT/'archive/before-adjustment-rule-2024-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def n(v):return Decimal(str(v or 0))
def blank(v):return None if v is None or v=='' else v
def main():
    a=read(OUT/'authoring.json');d=read(OUT/'adjustment-rule-2024-audit.json');plan=a['adjustment_rule_2024'];o=next(o for o in a['outputs'] if o['kind']=='adjustment');old=read(BASE/'authoring.json')
    wanted=defaultdict(set)
    for h in d['headers'].values():wanted[h['source']['file']].add(h['source']['row'])
    for ls in d['detail'].values():
        for l in ls:wanted[l['source']['file']].add(l['source']['row'])
    live={}
    for file,ns in wanted.items():
        for rn,r in rows(ROOT/file):
            if rn in ns:live[(file,rn)]=r
    for k,h in d['headers'].items():assert live[(h['source']['file'],h['source']['row'])]==h['values']
    for ls in d['detail'].values():
        for l in ls:assert live[(l['source']['file'],l['source']['row'])]==l['values']
    w=load_workbook(OUT/'deliverables'/o['file'],data_only=True,read_only=True);s=w.active;h=list(next(s.iter_rows(min_row=9,max_row=9,max_col=29,values_only=True)));assert h==o['headers']
    exported=[[to_excel(v) if hasattr(v,'date') else v for v in rr] for rr in s.iter_rows(min_row=10,max_row=42,max_col=29,values_only=True)]
    assert len(exported)==33 and all([blank(v) for v in x]==[blank(v) for v in y] for x,y in zip(exported,o['rows']))
    grouped=defaultdict(list)
    for r in exported:grouped[r[0]].append(r)
    assert sorted(grouped)==list(range(1,26))
    byroot=defaultdict(lambda:[Decimal(0)]*3);roles={};source_reversed=set()
    for seq,rs in grouped.items():
        metas=[m for m in plan['rows'] if m['seq']==seq];root=metas[0]['root'];source=metas[0]['source_invoice'];raw=d['headers'][root]['values'];first=rs[0]
        assert len({m['role'] for m in metas})==1
        for row in rs:
            assert blank(row[2])==blank(raw.get('H')) and blank(row[5])==blank(raw.get('L'))
            assert row[9:11]==root.split('/') and int(row[11])==to_excel(invoice_day(raw['D']))
            assert blank(row[12])==blank(raw.get('E')) and blank(row[3])==blank(raw.get('I')) and blank(row[4])==blank(raw.get('J'))
        assert sum(n(r[18]) for r in rs)==n(first[19]) and n(first[19])+n(first[21])==n(first[22])
        assert all(n(first[i])+n(first[j])==0 for i,j in [(19,26),(21,27),(22,28)])
        assert all(all(r[i] is None for i in [19,20,21,22,26,27,28]) for r in rs[1:])
        if metas[0]['role']=='Đảo khoản điều chỉnh cũ':
            assert len(rs)==1 and first[16] is None and first[17] is None and first[23] is None and source in first[13]
            source_reversed.add(source);m=d['headers'][source]['values'];assert [n(first[i]) for i in [19,21,22]]==[-n(m.get(c)) for c in ['P','Q','R']]
            assert source!=root and source in next(c['keys'] for c in d['chains'] if c['root']==root)
        else:
            ls=d['detail'][root];assert len(rs)==len(ls)
            for row,l in zip(rs,ls):
                rawline=live[(l['source']['file'],l['source']['row'])];mt=l['mtt'];assert not l['latest']
                assert row[14]==rawline['S' if mt else 'L'] and blank(row[15])==blank(rawline.get('T' if mt else 'M'))
                assert n(row[16])==-n(rawline['U' if mt else 'N']) and n(row[16])<0
                assert n(row[17])==n(rawline['V' if mt else 'O']) and n(row[18])==-n(rawline['W' if mt else 'P'])
                assert row[23]==(1 if n(row[18])==0 else None)
            m=d['headers'][root]['values'];assert [n(first[i]) for i in [19,21,22]]==[-n(m.get(c)) for c in ['P','Q','R']]
        roles[seq]=metas[0]['role']
        for i,j in enumerate([19,21,22]):byroot[root][i]+=n(first[j])
    assert len(source_reversed)==7 and len(plan['skipped_zero_adjustments'])==13
    for skip in plan['skipped_zero_adjustments']:
        assert skip['invoice'] not in source_reversed and all(n(v)==0 for v in skip['money'])
    assert len(byroot)==18
    for c in d['chains']:assert all(n(v)+z==0 for v,z in zip(c['before'],byroot[c['root']]))
    previous=next(o for o in old['outputs'] if o['kind']=='adjustment');oldtotal=[sum(n(r[j]) for r in previous['rows']) for j in [19,21,22]];newtotal=[sum(v[i] for v in byroot.values()) for i in range(3)];assert newtotal==oldtotal
    for kind in ['new','replacement','return_adjustment']:
        x=next(o for o in a['outputs'] if o['kind']==kind);y=next(o for o in old['outputs'] if o['kind']==kind);assert x==y and (BASE/x['file']).read_bytes()==(OUT/'deliverables'/x['file']).read_bytes()
    w.close();result={'passed':True,'chains':18,'invoices':25,'rows':33,'reverse_prior_nonzero':7,'reverse_original_detail':18,'zero_prior_verified_not_reissued':13,'original_names_exact':True,'all_references_f0':True,'all_three_balances_zero':True,'totals_unchanged':list(map(float,newtotal)),'other_three_imports_binary_unchanged':True,'source_price_preserved_flags':len(plan['source_price_flags']),'live_source_rows_checked':len(live)}
    (OUT/'adjustment-rule-2024-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
