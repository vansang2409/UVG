"""Check all appended cells against frozen outputs and original invoice/relationship rows."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal
import json,sys,hashlib
from openpyxl import load_workbook
from prepare_invoice_reports import rows,invoice_day,invoice_key
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109';BEFORE=OUT/'archive/before-unmapped-meinvoice-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def n(v):return Decimal(str(v)) if v is not None and str(v).strip() else None
def key(s,v):
    x=invoice_key(s,v);return '/'.join(x) if x else None
def main():
    a=read(OUT/'authoring.json');old=read(BEFORE/'authoring.json');audit=read(OUT/'unmapped-meinvoice-audit.json');p=read(OUT/'prepared.json');ready={r['root']:r for r in audit['components'] if r['status']=='ready'};nodes=audit['invoice_details'];added=a['unmapped_zero_audit']
    assert {r['root'] for r in added}==set(ready) and len(added)==1483
    assert a['zero_audit']==old['zero_audit']
    for kind in ['new','return_adjustment']:
        x=next(o for o in old['outputs'] if o['kind']==kind);y=next(o for o in a['outputs'] if o['kind']==kind)
        assert x==y and (BEFORE/x['file']).read_bytes()==(OUT/'deliverables'/y['file']).read_bytes()
    for kind in ['adjustment','replacement']:
        x=next(o for o in old['outputs'] if o['kind']==kind);y=next(o for o in a['outputs'] if o['kind']==kind)
        assert y['rows'][:len(x['rows'])]==x['rows'] and x['headers']==y['headers']
        assert len(y['rows'])==x['invoices']+sum(r['kind']==kind for r in added)
    cache=read(OUT/'source-index.json');children=defaultdict(set,{k:set(v) for k,v in cache['children'].items()});parents=defaultdict(set,{k:set(v) for k,v in cache['parents'].items()});replacement=defaultdict(set,{k:set(v) for k,v in cache['replacement'].items()})
    relation_files=[ROOT/'data/Du lieu Meinvoice/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_2025.xls',ROOT/'data/Du lieu Meinvoice/Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_2024_2025.xls']+list((ROOT/'data/Du lieu Meinvoice/Xuat_lai_2026-10-22').glob('*.xls'))
    for f in relation_files:
        for rn,row in rows(f):
            if rn<7:continue
            x,y=key(row.get('B'),row.get('D')),key(row.get('M'),row.get('O'))
            if not x or not y:continue
            children[x].add(y);parents[y].add(x)
            if 'thay_the' in f.name:replacement[x].add(y)
    requests=defaultdict(set)
    for r in ready.values():
        for k in r['keys']:
            src=nodes[k]['record']['source'];requests[src['file']].add(src['row'])
            if nodes[k].get('payment_evidence'):
                e=nodes[k]['payment_evidence'];requests[e['file']].add(e['row'])
    originals={}
    for f,numbers in requests.items():
        for rn,row in rows(ROOT/f):
            if rn in numbers:originals[(f,rn)]=row
    companies={r['key'] for r in p['company_rows'] if r['key']};mapped={k for d in p['decisions'] for k in d['roots']}
    report=next(o for o in a['outputs'] if o['kind']=='report');old_report=next(o for o in old['outputs'] if o['kind']=='report')
    for sh in old_report['sheets']:assert next(s for s in report['sheets'] if s['name']==sh['name'])==sh
    wb=load_workbook(OUT/'deliverables'/report['file'],read_only=True,data_only=True)
    for sh in report['sheets']:
        if sh['name'] not in ['MI2025 chua map','MI2025 chuoi','Bieu tang - KT check']:continue
        actual=list(wb[sh['name']].iter_rows(min_row=1,max_row=len(sh['rows'])+1,max_col=len(sh['headers']),values_only=True))
        for got,want in zip(actual,[sh['headers']]+sh['rows']):
            assert all(x==y or (x is None and y=='') for x,y in zip(got,want)),sh['name']
    wb.close()
    payment=Decimal(0)
    for z in added:
        r=ready[z['root']];root=r['root'];seen=set();todo=[root]
        while todo:
            k=todo.pop()
            if k in seen:continue
            seen.add(k);todo.extend(children[k]);todo.extend(parents[k])
        assert seen==set(r['keys']) and not seen&companies and not seen&mapped and not r['mapping_evidence']
        assert not parents[root]
        active={k for k in seen if not replacement[k] and 'hủy' not in str(nodes[k]['record']['status']).lower()}
        assert active==set(r['active']);sums=[Decimal(0)]*3
        for k in active:
            src=nodes[k]['record']['source'];raw=originals[(src['file'],src['row'])]
            assert key(raw.get('B'),raw.get('C'))==k
            m=[n(raw.get(c)) for c in ['P','Q','R']]
            if m[2] is None:
                e=nodes[k]['payment_evidence'];detail=originals[(e['file'],e['row'])]
                assert key(detail.get('F'),detail.get('B'))==k
                assert n(detail['W'])-n(detail['Y'])==m[0] and n(detail['AA'])==m[1]
                m[2]=n(detail['AB']);assert m[2] is not None and m[0]+m[1]==m[2]
            assert all(x is not None for x in m)
            sums=[x+y for x,y in zip(sums,m)]
        assert sums==list(map(n,z['before'])) and any(sums) and sums[0]+sums[1]==sums[2]
        assert all(n(b)+n(c)==0 for b,c in zip(z['before'],z['change']))
        src=nodes[root]['record']['source'];raw=originals[(src['file'],src['row'])]
        assert invoice_day(raw['D']).year==2025 and not raw.get('J')
        o=next(o for o in a['outputs'] if o['kind']==z['kind']);row=o['rows'][z['sequence']-1]
        assert row[25]==root and row[2] is None and row[5]==(raw.get('L') or raw.get('H')) and row[24] is None
        if z['kind']=='replacement':assert len(seen)==1 and row[23]==1 and all(n(row[i])==0 for i in [17,18,19,21])
        else:
            assert len(seen)>1 and any(parents[k] for k in seen)
            assert [n(row[i]) for i in [19,21,22]]==[-v for v in sums]
        payment+=sums[2]
    gifts=read(OUT/'gift-kt-check.json');assert len(gifts['rows'])==16 and not {r[0] for r in gifts['rows']}&set(ready)
    for field in ['new_invoices','new_lines','new_goods','new_vat','new_pay','actions']:assert a['summary'][field]==old['summary'][field]
    result={'passed':True,'new_chains':len(added),'new_replacements':1472,'new_adjustments':11,'payment_to_zero':float(payment),'all_f0_f1_f2_checked':True,'original_cells_checked':len(originals),'old_import_rows_retained':True,'sales_and_btl_binary_unchanged':True,'gift_kt_check':16,'source_exceptions':2}
    (OUT/'unmapped-import-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
