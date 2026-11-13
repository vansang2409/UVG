from pathlib import Path
from zipfile import ZipFile
from openpyxl import load_workbook
from decimal import Decimal
from collections import defaultdict
import io,json,sys,hashlib
sys.stdout.reconfigure(encoding='utf-8')
R=Path('D:/UVG');O=Path(__file__).parent
sys.path.insert(0,str(R/'scripts'));from prepare_invoice_reports import rows,invoice_key
e=json.loads((O/'five-invoice-deep-evidence.json').read_text(encoding='utf-8'));keys={k for c in e['cases'] for k in [c['orphan'],c['kept']]}
matches=[];relationship_rows=0
for p in (R/'data/1311/meinvoice').glob('*.zip'):
    if not any(x in p.name for x in ['dieu_chinh','thay_the']):continue
    with ZipFile(p) as z:
        for member in z.namelist():
            if member.endswith('/'):continue
            for rn,r in rows(io.BytesIO(z.read(member))):
                if rn<7 or not str(r.get('A') or '').isdigit():continue
                a,b=invoice_key(r.get('B'),r.get('D')),invoice_key(r.get('M'),r.get('O'));a='/'.join(a) if a else None;b='/'.join(b) if b else None;relationship_rows+=1
                if a in keys or b in keys:matches.append({'file':member,'row':rn,'old':a,'new':b})
assert not matches
print('CHECKED relationships',relationship_rows,flush=True)
targets={}
for c in e['cases']:
    for k,ls in c['invoice_lines'].items():
        for l in ls:targets[l['row']]=(k,l)
def n(v):return None if v in (None,'') else Decimal(str(v))
p=next((R/'data/1311/meinvoice').glob('*chi_tiet*.zip'))
with ZipFile(p) as z:
    names=[x for x in z.namelist() if not x.endswith('/')];assert len(names)==1
    w=load_workbook(io.BytesIO(z.read(names[0])),read_only=True,data_only=True);checked=0
    for rn,r in enumerate(w.active.iter_rows(min_row=min(targets),max_row=max(targets),max_col=19,values_only=True),min(targets)):
        if rn not in targets:continue
        k,l=targets[rn];assert str(r[1]).upper()+'/'+str(r[2]).removesuffix('.0').zfill(8)==k
        assert r[11]==l['name'] and str(r[12] or '')==str(l['unit'] or '')
        for col,field in [(13,'qty'),(14,'price'),(15,'goods'),(17,'vat'),(18,'payment')]:assert n(r[col])==n(l[field]),(rn,field)
        checked+=1
    w.close()
assert checked==20
for s in e['sources']:assert hashlib.sha256((R/s['file']).read_bytes()).hexdigest()==s['sha256']
full=json.loads((O/'reverse-chain-audit.json').read_text(encoding='utf-8'))
assert not [r for r in full['recovered_edges'] if r['old'] in keys or r['new'] in keys]
result={'passed':True,'invoice_keys_checked':len(keys),'independent_detail_rows_checked':checked,'current_relationship_rows_scanned':relationship_rows,'related_adjustment_replacement_rows':matches,'sources_unchanged':True}
(O/'five-deep-source-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False),flush=True)
