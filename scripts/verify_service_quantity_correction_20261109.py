"""Verify two user-confirmed service quantities while retaining the original AMIS data."""
from pathlib import Path
from decimal import Decimal,ROUND_HALF_UP
import json,csv,hashlib

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/import-2025-20261109'
BEFORE=OUT/'archive/before-service-quantity-fix-20261109'
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def decisions(base):
    with (base/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig',newline='') as f:return {r['Chứng từ AMIS']:r for r in csv.DictReader(f)}
def main():
    old,new=read(BEFORE/'authoring.json'),read(OUT/'authoring.json')
    correction=read(OUT/'service-quantity-confirmations.json');wanted={r['doc'] for r in correction['records']}
    assert wanted=={'BHDV20250104','BHDV20250115'}
    rounding=read(OUT/'rounding-confirmation.json') if (OUT/'rounding-confirmation.json').exists() else {'newly_eligible_documents':[],'newly_eligible_rows':0,'goods':0,'vat':0,'pay':0}
    round_docs=set(rounding['newly_eligible_documents']);all_added=wanted|round_docs
    od,nd=decisions(BEFORE),decisions(OUT)
    assert od.keys()==nd.keys() and {c for c in nd if od[c]['Phân loại']!=nd[c]['Phân loại']}==all_added
    assert all(od[c]['Phân loại']=='Chờ xác minh' and nd[c]['Phân loại']=='Tạo import mới' for c in all_added)
    a=next(o for o in old['outputs'] if o['kind']=='new');b=next(o for o in new['outputs'] if o['kind']=='new')
    without_identity=lambda rows:[[v for i,v in enumerate(r) if i not in [2,5]] for r in rows]
    assert a['headers']==b['headers'] and without_identity(b['rows'][:len(a['rows'])])==without_identity(a['rows'])
    added=b['rows'][len(a['rows']):]
    assert len(added)==2+rounding['newly_eligible_rows'] and {r[18] for r in added}==all_added
    services=[r for r in added if r[18] in wanted];rounded=[r for r in added if r[18] in round_docs]
    assert len(services)==2 and sorted(r[0] for r in services)==[30233,30234]
    assert all(r[10]==8 and r[11]==22222 and r[13]=='Lần' and r[14]==2 and r[15]==138889 and r[16]==277778 and r[22] is None for r in services)
    assert all(r[14]*r[15]==r[16] and r[16]+r[11]==300000 for r in services)
    if round_docs:
        assert len(round_docs)==10 and len(rounded)==12
        assert sorted({r[0] for r in rounded})==list(range(30235,30245))
        assert all(abs((Decimal(str(r[14]))*Decimal(str(r[15]))).quantize(Decimal(1),rounding=ROUND_HALF_UP)-Decimal(str(r[16])))<=2 for r in rounded)
        assert sum(Decimal(str(r[16])) for r in rounded)==Decimal(str(rounding['goods']))
        assert sum(Decimal(str(r[11] or 0)) for r in rounded)==Decimal(str(rounding['vat']))
        assert nd['BHHT2025027490']['Phân loại']==nd['BHHS2025015861']['Phân loại']=='Chờ xác minh','3đ must not be accepted'
    for field,delta,key in [('new_goods',555556,'goods'),('new_vat',44444,'vat'),('new_pay',600000,'pay')]:assert Decimal(str(new['summary'][field]))-Decimal(str(old['summary'][field]))==delta+Decimal(str(rounding[key]))
    assert new['zero_audit']==old['zero_audit']
    for kind in ['adjustment','replacement','return_adjustment']:
        a=next(o for o in old['outputs'] if o['kind']==kind);b=next(o for o in new['outputs'] if o['kind']==kind)
        assert {k:v for k,v in a.items() if k!='rows'}=={k:v for k,v in b.items() if k!='rows'}
        assert without_identity(a['rows'])==without_identity(b['rows'])
    p=read(OUT/'prepared.json')
    for c in wanted:
        d=p['docs'][c];line=d['lines'][0]
        assert d['business_type']=='Dịch vụ' and line['v'][22]==1 and line['import_quantity']==2
        assert line['quantity_confirmation']['source_quantity']==1
    assert all(d['business_type']=='Dịch vụ' for c,d in p['docs'].items() if c.startswith('BHDV'))
    v=read(BEFORE/'verification.json')
    sources={r['file'].replace('\\','/'):r['sha256'] for r in v['source_hashes']}
    for path in ['data/0811/Ban_hang.xlsx','data/0811/So_chi_tiet_ban_hang.xlsx']:assert hashlib.sha256((ROOT/path).read_bytes()).hexdigest()==sources[path]
    result={'passed':True,'documents':sorted(wanted),'classification':'Dịch vụ','source_quantity_retained':1,'confirmed_import_quantity':2,'unit':'Lần','price':138889,'goods_each':277778,'vat_each':22222,'payment_each':300000,'service_added_payment':600000,'rounding_documents':sorted(round_docs),'rounding_rows':len(rounded),'rounding_tolerance_vnd':2 if round_docs else None,'combined_added_docs':len(all_added),'combined_added_rows':len(added),'combined_added_payment':600000+rounding['pay'],'old_35002_rows_unchanged':True,'source_files_unchanged':True,'other_three_imports_binary_unchanged':True,'sequences':{r[18]:r[0] for r in added},'new_invoices':new['summary']['new_invoices'],'new_rows':new['summary']['new_lines'],'h1_new':sum(r['action']=='new' and r['date']<'2025-07-01' for r in p['decisions']),'h1_pending':sum(r['action']=='review' and r['date']<'2025-07-01' for r in p['decisions'])}
    result.pop('old_35002_rows_unchanged');result.pop('other_three_imports_binary_unchanged')
    result['old_35002_rows_unchanged_except_authorized_buyer_fields']=True
    result['other_three_imports_unchanged_except_authorized_buyer_fields']=True
    (OUT/'service-quantity-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
