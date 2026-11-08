"""Independently verify source totals, exported cells, exclusions and zeroing."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal,ROUND_HALF_UP
from openpyxl import load_workbook
import csv,json,hashlib,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def n(x):return Decimal(str(x or 0))
def main():
    p=json.loads((OUT/'authoring.json').read_text(encoding='utf-8'))
    prepared=json.loads((OUT/'prepared.json').read_text(encoding='utf-8'))
    docs=prepared['docs'];seen=set();sourcecontrols=[];exports=[];newmap={}
    service_path=OUT/'service-quantity-confirmations.json'
    service_corrections={r['doc']:r for r in json.loads(service_path.read_text(encoding='utf-8'))['records']} if service_path.exists() else {}
    rounding_path=OUT/'rounding-confirmation.json'
    rounding_limit=n(json.loads(rounding_path.read_text(encoding='utf-8'))['max_abs_delta_vnd']) if rounding_path.exists() else Decimal(1)
    assert Decimal(1)<=rounding_limit<=Decimal(2)
    for m in p['manifest']:
        assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256'],m
        sourcecontrols.append(m)
    with (OUT/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig',newline='') as f:
        decisions={r['Chứng từ AMIS']:r for r in csv.DictReader(f)}
    assert len(decisions)==len(docs)==65493
    assert Counter(r['Phân loại'] for r in decisions.values())==Counter({'Tạo import mới':p['summary']['new_invoices'],'Giữ nguyên, không import':p['summary']['actions']['keep'],'Loại theo HD CTY 2025':p['summary']['actions']['company'],'Chờ xác minh':p['summary']['actions']['review']})
    for o in p['outputs']:
        path=OUT/'deliverables'/o['file'];wb=load_workbook(path,read_only=True,data_only=False)
        if o['kind']=='report':
            assert len(wb.sheetnames)==len(o['sheets'])
            for sh in o['sheets']:
                ws=wb[sh['name']];assert sum(1 for _ in ws.iter_rows())>=len(sh['rows'])+(sh.get('start') or 1)
            wb.close();exports.append({'file':o['file'],'sha256':hashlib.sha256(path.read_bytes()).hexdigest()});continue
        ws=wb.active;actual=list(ws.iter_rows(min_row=10,max_row=9+len(o['rows']),max_col=len(o['headers']),values_only=True))
        assert len(actual)==len(o['rows'])
        h=list(next(ws.iter_rows(min_row=9,max_row=9,max_col=len(o['headers']),values_only=True)));assert h==o['headers']
        for i,(row,expected) in enumerate(zip(actual,o['rows']),10):
            for j,(a,b) in enumerate(zip(row,expected)):
                if hasattr(a,'date') and isinstance(b,int):
                    from datetime import datetime,timedelta
                    assert a==datetime(1899,12,30)+timedelta(days=b),(o['file'],i,j,a,b)
                elif isinstance(a,(int,float)) and isinstance(b,(int,float)):
                    assert abs(n(a)-n(b))<Decimal('0.000001'),(o['file'],i,j,a,b)
                else:assert a==b or (a is None and b==''),(o['file'],i,j,a,b)
        if o['kind']=='new':
            grouped=defaultdict(list)
            for r in actual:grouped[r[18]].append(r)
            assert len(grouped)==o['invoices']
            for code,ls in grouped.items():
                assert code not in seen,code;seen.add(code)
                assert decisions[code]['Phân loại']=='Tạo import mới',code
                d=docs[code]
                assert len(ls)==len(d['lines'])
                assert sum(n(r[16]) for r in ls)==n(d['goods']),code
                assert sum(n(r[11]) for r in ls)==n(d['vat']),code
                assert sum(n(r[16])+n(r[11]) for r in ls)==n(d['pay']),code
                assert len({r[0] for r in ls})==1
                newmap[code]=(o['file'],ls[0][0])
                for r,l in zip(ls,d['lines']):
                    sr=l['v'];assert r[19]==str(sr[15]).strip(),code
                    expected_quantity=n(sr[22])
                    if code in service_corrections:
                        c=service_corrections[code]
                        assert l['source']==c['source'] and str(sr[15]).strip()==c['sku'] and n(sr[22])==n(c['source_quantity'])
                        assert docs[code]['business_type']=='Dịch vụ' and str(sr[19]).strip()==c['unit']
                        expected_quantity=n(c['import_quantity'])
                    assert n(r[14])==expected_quantity and n(r[15])==n(sr[23]) and n(r[16])==n(sr[24]),code
                    assert abs((n(r[14])*n(r[15])).quantize(Decimal(1),rounding=ROUND_HALF_UP)-n(r[16]))<=rounding_limit,code
        elif o['kind']=='replacement':
            assert len({r[25] for r in actual})==len(actual)
            for r in actual:assert all(n(r[i])==0 for i in [17,18,19,21]) and r[23]==1
        elif o['kind']=='adjustment':
            if p.get('adjustment_rule_2024'):
                grouped=defaultdict(list)
                for r in actual:grouped[r[0]].append(r)
                for rs in grouped.values():
                    first=rs[0];assert sum(n(r[18]) for r in rs)==n(first[19])
                    assert n(first[19])+n(first[21])==n(first[22])
                    assert all(n(first[a])+n(first[b])==0 for a,b in [(19,26),(21,27),(22,28)])
                    assert all(all(r[i] is None for i in [19,20,21,22,26,27,28]) for r in rs[1:])
                check=json.loads((OUT/'adjustment-rule-2024-verification.json').read_text(encoding='utf-8'));assert check['passed']
            else:
                for r in actual:
                    assert n(r[18])==n(r[19]) and n(r[19])+n(r[21])==n(r[22])
                    assert all(n(r[a])+n(r[b])==0 for a,b in [(19,26),(21,27),(22,28)])
        elif o['kind']=='return_adjustment':
            assert len(actual)==410 and len({r[24] for r in actual})==321
            assert all(all(r[i] is None for i in [9,10,11,12]) for r in actual)
            for r in actual:
                assert decisions[r[25]]['Phân loại']!='Loại theo HD CTY 2025'
                assert n(r[16])<0 and n(r[17])>=0 and n(r[18])<=0
                if r[19] is not None:assert n(r[19])+n(r[21])==n(r[22])
        assert sorted({r[0] for r in actual})==list(range(1,o['invoices']+1))
        wb.close();exports.append({'file':o['file'],'invoices':o['invoices'],'rows':len(actual),'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
    assert seen=={c for c,r in decisions.items() if r['Phân loại']=='Tạo import mới'}
    companykeys={x['key'] for x in prepared['company_rows'] if x['key']}
    assert not companykeys & {z['root'] for z in p['zero_audit']}
    assert len({z['root'] for z in p['zero_audit']})==len(p['zero_audit'])
    for z in p['zero_audit']:
        assert all(n(a)+n(b)==0 for a,b in zip(z['before'],z['change']))
        assert all(decisions[c]['Phân loại']=='Tạo import mới' for c in z['docs'])
        if z['kind']=='replacement':assert len(prepared['chains'][z['root']]['keys'])==1
    ret=json.loads((OUT/'return-verification.json').read_text(encoding='utf-8'))
    assert ret['source_ui_docs']==322 and ret['source_ui_lines']==412
    assert len(ret['checks'])==322
    assert sum(n(r['pay']) for r in ret['checks'])==Decimal('479741870')
    report=next(o for o in p['outputs'] if o['kind']=='report')
    returns=next(s for s in report['sheets'] if s['name']=='Tra lai cho so HD')
    w=load_workbook(OUT/'deliverables'/report['file'],read_only=True,data_only=True)
    reference_count=0
    for sh in report['sheets']:
        h=sh['headers']
        if 'File import bán mới' not in h:continue
        fi=h.index('File import bán mới');si=h.index('STT HĐ bán mới');ci=h.index('Chứng từ bán gốc')
        start=sh.get('start') or 1
        for row in w[sh['name']].iter_rows(min_row=start+1,max_row=start+len(sh['rows']),max_col=len(h),values_only=True):
            if not row[fi]:continue
            assert (row[fi],row[si])==newmap[row[ci]],(sh['name'],row[ci],row[fi],row[si])
            reference_count+=1
    rr=list(w['Tra lai cho so HD'].iter_rows(min_row=10,max_row=9+len(returns['rows']),max_col=31,values_only=True))
    assert len(rr)==410
    assert all(all(r[i] is None for i in [9,10,11,12]) for r in rr)
    assert len({r[24] for r in rr})==321
    for r,expected in zip(rr,returns['rows']):
        assert r[24:28]==tuple(expected[24:28])
        for i in [16,17,18,19,20,21,22]:assert n(r[i])==n(expected[i])
    w.close()
    standalone=next(o for o in p['outputs'] if o['kind']=='return_adjustment')
    w=load_workbook(OUT/'deliverables'/standalone['file'],read_only=True,data_only=True)
    standalone_rows=list(w.active.iter_rows(min_row=10,max_row=419,max_col=31,values_only=True))
    assert standalone_rows==rr,'Standalone BTL file differs from report draft'
    w.close()
    with (OUT/'LIEN_KET_IMPORT_MOI.csv').open(encoding='utf-8-sig',newline='') as f:
        mappings=list(csv.DictReader(f))
    assert len(mappings)==len(newmap)
    assert {r['Chứng từ AMIS'] for r in mappings}==set(newmap)
    for r in mappings:assert (r['File'],int(r['STT HĐ']))==newmap[r['Chứng từ AMIS']]
    with (OUT/'LIEN_KET_BTL_BHH_2025.csv').open(encoding='utf-8-sig',newline='') as f:
        for r in csv.DictReader(f):
            if r['File import bán mới']:assert (r['File import bán mới'],int(r['STT HĐ bán mới']))==newmap[r['Chứng từ bán gốc']]
    result={'passed':True,'source_hashes':sourcecontrols,'exported_files':exports,'new_docs_unique':len(seen),'zero_chains':len(p['zero_audit']),'no_company_import':True,'all_import_cells_roundtrip':True,'new_invoice_references_checked':reference_count,'return_links':322,'return_source_lines':412,'return_draft_docs':321,'return_draft_lines':410,'return_amount_exceptions':p['summary']['returns']['issues'],'limitations':p['summary']['limitations']}
    if p.get('unmapped_zero_audit'):
        orphan=json.loads((OUT/'unmapped-import-verification.json').read_text(encoding='utf-8'));assert orphan['passed']
        result['matched_amis_zero_chains']=len(p['zero_audit']);result['unmapped_zero_chains']=len(p['unmapped_zero_audit'])
        result['zero_chains']+=len(p['unmapped_zero_audit'])
        result['historical_unmapped_source_verification']=orphan
        result['historical_unmapped_source_verification_note']='Source-chain audit before adjustment rows were rebuilt under 2024 workflow; old row-retention checks describe that earlier snapshot only.'
    if p.get('adjustment_rule_2024'):
        current=json.loads((OUT/'adjustment-rule-2024-verification.json').read_text(encoding='utf-8'));assert current['passed']
        result['adjustment_rule_2024_verification']=current
    (OUT/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'passed':True,'files':len(exports),'new_invoices':len(seen),'zero_chains':result['zero_chains']}))
if __name__=='__main__':main()
