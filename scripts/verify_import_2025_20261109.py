"""Independently verify source totals, exported cells, exclusions and zeroing."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal,ROUND_HALF_UP
from openpyxl import load_workbook
import csv,json,hashlib,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
if len(sys.argv)>1:
    OUT=Path(sys.argv[1]).resolve()
sys.stdout.reconfigure(encoding='utf-8')
def n(x):return Decimal(str(x or 0))
def main():
    p=json.loads((OUT/'authoring.json').read_text(encoding='utf-8'))
    prepared=json.loads((OUT/'prepared.json').read_text(encoding='utf-8'))
    docs=prepared['docs'];seen=set();sourcecontrols=[];exports=[];newmap={}
    normalized_scripts=[]
    exchanges=p['summary'].get('psc2_exchanges',{})
    reissues=p['summary'].get('rounding_reissue_20261109',{})
    effective_prices={r['source']:r for r in reissues.get('effective_prices',[])}
    replacement_overrides={r['root']:r for r in reissues.get('records',[])}
    continuation=p['summary'].get('continued_reissue_20261109',{}).get('records',[])
    replacement_overrides.update({r['root']:r for r in continuation})
    confirmed_reissue=p['summary'].get('confirmed_reissue_7502')
    if confirmed_reissue:
        replacement_overrides[confirmed_reissue['root']]=confirmed_reissue
    final_reissue=p['summary'].get('final_pending_reissues_20261109')
    confirmed_keep=p['summary'].get('confirmed_keep_12515_20261109')
    confirmed_duplicate=p['summary'].get('confirmed_duplicate_12392')
    replacement_rows={}
    service_path=OUT/'service-quantity-confirmations.json'
    service_corrections={r['doc']:r for r in json.loads(service_path.read_text(encoding='utf-8'))['records']} if service_path.exists() else {}
    rounding_path=OUT/'rounding-confirmation.json'
    rounding_limit=n(json.loads(rounding_path.read_text(encoding='utf-8'))['max_abs_delta_vnd']) if rounding_path.exists() else Decimal(1)
    assert Decimal(1)<=rounding_limit<=Decimal(2)
    for m in p['manifest']:
        path=ROOT/m['file']
        source=path.read_bytes()
        digest=hashlib.sha256(source).hexdigest()
        if digest!=m['sha256'] and path.suffix in {'.py','.mjs'}:
            lf=source.replace(b'\r\n',b'\n')
            digests={hashlib.sha256(v).hexdigest() for v in (lf,lf.replace(b'\n',b'\r\n'))}
            assert m['sha256'] in digests,m
            normalized_scripts.append(m['file'])
        else:
            assert digest==m['sha256'],m
        sourcecontrols.append(m)
    with (OUT/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig',newline='') as f:
        decisions={r['Chứng từ AMIS']:r for r in csv.DictReader(f)}
    assert len(decisions)==len(docs)==65493
    gift_update=p['summary'].get('gift_keeps_and_amis_deletion_20261109',{})
    deleted=set(gift_update.get('deleted_amis',[]))
    assert {code for code,row in decisions.items() if row['Phân loại']=='Đã xóa trên AMIS'}==deleted
    counts=Counter(row['Phân loại'] for code,row in decisions.items() if code not in deleted)
    assert counts==Counter({'Tạo import mới':p['summary']['new_invoices'],'Giữ nguyên, không import':p['summary']['actions']['keep'],'Loại theo HD CTY 2025':p['summary']['actions']['company'],'Chờ xác minh':p['summary']['actions']['review']})
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
                source_lines=d['lines']
                excluded_goods=Decimal(0)
                expected_vat=n(d['vat'])
                assert n(d['goods'])+expected_vat==n(d['pay']),code
                if code in exchanges:
                    exchange=exchanges[code]
                    excluded=[line for line in source_lines if line['v'][15]=='PSC2']
                    assert len(excluded)==1 and excluded[0]['source']==exchange['psc2_source'],code
                    assert n(excluded[0]['v'][41])==0 and n(excluded[0]['v'][43])==n(exchange['excluded_goods']),code
                    excluded_goods=n(excluded[0]['v'][24])
                    assert excluded_goods==n(exchange['excluded_goods']),code
                    source_lines=[line for line in source_lines if line not in excluded]
                    expected_sources=(exchange['product_sources'] if 'product_sources' in exchange else [exchange['product_source']])
                    assert [line['source'] for line in source_lines]==expected_sources,code
                    goods=sum(n(line['v'][24]) for line in source_lines)
                    source_vat=sum(n(line['v'][41]) for line in source_lines)
                    assert source_vat==n(d['vat']),code
                    if exchange.get('rate_policy')=='recalculate':
                        assert exchange['rate']==8 and n(exchange['source_vat'])==source_vat,code
                        expected_vat=(goods*Decimal('0.08')).quantize(Decimal(1),rounding=ROUND_HALF_UP)
                    else:
                        expected_vat=source_vat
                    assert n(exchange['vat'])==expected_vat and all(r[10]==exchange['rate'] for r in ls),code
                    assert goods==n(exchange['goods']) and goods+expected_vat==n(exchange['payment']),code
                assert len(ls)==len(source_lines)
                assert sum(n(r[16]) for r in ls)==n(d['goods'])-excluded_goods,code
                assert sum(n(r[11]) for r in ls)==expected_vat,code
                assert sum(n(r[16])+n(r[11]) for r in ls)==n(d['goods'])-excluded_goods+expected_vat,code
                assert len({r[0] for r in ls})==1
                newmap[code]=(o['file'],ls[0][0])
                for r,l in zip(ls,source_lines):
                    sr=l['v'];assert r[19]==str(sr[15]).strip(),code
                    expected_quantity=n(sr[22])
                    if code in service_corrections:
                        c=service_corrections[code]
                        assert l['source']==c['source'] and str(sr[15]).strip()==c['sku'] and n(sr[22])==n(c['source_quantity'])
                        assert docs[code]['business_type']=='Dịch vụ' and str(sr[19]).strip()==c['unit']
                        expected_quantity=n(c['import_quantity'])
                    expected_price=n(sr[23])
                    if l['source'] in effective_prices:
                        price=effective_prices[l['source']]
                        assert price['doc']==code and n(price['source_unit_price'])==expected_price
                        expected_price=n(price['import_unit_price'])
                        assert expected_quantity*expected_price==n(sr[24]),code
                    assert n(r[14])==expected_quantity and n(r[15])==expected_price and n(r[16])==n(sr[24]),code
                    assert abs((n(r[14])*n(r[15])).quantize(Decimal(1),rounding=ROUND_HALF_UP)-n(r[16]))<=rounding_limit,code
        elif o['kind']=='replacement':
            assert len({r[25] for r in actual})==len(actual)
            for r in actual:assert all(n(r[i])==0 for i in [17,18,19,21]) and r[23]==1
            replacement_rows.update({r[25]:r for r in actual})
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
                if final_reissue:
                    plan=p['adjustment_rule_2024']
                    assert len(plan['chains'])==p['summary']['adjustment_chains']==19
                    assert len(actual)==plan['line_count']==35 and len(grouped)==plan['invoices']==27
                    reissue=next(r for r in final_reissue['records'] if r['kind']=='adjustment')
                    reversal,reduction=[grouped[seq][0] for seq in reissue['zero_seqs']]
                    f0=reissue['root'].split('/')
                    assert all(list(row[9:11])==f0 and row[24:26]==(reissue['code'],reissue['root'])
                               for row in (reversal,reduction))
                    assert reversal[2:6]==reduction[2:6]
                    assert [reversal[i]+reduction[i] for i in (19,21,22)]==[-v for v in reissue['before']]
                    assert reversal[16:18]==(None,None) and reduction[16:18]==(-1,600411)
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
    if final_reissue:
        pending=final_reissue['pending']
        assert pending['candidate'] not in replacement_rows
        assert all(pending['candidate'] not in z.get('keys',[z['root']])
                   for z in p['zero_audit']+p['unmapped_zero_audit'])
        assert pending['doc'] not in seen
        if confirmed_keep:
            row=decisions[pending['doc']]
            assert row['Phân loại']=='Giữ nguyên, không import' and confirmed_keep['doc']==pending['doc']
            assert row['Chuỗi liên quan']==confirmed_keep['invoice']==pending['candidate']
            assert confirmed_keep['detached_invoice']==pending['sheet_link']
            detached=confirmed_keep['detached_invoice']
            if confirmed_duplicate and confirmed_duplicate['root']==detached:
                assert confirmed_duplicate['doc']!=confirmed_keep['doc']
            else:
                assert detached not in replacement_rows
                assert not any(detached in z.get('keys',[z['root']])
                               for z in p['zero_audit']+p['unmapped_zero_audit'])
        else:
            assert decisions[pending['doc']]['Phân loại']=='Chờ xác minh'
        for record in final_reissue['records']:
            assert newmap[record['code']][1]==record['new_seq']
            if record['kind']=='replacement':
                row=replacement_rows[record['root']]
                assert row[0]==record['zero_seqs'][0] and row[24]==record['code']
                assert '/'.join(row[11:13])==record['root'] and row[14]
    for z in p['zero_audit']:
        assert all(n(a)+n(b)==0 for a,b in zip(z['before'],z['change']))
        if confirmed_duplicate and z['root']==confirmed_duplicate['root']:
            doc=confirmed_duplicate['doc']
            decision=decisions[doc]
            assert decision['Phân loại']=='Giữ nguyên, không import' and doc not in seen
            assert decision['Chuỗi liên quan']==confirmed_duplicate['kept_invoice']
            assert confirmed_duplicate['kept_invoice'] not in replacement_rows
            assert z['kind']=='replacement' and z['keys']==[z['root']] and z['docs']==[doc]
            assert z['before']==confirmed_duplicate['money']
            row=replacement_rows[z['root']]
            assert row[0]==z['sequence']==confirmed_duplicate['zero_seq']
            assert '/'.join(row[11:13])==z['root'] and row[14] and row[24]==doc
            continue
        assert all(decisions[c]['Phân loại']=='Tạo import mới' for c in z['docs'])
        if z['kind']=='replacement':
            root=z['root']
            if root not in replacement_overrides:
                assert len(prepared['chains'][root]['keys'])==1
                continue
            override=replacement_overrides[root]
            row=replacement_rows[root]
            assert row[0]==z['sequence']==override['zero_seq']
            assert row[11]+'/'+row[12]==override['active'] and row[14]
            if 'new_sequences' in override:
                assert row[24]==', '.join(override['docs'])
                assert {code:newmap[code][1] for code in override['docs']}==override['new_sequences']
            else:
                assert row[24]==override['doc'] and newmap[row[24]][1]==override['new_seq']
            assert [n(x) for x in z['before']]==[n(x) for x in override['before']]
            chain=prepared['chains'].get(root)
            if chain:
                assert chain['active']==[override['active']] and not chain['issues']
                assert chain['keys']==override['keys'] and chain['balance']==z['before']
            else:
                assert override['keys']==[root] and override['active']==root
    ret=json.loads((OUT/'return-verification.json').read_text(encoding='utf-8'))
    assert ret['source_ui_docs']==322 and ret['source_ui_lines']==412
    assert len(ret['checks'])==322
    assert sum(n(r['pay']) for r in ret['checks'])==Decimal('479741870')
    report=next(o for o in p['outputs'] if o['kind']=='report')
    returns=next(s for s in report['sheets'] if s['name']=='Tra lai cho so HD')
    w=load_workbook(OUT/'deliverables'/report['file'],read_only=True,data_only=True)
    reporting=p['summary'].get('amis_reporting_0911')
    if reporting:
        scope,projection,gap=list(w['Tong hop'].iter_rows(min_row=34,max_row=36,min_col=6,max_col=7,values_only=True))
        assert [list(r) for r in (scope,projection,gap)]==[reporting['scope_current'],reporting['projection'],reporting['gap']]
        source_note=next(w['Tong hop'].iter_rows(min_row=17,max_row=17,min_col=2,max_col=2,values_only=True))[0]
        assert f"Báo cáo lấy AMIS 2025 xuất {reporting['source_date']}" in source_note
        for code,entry in reporting['psc2_documents'].items():
            row=decisions[code]
            assert [int(row[k]) for k in ('Tiền hàng AMIS','VAT AMIS','Thanh toán AMIS')]==entry['money']
            assert entry['draft'][0]==entry['money'][0]==exchanges[code]['goods']
        assert sum(e['draft'][1]-e['money'][1] for e in reporting['psc2_documents'].values())==reporting['vat_policy_delta']
    if gift_update:
        gifts={entry['root']:entry['money'] for entry in gift_update['records']}
        kept={row[0]:row for row in w['MI2025 chua map'].iter_rows(min_row=2,values_only=True) if row[0] in gifts}
        assert kept.keys()==gifts.keys() and len(kept)==3
        assert all(row[1]=='Giữ HĐ biếu tặng KOL/KOC' and row[2] is None and row[10] is None
                   and list(row[4:7])==gifts[root] for root,row in kept.items())
        assert not any(gifts.keys() & set(plan.get('keys',[plan['root']]))
                       for plan in p['zero_audit']+p['unmapped_zero_audit'])
        assert [row[0] for row in w['Cho xac minh'].iter_rows(min_row=2,values_only=True) if row[0]]==[
            code for code,row in decisions.items() if row['Phân loại']=='Chờ xác minh']
        assert sum(p['summary']['actions'].values())==p['summary']['total']==len(decisions)-len(deleted)
    kept_gifts=p['summary'].get('unmapped_meinvoice',{}).get('gift_kt_kept',0)
    if kept_gifts:
        policy=json.loads((OUT/'unmapped-gift-confirmation.json').read_text(encoding='utf-8'))
        keys=set(policy['invoice_roots'])
        assert policy['accountant_decision']=='keep' and len(keys)==kept_gifts==16
        gift_rows=list(w['Bieu tang - KT check'].iter_rows(min_row=2,values_only=True))
        assert {row[0] for row in gift_rows}==keys and len(gift_rows)==kept_gifts
        assert all('Kế toán chốt giữ lại' in row[9] for row in gift_rows)
        mapped=[row for row in w['MI2025 chua map'].iter_rows(min_row=2,values_only=True) if row[0] in keys]
        assert len(mapped)==kept_gifts
        assert all('Kế toán chốt giữ lại' in row[1] and row[2] is None and row[10] is None for row in mapped)
        assert not any(keys.intersection(plan.get('keys',[plan['root']])) for plan in p['zero_audit']+p['unmapped_zero_audit'])
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
    if exchanges:
        assert exchanges.keys()<=seen
        result['psc2_exchanges']=exchanges
    if normalized_scripts:
        result['scripts_verified_with_normalized_line_endings']=normalized_scripts
    if kept_gifts:
        result['gift_invoices_kept']=kept_gifts
    (OUT/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'passed':True,'files':len(exports),'new_invoices':len(seen),'zero_chains':result['zero_chains']}))
if __name__=='__main__':main()
