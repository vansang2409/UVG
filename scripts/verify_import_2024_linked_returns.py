"""Đọc độc lập Excel cuối cùng và sổ nguồn; kiểm từng dòng, tổng và liên kết."""
from pathlib import Path
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from io import BytesIO
from zipfile import ZipFile
import hashlib
import json
import re
import os
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/amis-2024-linked-20261026'
OUT=Path(os.environ.get('UVG_OUTPUT_DIR',ROOT/'outputs/amis-2024-theo-btl-bhh-20261026'))
def d(v):return Decimal(str(v or 0))
def canon(v):return None if v in ('',None) else v
def add(target,key,values):
    target[key]=[a+d(b) for a,b in zip(target[key],values)]

def main():
    meta=json.loads((WORK/'prepared.json').read_text(encoding='utf8'))
    for m in meta['manifest']:assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256']
    src=load_workbook(ROOT/'data/Du lieu Amis/So_chi_tiet_ban_hang_2024.xlsx',read_only=True,data_only=True).active
    rows={i:list(r) for i,r in enumerate(src.iter_rows(min_row=5,values_only=True),5) if isinstance(r[0],datetime) and r[2]}
    docs=defaultdict(list)
    for i,r in rows.items():docs[r[2]].append(i)
    assert len(rows)==44578 and len(docs)==35017
    # Đối chiếu bản đọc trung gian với toàn bộ sổ thật để phát hiện dữ liệu cũ.
    for record in meta['docs']:
        assert record['rows']==docs[record['doc']]
        for l in record['original_lines']:
            expected=[v.isoformat() if isinstance(v,datetime) else v for v in rows[l['row']]]
            assert expected==l['values'],l['row']
    exclusions=load_workbook(ROOT/'data/HD Xuất CTY 2024.xlsx',read_only=True,data_only=True)
    exclude=set()
    for r in exclusions['HD cty 2024'].iter_rows(min_row=2,values_only=True):
        if r[0] and r[1]:exclude.update(re.findall(r'BH[A-Z]*\d+(?:\.\d+)?',str(r[12] or '')+' '+str(r[23] or '')))
    assert len(exclude&set(docs))==115
    linkfile=ROOT/'outputs/btl-bhh-20261026/LIEN_KET_BTL_BHH_2024.xlsx'
    returns=defaultdict(lambda:defaultdict(lambda:[Decimal(0)]*3))
    btl_by_bhh=defaultdict(set);unlinked=[]
    for r in load_workbook(linkfile,read_only=True,data_only=True)['Chi tiet BTL-BHH'].iter_rows(min_row=2,values_only=True):
        if r[2]:add(returns[r[2]],(r[4],r[6]),[r[7],r[9],r[11]]);btl_by_bhh[r[2]].add(r[0])
        else:unlinked.append(r[0])
    assert unlinked==[] and 'BTL00145' in btl_by_bhh['BHHT2024024972']
    sales=defaultdict(lambda:defaultdict(lambda:[Decimal(0)]*3))
    for code,indices in docs.items():
        if not code.startswith('BTL'):
            for i in indices:
                r=rows[i];add(sales[code],(r[12],r[16]),[r[17],r[20],r[32]])
    full={k for k in returns if dict(sales[k])==dict(returns[k])}
    assert len(full)==159 and 'BHHT2024024972' in full
    partial=set(returns)-full
    assert partial=={'BHHS2024000629','BHHS2024005761','BHHT2024006103','BHHT2024024887'}
    check=load_workbook(OUT/'KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx',read_only=True,data_only=True)
    pending_rows=list(check['Can kiem tra'].iter_rows(min_row=2,values_only=True));pending={r[0] for r in pending_rows}
    assert pending==set() and meta['return_mode']=='separate_adjustment_after_original_issuance'
    expected_codes={k for k in docs if not k.startswith('BTL')}-exclude
    accepted={'BHHT2024003185','BHHT2024011030','BHHT2024015896','BHHT2024017438','BHHN2024000179','BHHT2024023472','BHHT2024025057'}
    promotions={'BHHT2024013153','BHHT2024013026'}
    assert len(expected_codes)==34739 and (accepted|promotions|full)<=expected_codes and 'BHHT2024004546' in expected_codes
    impfile=OUT/'IMPORT_HOA_DON_MOI_AMIS_2024.xlsx'
    imp=load_workbook(impfile,read_only=True,data_only=True)['Hóa đơn GTGT']
    template=load_workbook(BytesIO((ROOT/'data/Du lieu Meinvoice/Mau_hoa_don_GTGT_MTT_1_thue_suat.xls').read_bytes()),read_only=True,data_only=True).active
    assert tuple(next(imp.iter_rows(min_row=9,max_row=9,max_col=17,values_only=True)))==tuple(next(template.iter_rows(min_row=9,max_row=9,values_only=True)))
    imported=defaultdict(lambda:defaultdict(lambda:[Decimal(0)]*3));invoices={};used=set();linecount=0
    compact=imp.max_column in [17,18,19]
    assert imp.max_column==19 and imp.cell(9,18).value=='Hàng khuyến mại' and imp.cell(9,19).value=='Số chứng từ'
    if compact:
        assert imp.parent.sheetnames==['Hóa đơn GTGT']
        audit=list(check['Doi chieu import'].iter_rows(min_row=2,values_only=True))
        assert len(audit)==44189
    for row_index,r in enumerate(imp.iter_rows(min_row=10,values_only=True)):
        if compact:
            a=audit[row_index]
            assert a[0]==r[0] and a[1]==row_index+10
            assert r[17]==(1 if a[2] in promotions else None)
            assert r[18]==a[2]
            r=tuple(r[:17])+(a[2],a[6],a[7],a[3],a[4],a[5])
        linecount+=1;code=r[17];row=r[20];orig=rows[row]
        assert code in expected_codes and orig[2]==code and row not in used;used.add(row)
        assert r[12]==orig[7] and r[13]==orig[16] and r[15]==orig[19] and r[21]==orig[12]
        assert canon(r[22])==canon(', '.join(sorted(btl_by_bhh[code])))
        assert [d(r[i]) for i in [14,16]]==[d(orig[17]),d(orig[20])]
        add(imported[code],(orig[12],orig[16]),[r[14],r[16],0])
        assert d(r[14])>0
        if code in promotions:assert r[21]=='KM300FULL' and d(r[14])==1 and all(d(r[i])==0 for i in [11,15,16,18,19]) and r[10]==8
        else:assert d(r[16])>0
        if code not in invoices:
            assert r[0]==len(invoices)+1 and r[1]==orig[1] and r[1].year==2024
            for c,sc in [(2,9),(3,11),(4,10),(5,9)]:assert canon(r[c])==canon(orig[sc])
            assert r[9]=='TM/CK' and r[10]==8 and all(r[c] is None for c in [6,7,8])
            invoices[code]={'number':r[0],'vat':d(r[11]),'goods':d(r[18]),'payment':d(r[19])}
        else:
            assert r[0]==invoices[code]['number'] and all(r[c] is None for c in [1,2,3,4,5,6,7,8,9,10,11,18,19])
    assert linecount==44189 and set(invoices)==expected_codes
    for code in expected_codes:
        expected=dict(sales[code])
        expected={k:v for k,v in expected.items() if any(v)}
        assert set(imported[code])==set(expected),code
        for key,v in expected.items():assert imported[code][key][:2]==v[:2],(code,key)
        inv=invoices[code]
        assert inv['goods']==sum(v[1] for v in expected.values())
        assert inv['vat']==sum(v[2] for v in expected.values())
        assert inv['goods']+inv['vat']==inv['payment']
    assert sum(v['payment'] for v in invoices.values())==53496119342
    assert sum(v['goods'] for v in invoices.values())==49533443555
    assert sum(v['vat'] for v in invoices.values())==3962675787
    assert sum(invoices[c]['payment'] for c in accepted)==47366000
    # Các dòng giữ kiểm tra vẫn nguyên 42 trường nguồn.
    raw=list(check['Dong nguon can kiem tra'].iter_rows(min_row=2,values_only=True))
    assert len(raw)==0
    for r in raw:assert [canon(v) for v in r[1:]]==[canon(v) for v in rows[r[0]]]
    omitted=list(check['Da loai va tra het'].iter_rows(min_row=2,values_only=True))
    assert {r[0] for r in omitted}==(exclude&set(docs))
    assert 'Tra lai cho dieu chinh' not in check.sheetnames
    adjustment=json.loads((WORK/'return-adjustment-import.json').read_text(encoding='utf8'))
    ready=list(check['Import dieu chinh'].iter_rows(min_row=10,values_only=True))
    assert len(ready)==219 and len({r[0] for r in ready})==163
    assert tuple(next(check['Import dieu chinh'].iter_rows(min_row=9,max_row=9,values_only=True)))==tuple(adjustment['headers'])
    assert all(all(r[i] is None for i in [1,8,9,10,11,12]) for r in ready)
    used_return_sources=set()
    for r,a in zip(ready,adjustment['audit']):
        # Các dòng nguồn giống hệt nhau có danh sách ứng viên chung; dùng mỗi dòng đúng một lần.
        candidates=[int(n.strip()) for n in a['source_rows'].split(',')]
        matches=[n for n in candidates if n not in used_return_sources and rows[n][2]==a['btl'] and rows[n][12]==a['sku'] and [d(r[i]) for i in [16,17,18]]==[-d(rows[n][26]),d(rows[n][19]),-d(rows[n][27])]]
        assert matches,(a,r)
        used_return_sources.add(matches[0]);assert r[25]==a['bhh']
    assert len(used_return_sources)==219
    assert sum(d(r[18]) for r in ready)==-225865926 and sum(d(r[21]) for r in ready)==-18069275 and sum(d(r[22]) for r in ready)==-243935201
    assert 'Cho KT tra lai' not in check.sheetnames
    drafts=[]
    assert len(drafts)==0
    assert all(all(r[i] is None for i in [12,13,14,15]) for r in drafts)
    assert len(list(check['Loai chua ro'].iter_rows(min_row=2,values_only=True)))==7
    assert list(next(check['Tong hop'].iter_rows(min_row=7,max_row=7,values_only=True)))[1:]==[35017,44578,50154037726,225865926,3994254050,53922425850]
    for name in check.sheetnames:
        for r in check[name].iter_rows():
            for cell in r:assert cell.data_type!='e',(name,cell.coordinate)
    for file in [impfile,OUT/'KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx']:
        with ZipFile(file) as z:
            for n in z.namelist():
                if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n):
                    xml=z.read(n);assert b't="e"' not in xml
                    assert 'Có thể liên quan trả lại'.encode() not in xml
                    if file==impfile:assert not re.search(rb'<(?:\w+:)?f(?:\s|>)',xml)
    for m in meta['manifest']:assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256']
    report={'passed':True,'return_mode':meta['return_mode'],'invoices':len(invoices),'lines':linecount,'goods':49533443555,'vat':3962675787,'payment':53496119342,'fully_returned_restored':159,'excluded_sales':115,'pending_sales_documents':0,'pending_returns_documents':163,'return_lines':219,'return_payment':243935201,'unresolved_exclusions':7,'partial_applied':[],'return_price_review':['BHHS2024000629','BHHS2024005761'],'unlinked_btl':unlinked,'accepted_promotions':sorted(promotions),'accepted_amis_rounding':sorted(accepted),'checks':['17 cột đúng mẫu, tiêu đề dòng 9','Đọc lại sổ thật và đối chiếu 44.578 dòng trung gian','Toàn bộ 44.189 dòng import đúng SL, đơn giá, tiền hàng bán gốc; không trừ BTL','159 BHH trả hết được khôi phục; 115 chứng từ loại trừ vẫn vắng mặt','163 BTL / 219 dòng tách riêng giữ nguyên nguồn, chưa có số HĐ và ngày điều chỉnh','Hai chênh lệch tiền trả tiếp tục giữ kiểm tra, không tự giảm giá','Tổng nguồn và số dòng khớp đủ, không trùng thiếu','Không lỗi công thức; import chỉ có giá trị','Hash các nguồn giữ nguyên'],'files':[{'path':str(f),'bytes':f.stat().st_size,'sha256':hashlib.sha256(f.read_bytes()).hexdigest()} for f in [impfile,OUT/'KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx']]}
    report['return_price_review']=[]
    report['confirmed_recorded_return_prices']=['BTL00023','BTL00024']
    report['return_import_invoices']=163;report['return_import_lines']=219;report['pending_return_price_documents']=0
    report['checks']=[c.replace('Hai chênh lệch tiền trả tiếp tục giữ kiểm tra, không tự giảm giá','BTL00023/BTL00024 dùng Giá ghi trả theo xác nhận 28/10/2026; không tự giảm giá') for c in report['checks']]
    (WORK/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
