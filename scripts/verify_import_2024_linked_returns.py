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
    assert unlinked==['BTL00145']
    sales=defaultdict(lambda:defaultdict(lambda:[Decimal(0)]*3))
    for code,indices in docs.items():
        if not code.startswith('BTL'):
            for i in indices:
                r=rows[i];add(sales[code],(r[12],r[16]),[r[17],r[20],r[32]])
    full={k for k in returns if dict(sales[k])==dict(returns[k])}
    assert len(full)==158
    partial=set(returns)-full
    assert partial=={'BHHS2024000629','BHHS2024005761','BHHT2024006103','BHHT2024024887'}
    check=load_workbook(OUT/'KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx',read_only=True,data_only=True)
    pending_rows=list(check['Can kiem tra'].iter_rows(min_row=2,values_only=True));pending={r[0] for r in pending_rows}
    assert len(pending)==12 and {'BTL00145','BHHS2024000629','BHHS2024005761'}<=pending and 'BHHT2024004546' not in pending
    expected_codes={k for k in docs if not k.startswith('BTL')}-exclude-full-pending
    assert len(expected_codes)==34570 and 'BHHT2024004546' in expected_codes
    impfile=OUT/'IMPORT_HOA_DON_MOI_AMIS_2024.xlsx'
    imp=load_workbook(impfile,read_only=True,data_only=True)['Hóa đơn GTGT']
    template=load_workbook(BytesIO((ROOT/'data/Du lieu Meinvoice/Mau_hoa_don_GTGT_MTT_1_thue_suat.xls').read_bytes()),read_only=True,data_only=True).active
    assert tuple(next(imp.iter_rows(min_row=9,max_row=9,max_col=17,values_only=True)))==tuple(next(template.iter_rows(min_row=9,max_row=9,values_only=True)))
    imported=defaultdict(lambda:defaultdict(lambda:[Decimal(0)]*3));invoices={};used=set();linecount=0
    for r in imp.iter_rows(min_row=10,values_only=True):
        linecount+=1;code=r[17];row=r[20];orig=rows[row]
        assert code in expected_codes and orig[2]==code and row not in used;used.add(row)
        assert r[12]==orig[7] and r[13]==orig[16] and r[15]==orig[19] and r[21]==orig[12]
        assert canon(r[22])==canon(', '.join(sorted(btl_by_bhh[code])))
        if code not in returns:assert [d(r[i]) for i in [14,16]]==[d(orig[17]),d(orig[20])]
        add(imported[code],(orig[12],orig[16]),[r[14],r[16],0])
        assert d(r[14])>0 and d(r[16])>0
        if code not in invoices:
            assert r[0]==len(invoices)+1 and r[1]==orig[1] and r[1].year==2024
            for c,sc in [(2,9),(3,11),(4,10),(5,9)]:assert canon(r[c])==canon(orig[sc])
            assert r[9]=='TM/CK' and r[10]==8 and all(r[c] is None for c in [6,7,8])
            invoices[code]={'number':r[0],'vat':d(r[11]),'goods':d(r[18]),'payment':d(r[19])}
        else:
            assert r[0]==invoices[code]['number'] and all(r[c] is None for c in [1,2,3,4,5,6,7,8,9,10,11,18,19])
    assert linecount==43960 and set(invoices)==expected_codes
    for code in expected_codes:
        expected={k:[a-b for a,b in zip(v,returns[code].get(k,[0,0,0]))] for k,v in sales[code].items()}
        expected={k:v for k,v in expected.items() if any(v)}
        assert set(imported[code])==set(expected),code
        for key,v in expected.items():assert imported[code][key][:2]==v[:2],(code,key)
        inv=invoices[code]
        assert inv['goods']==sum(v[1] for v in expected.values())
        assert inv['vat']==sum(v[2] for v in expected.values())
        assert inv['goods']+inv['vat']==inv['payment']
    assert sum(v['payment'] for v in invoices.values())==53202924891
    assert sum(v['goods'] for v in invoices.values())==49261967212
    assert sum(v['vat'] for v in invoices.values())==3940957679
    # Các dòng giữ kiểm tra vẫn nguyên 42 trường nguồn.
    raw=list(check['Dong nguon can kiem tra'].iter_rows(min_row=2,values_only=True))
    assert len(raw)==17
    for r in raw:assert [canon(v) for v in r[1:]]==[canon(v) for v in rows[r[0]]]
    omitted=list(check['Da loai va tra het'].iter_rows(min_row=2,values_only=True))
    assert {r[0] for r in omitted}==(exclude&set(docs))|full
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
    report={'passed':True,'invoices':len(invoices),'lines':linecount,'goods':49261967212,'vat':3940957679,'payment':53202924891,'fully_returned':158,'excluded_sales':115,'pending_documents':12,'unresolved_exclusions':7,'partial_applied':sorted(partial-pending),'partial_review':sorted(partial&pending),'unlinked_btl':unlinked,'checks':['17 cột đúng mẫu, tiêu đề dòng 9','Đọc lại sổ thật và đối chiếu 44.578 dòng trung gian','Từng dòng import đúng BHH, SKU, ngày, giá và số dòng nguồn','SL, tiền hàng và VAT mỗi BHH bằng bán gốc trừ trả lại thực tế','Không còn điều kiện dò ứng viên trả lại','158 BHH trả hết và 115 BHH loại trừ vắng mặt trong import','12 chứng từ giữ nguyên dòng nguồn, tách kiểm tra','Tổng nguồn và số dòng khớp đủ, không trùng thiếu','Không lỗi công thức; import chỉ có giá trị','Hash của bốn nguồn giữ nguyên'],'files':[{'path':str(f),'bytes':f.stat().st_size,'sha256':hashlib.sha256(f.read_bytes()).hexdigest()} for f in [impfile,OUT/'KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx']]}
    (WORK/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
