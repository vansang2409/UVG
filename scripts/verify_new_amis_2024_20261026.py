"""Đọc độc lập hai workbook đã xuất, đối chiếu từng dòng và bảo toàn nguồn."""
from pathlib import Path
from io import BytesIO
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from zipfile import ZipFile
import hashlib
import json
import re
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT/'scripts/data/new-amis-2024-20261026'
OUT = ROOT/'outputs/amis-2024-moi-20261026'


def canon(v):
    if v in (None, ''):
        return None
    if isinstance(v, datetime):
        return v.isoformat()
    return v


def main():
    p=json.loads((WORK/'prepared.json').read_text(encoding='utf8'))
    for source in p['manifest']:
        assert hashlib.sha256((ROOT/source['file']).read_bytes()).hexdigest()==source['sha256']
    template=ROOT/'data/Du lieu Meinvoice/Mau_hoa_don_GTGT_MTT_1_thue_suat.xls'
    template_headers=list(load_workbook(BytesIO(template.read_bytes()),read_only=True,data_only=True).active.iter_rows(min_row=9,max_row=9,values_only=True))[0]
    source=load_workbook(ROOT/'data/Du lieu Amis/So_chi_tiet_ban_hang_2024.xlsx',read_only=True,data_only=True).active
    source_rows={i:list(r) for i,r in enumerate(source.iter_rows(min_row=5,values_only=True),5) if r[2]}
    source_docs=defaultdict(list)
    for i,r in source_rows.items():
        source_docs[str(r[2])].append(i)
    import_path=OUT/'IMPORT_HOA_DON_MOI_AMIS_2024.xlsx'
    check_path=OUT/'KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx'
    imported=load_workbook(import_path,read_only=True,data_only=True)
    s=imported['Hóa đơn GTGT']
    assert tuple(c.value for c in next(s.iter_rows(min_row=9,max_row=9,max_col=17)))==template_headers
    ready=[d for d in p['docs'] if d['category']=='Đưa vào import']
    flat=[(n,d,j,l) for n,d in enumerate(ready,1) for j,l in enumerate(d['lines'])]
    actual=list(s.iter_rows(min_row=10,values_only=True))
    assert len(actual)==len(flat)==24677
    seen=set()
    for r,(n,d,j,l) in zip(actual,flat):
        original=source_rows[l['row']]
        first=j==0
        expected=[n,original[1] if first else None,original[9] if first else None,
                  original[11] if first else None,original[10] if first else None,
                  original[9] if first else None,None,None,None,'TM/CK' if first else None,
                  d['rate'] if first else None,d['sums']['vat'] if first else None,
                  original[7],original[16],original[17],original[19],original[20],d['doc'],
                  d['sums']['goods'] if first else None,d['sums']['payment'] if first else None,l['row'],original[12]]
        assert [canon(v) for v in r]==[canon(v) for v in expected],(n,l['row'])
        assert l['row'] not in seen
        seen.add(l['row'])
        assert not any(d['doc']==code for e in p['exclusions'] for code in e['codes'])
        if first:
            assert original[1].year==2024
            assert d['rate']==8
    review=load_workbook(check_path,read_only=True,data_only=True)
    groups={}
    for sheet,category in [('Can kiem tra','Cần kiểm tra'),('Da loai','Đã loại')]:
        rs=[r for r in review[sheet].iter_rows(min_row=2,values_only=True) if r[0]]
        ds=[d for d in p['docs'] if d['category']==category]
        assert len(rs)==len(ds)
        for r,d in zip(rs,ds):
            assert r[0]==d['doc'] and r[9]==len(d['lines'])
            assert r[11]=='; '.join(d['reasons'])
            for c,k in [(5,'goods'),(6,'returns'),(7,'vat'),(8,'payment')]:
                assert Decimal(str(r[c]))==Decimal(str(d['sums'][k]))
            groups[d['doc']]=category
    for d in ready:
        assert d['doc'] not in groups
        groups[d['doc']]='Đưa vào import'
    assert set(groups)==set(source_docs) and len(groups)==35017
    raw=[r for r in review['Chi tiet can kiem tra'].iter_rows(min_row=2,values_only=True) if r[0]]
    assert len(raw)==19731
    for r in raw:
        assert [canon(v) for v in r[1:]]==[canon(v) for v in source_rows[r[0]]],r[0]
        assert groups[str(source_rows[r[0]][2])]=='Cần kiểm tra'
    excluded_rows={i for d in p['docs'] if d['category']=='Đã loại' for i in d['rows']}
    assert len(excluded_rows)==170 and len(excluded_rows|seen|{r[0] for r in raw})==44578
    assert not (excluded_rows&seen)
    # Kiểm tra cả ứng viên trả lại, không chỉ nhãn phân loại.
    for ret,row,candidates in p['return_evidence']:
        assert groups[ret]=='Cần kiểm tra'
        for sale,source_row in candidates:
            assert groups[sale]!='Đưa vào import'
    for file in (import_path,check_path):
        z=ZipFile(file)
        errors=[]
        for name in z.namelist():
            if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name):
                xml=z.read(name)
                assert b't="e"' not in xml,(file.name,name)
                if file==import_path:
                    assert not re.search(rb'<(?:\w+:)?f(?:\s|>)',xml),(file.name,name)
    total=[c.value for c in next(review['Tong hop'].iter_rows(min_row=5,max_row=5))]
    assert total[1:]==[35017,44578,50154037726,225865926,3994254050,53922425850],total
    for source in p['manifest']:
        assert hashlib.sha256((ROOT/source['file']).read_bytes()).hexdigest()==source['sha256']
    report={'passed':True,'source_documents':35017,'source_lines':44578,'summary':p['summary'],
            'checks':['17 cột mẫu đúng thứ tự; dòng tiêu đề 9','Mọi dòng import khớp giá trị nguồn',
                      'Mọi dòng kiểm tra khớp 42 cột nguồn','Phân nhóm đủ, không trùng/thiếu dòng',
                      'Không có BHH loại trừ hoặc ứng viên trả lại trong import',
                      'Tổng 3 nhóm khớp dòng tổng sổ AMIS','Không có lỗi ô Excel hoặc công thức trong import','Hash nguồn giữ nguyên'],
            'template_sha256':hashlib.sha256(template.read_bytes()).hexdigest(),
            'files':[{'path':str(f),'sha256':hashlib.sha256(f.read_bytes()).hexdigest(),'bytes':f.stat().st_size} for f in [import_path,check_path]]}
    (WORK/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__=='__main__':
    main()
