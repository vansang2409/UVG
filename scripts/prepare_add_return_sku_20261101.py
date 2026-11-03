"""Map each row of the user's local import to a verified AMIS product code."""
from collections import defaultdict, Counter
from pathlib import Path
import hashlib
import json
import sys

from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/add-return-sku-20261101'
TARGET=Path(r'C:\Users\thuong ngo\Downloads\CHECK HOA DON XUAT MOI 2024 31.10.2026\IMPORT_DIEU_CHINH_163_PHIEU_TRA_HANG_2024_MOI_NHAT.xlsx')


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    WORK.mkdir(parents=True,exist_ok=True)
    raw=TARGET.read_bytes()
    backup=WORK/'before.xlsx'
    if backup.exists():
        assert backup.read_bytes()==raw, 'The input changed since the original snapshot'
    else:
        backup.write_bytes(raw)
    w=load_workbook(TARGET,read_only=False,data_only=False)
    assert w.sheetnames==['Import dieu chinh']
    s=w.active
    header=list(next(s.iter_rows(min_row=9,max_row=9,max_col=25,values_only=True)))
    assert s.max_row==228 and s.max_column==25 and header[24]=='Mã đơn hàng'
    assert all('Mã hàng' not in str(v) for v in header)
    assert all(c.data_type!='f' for row in s.iter_rows() for c in row)
    source_file=ROOT/'scripts/data/amis-2024-linked-20261026/return-adjustment-import.json'
    reference_file=ROOT/'scripts/data/return-adjustment-issued-20261101/prepared.json'
    source=json.loads(source_file.read_text(encoding='utf-8'))
    references=json.loads(reference_file.read_text(encoding='utf-8'))
    invoice_doc={r['invoice']:r['doc'] for r in references['audit']}
    invoice_order={r['invoice']:r['order'] for r in references['audit']}
    index=defaultdict(list)
    for row,a in zip(source['rows'],source['audit']):
        assert a['sku']==a['source'][4] and a['sku']
        key=(row[25],row[14],row[15],row[16],row[17],row[18])
        index[key].append(a)
    result=[]
    current=None
    body=list(s.iter_rows(min_row=10,max_row=228,max_col=25,values_only=True))
    for rn,r in enumerate(body,10):
        if r[9]:
            key=r[9]+'/'+r[10]
            assert key in invoice_doc and str(r[24])==invoice_order[key]
            current={'seq':r[0],'invoice':key,'doc':invoice_doc[key]}
        assert current and r[0]==current['seq'] and str(r[24])==invoice_order[current['invoice']]
        matches=index.get((current['doc'],r[14],r[15],r[16],r[17],r[18]),[])
        codes={a['sku'] for a in matches}
        assert len(codes)==1,(rn,current['doc'],codes)
        result.append({'row':rn,'code':next(iter(codes)),'invoice':current['invoice'],'doc':current['doc'],
                       'name':r[14],'unit':r[15],'quantity':r[16],'price':r[17],'goods':r[18],
                       'btl':sorted({a['btl'] for a in matches}),'amis_source_rows':sorted({a['source_rows'] for a in matches})})
    assert len(result)==219 and len({r[0] for r in body})==163
    assert [sum(r[i] or 0 for r in body) for i in [18,21,22]]==[-225865926,-18069275,-243935201]
    stage=TARGET.with_name('._codex_ADD_SKU_163_20261101.tmp.xlsx')
    prepared={'target':str(TARGET),'backup':str(backup),'staged':str(stage),'before_sha256':sha(TARGET),
              'source_audit':str(source_file),'source_audit_sha256':sha(source_file),'source_references':str(reference_file),'source_references_sha256':sha(reference_file),
              'sheet':s.title,'header_row':9,'new_column':'Z','header':'Mã hàng hóa/dịch vụ','invoices':163,'rows':219,
              'sku_counts':dict(Counter(x['code'] for x in result)),'mapping':result}
    w.close()
    (WORK/'prepared.json').write_text(json.dumps(prepared,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'rows':219,'invoices':163,'unique_codes':len(prepared['sku_counts']),'column':'Z','header':prepared['header']},ensure_ascii=False))


if __name__=='__main__':
    main()
