"""Read back the exported import and compare every value with the fresh live snapshot."""
from collections import Counter
from datetime import datetime
from pathlib import Path
import hashlib
import json
import sys

from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/export-return-adjustment-163-20261101'
OUTPUT=ROOT/'outputs/import-dieu-chinh-163-20261101/IMPORT_DIEU_CHINH_163_PHIEU_TRA_HANG_2024_MOI_NHAT.xlsx'


def value(cell):
    v=cell.get('effectiveValue',{})
    assert 'errorValue' not in v
    return next(iter(v.values()),None)


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    source_path=WORK/'source-cells.json'
    source=json.loads(source_path.read_text(encoding='utf-8'))
    sh=next(s for s in source['sheets'] if s['properties']['sheetId']==700000001)
    rows=sh['data'][0]['rowData']
    matrix=[[value(rows[r].get('values',[])[c]) if c<len(rows[r].get('values',[])) else None for c in range(26)] for r in range(228)]
    wb=load_workbook(OUTPUT,read_only=True,data_only=False)
    assert wb.sheetnames==['Import dieu chinh']
    sheet=wb.active
    got=list(sheet.iter_rows(min_row=1,max_row=228,max_col=26,values_only=False))
    checked=0
    for rn,(cells,expected) in enumerate(zip(got,matrix),1):
        for cn,(cell,want) in enumerate(zip(cells,expected),1):
            actual=cell.value
            assert cell.data_type!='f', (rn,cn,'Formula retained in import')
            if isinstance(actual,datetime):
                actual=(actual-datetime(1899,12,30)).days
            assert actual==want,(rn,cn,actual,want)
            checked+=1
    body=[[c.value for c in r] for r in got[9:]]
    assert len(body)==219 and len({r[0] for r in body})==163
    assert sum(bool(r[9]) for r in body)==163
    money=[sum(r[i] or 0 for r in body) for i in [18,19,21,22]]
    assert money==[-225865926,-225865926,-18069275,-243935201]
    current=None
    for rn,(r,cells) in enumerate(zip(body,got[9:]),10):
        if r[9]:
            assert r[1]==datetime(2026,11,1) and r[11]==datetime(2026,11,1)
            assert isinstance(r[10],str) and len(r[10])==8 and cells[10].data_type=='s'
            assert r[9]=='1C26MUV' and isinstance(r[12],str)
            current=(r[0],r[9],r[10],r[11],r[24],r[25])
        assert current and current[0]==r[0] and current[5]==r[25]
        assert isinstance(r[24],str) and isinstance(r[25],str) and cells[24].data_type=='s'
        expected_reason=f"Điều chỉnh cho hóa đơn số {current[2]}, ký hiệu {current[1]}, ngày {current[3].strftime('%d/%m/%Y')} do trả lại hàng của mã đơn hàng {current[4]}"
        assert r[13]==expected_reason,(rn,'Incorrect reason')
        assert r[16]<0 and r[17]>=0 and r[18]<0
    assert not any(c.value is not None for r in sheet.iter_rows(min_row=229,max_row=240,max_col=27) for c in r)
    assert not any(c.value is not None for r in sheet.iter_rows(min_row=9,max_row=228,min_col=27,max_col=27) for c in r)
    wb.close()
    result={'passed':True,'independent_reader':'openpyxl','source_spreadsheet_id':source['spreadsheetId'],'source_sheet':'Import dieu chinh',
            'source_range':'A1:Z228','exported_sheets':1,'invoices':163,'lines':219,'columns':26,'checked_source_cells':checked,
            'header_row':9,'adjustment_date':'2026-11-01','original_invoice_references':163,'reasons_verified':219,
            'cached_order_lookups_exported_as_exact_text':True,'no_formulas_or_extra_data':True,
            'goods':money[0],'vat':money[2],'payment':money[3],
            'source_snapshot_sha256':hashlib.sha256(source_path.read_bytes()).hexdigest(),'sha256':hashlib.sha256(OUTPUT.read_bytes()).hexdigest()}
    (WORK/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))


if __name__=='__main__':
    main()
