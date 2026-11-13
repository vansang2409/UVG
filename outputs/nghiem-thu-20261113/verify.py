"""Kiểm chứng file xuất bằng trình đọc độc lập và kiểm tra nguồn không bị sửa."""
from pathlib import Path
from zipfile import ZipFile
from collections import Counter,defaultdict
from openpyxl import load_workbook
import json,hashlib,sys,re,io
sys.stdout.reconfigure(encoding='utf-8')
ROOT=Path('D:/UVG');OUT=Path(__file__).parent
def read(n):return json.loads((OUT/n).read_text(encoding='utf-8'))
records=read('mapping.json');summary=read('summary.json');ledger=read('ledger.json');manifest=read('manifest.json');controls=read('amis-controls.json')
expected={2024:[r for r in records if r['year']==2024],2025:[r for r in records if r['year']==2025]}
assert len({r['id'] for r in records})==summary['docs']
for r in records:
    if r['invoice']:
        assert re.fullmatch(r'\d{8}',r['invoice'])
        assert r['symbol']+'/'+r['invoice'] in ledger
        assert r['symbol']+'/'+r['invoice'] in r['keys']
    else:assert r['symbol'] is None
    if r['status'] in ['Khớp tiền và VAT','Khớp theo nhóm chứng từ']:assert r['delta']==[0,0,0]
# Đối chiếu các tổng sổ mới với dòng Tổng cộng nguyên bản.
for y in (2024,2025):
    cs=[r for r in controls if str(y) in r['file']]
    totals=[sum(r['totals'][i] for r in cs) for i in range(3)]
    assert totals==summary['amis_totals'][str(y)]
    footer=cs[-1]['footers'][-1]['values']
    assert totals[0]==sum(float(footer[k])*(-1 if k in ['AD','AK','AL'] else 1) for k in ['Y','AD','AK','AL'])
    assert totals[1]==float(footer['AP']) and totals[2]==float(footer['AR'])
for m in manifest:
    if 'member' not in m:assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256']
path=OUT/'MAP_SO_KY_HIEU_HD_AMIS_1311.xlsx'
with ZipFile(path) as z:assert z.testzip() is None
w=load_workbook(path,read_only=True,data_only=True)
output_counts={};errors=[]
for y,rs in expected.items():
    s=w['Map '+str(y)];got=[];sums=[0,0,0]
    for i,row in enumerate(s.iter_rows(min_row=7,values_only=True)):
        e=rs[i]
        assert row[0]==e['doc'],(y,i,'doc')
        assert (row[1] or None)==e['invoice'] and (row[2] or None)==e['symbol'],(y,i,'invoice')
        assert row[3]==e['status'],(y,i,'status')
        if row[1]:assert isinstance(row[1],str) and len(row[1])==8
        for value in row:
            if isinstance(value,str) and re.fullmatch(r'#(?:REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|SPILL!|CALC!)',value):errors.append([y,i,value])
        got.append(row[0])
    assert len(got)==len(rs)
    output_counts[y]=len(got)
    print('VERIFIED',y,len(got),flush=True)
assert not errors
cached=w['Tổng hợp']
assert [cached.cell(8,c).value for c in [2,3,4]]==[summary['docs'],summary['mapped'],summary['unmapped']]
for row in cached.iter_rows(min_row=12,max_row=19,max_col=3,values_only=True):
    if row[0]:
        for i,y in enumerate((2024,2025),1):assert row[i]==sum(r['status']==row[0] for r in expected[y]),('status_count',row,y)
review=[r for r in records if r['status'] not in ['Khớp tiền và VAT','Khớp theo nhóm chứng từ']]
review_rows=list(w['Cần kiểm tra'].iter_rows(min_row=7,values_only=True));assert len(review_rows)==len(review)
w.close()
# Kiểm độc lập một số HĐ nguồn ở đầu/cuối mỗi năm.
sample=[expected[y][i] for y in (2024,2025) for i in (0,len(expected[y])//2,-1)]
sourcechecks=[]
archive=next((ROOT/'data/1311/meinvoice').glob('Bang_ke_hoa_don_da_su_dung_*.zip'))
with ZipFile(archive) as z:
    for r in sample:
        if not r['invoice']:continue
        member,rn=r['source_invoice'].rsplit('!',1);rn=int(rn)
        sbook=load_workbook(io.BytesIO(z.read(member)),read_only=True,data_only=True)
        raw=next(sbook.worksheets[0].iter_rows(min_row=rn,max_row=rn,values_only=True))
        assert str(raw[1]).upper()==r['symbol'] and str(raw[2]).removesuffix('.0').zfill(8)==r['invoice']
        sourcechecks.append({'doc':r['doc'],'invoice':r['symbol']+'/'+r['invoice'],'source':r['source_invoice']});sbook.close()
result={'passed':True,'output_rows':output_counts,'mapped':summary['mapped'],'unmapped':summary['unmapped'],'review':len(review),'source_files_unchanged':6,'source_header_spot_checks':sourcechecks,'invoice_ids_are_text_8_digits':True,'amis_totals_match_source_footers':True,'formula_errors':errors,'output_sha256':hashlib.sha256(path.read_bytes()).hexdigest()}
(OUT/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False),flush=True)
