"""Đọc độc lập file đã xuất để kiểm đủ chứng từ, tiền và kiểu mã."""
from pathlib import Path
import json
from openpyxl import load_workbook
P=Path(__file__).resolve().parents[1]/'outputs/phan-loai-2025-20261105'
data=json.loads((P/'workbook-data-compact.json').read_text(encoding='utf-8'))
wb=load_workbook(P/'PHAN_LOAI_XU_LY_HD_2025.xlsx',read_only=True,data_only=False)
seen=set();sums={};errors=[]
for t in data['tables']:
    ws=wb[t['name']];n=0;money=0;vat=0
    for row in ws.iter_rows(min_row=7,max_col=len(t['headers'])):
        if not row[0].value:continue
        n+=1
        if t['name']!='HĐ cần kiểm riêng':
            assert row[0].value not in seen;seen.add(row[0].value)
            money+=row[3].value or 0;vat+=row[4].value or 0
            assert row[2].value is None or isinstance(row[2].value,str)
            assert row[12].value is None or isinstance(row[12].value,str)
        for c in row:
            if c.data_type=='e':errors.append((ws.title,c.coordinate,c.value))
    assert n==len(t['rows']),(t['name'],n,len(t['rows']))
    sums[t['name']]={'rows':n,'money':money,'vat':vat}
assert len(seen)==65494
assert sum(x['money'] for k,x in sums.items() if k!='HĐ cần kiểm riêng')==data['source_total']
assert not errors,errors[:5]
wb.close()
cached=load_workbook(P/'PHAN_LOAI_XU_LY_HD_2025.xlsx',read_only=True,data_only=True)
has_cache=cached['Tổng hợp']['B12'].value is not None
if has_cache:
    assert cached['Tổng hợp']['B12'].value==65494
    assert cached['Tổng hợp']['C12'].value==data['source_total']
    assert cached['Tổng hợp']['D12'].value==data['source_vat']
result={'source_rows_verified':True,'summary_recalculation':'cached_verified' if has_cache else 'verify_after_native_import','sheets':sums,'docs':len(seen),'money':data['source_total'],'vat':data['source_vat'],'cell_errors':errors}
(P/'xlsx-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
