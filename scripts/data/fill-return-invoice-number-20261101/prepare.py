from collections import Counter
from pathlib import Path
import hashlib,json
from openpyxl import load_workbook

work=Path(__file__).resolve().parent
target=Path(r'C:\Users\thuong ngo\Downloads\CHECK HOA DON XUAT MOI 2024 31.10.2026\IMPORT_DIEU_CHINH_163_PHIEU_TRA_HANG_2024_MOI_NHAT.xlsx')
backup=work/'before.xlsx'
raw=target.read_bytes()
if backup.exists():
    assert backup.read_bytes()==raw,'The input changed since the snapshot'
else:
    backup.write_bytes(raw)
w=load_workbook(backup,read_only=False,data_only=False)
assert w.sheetnames==['Import dieu chinh']
s=w.active
assert s['K9'].value=='Số hóa đơn bị điều chỉnh (*)' and s.max_row==228 and s.max_column==26
seen=set();current=None;changes=[];groups=Counter()
for rn,r in enumerate(s.iter_rows(min_row=10,max_row=228,max_col=26,values_only=True),10):
    seq=r[0]
    if current is None or seq!=current['sequence']:
        assert seq not in seen and isinstance(r[10],str) and len(r[10])==8
        seen.add(seq)
        current={'sequence':seq,'number':r[10],'series':r[9]}
    else:
        assert r[10] in (None,current['number']),('Conflicting invoice number',rn)
        if r[9]:assert r[9]==current['series']
        if r[10] is None:changes.append({'row':rn,'number':current['number'],'sequence':seq})
    groups[seq]+=1
assert len(seen)==163 and sum(groups.values())==219 and len(changes)==56
p={'target':str(target),'backup':str(backup),'staged':str(target.with_name('._codex_FILL_INVOICE_NUMBER_163_20261101.tmp.xlsx')),
   'before_sha256':hashlib.sha256(raw).hexdigest(),'sheet':s.title,'column':'K','invoices':163,'lines':219,
   'multiline_invoices':sum(v>1 for v in groups.values()),'changes':changes}
(work/'prepared.json').write_text(json.dumps(p,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
w.close()
print(json.dumps({'invoices':163,'lines':219,'cells_to_fill':len(changes),'multiline_invoices':p['multiline_invoices']}))
