from pathlib import Path
from collections import Counter
import hashlib,json,os,shutil,sys
from copy import deepcopy
from io import BytesIO
from zipfile import ZipFile,ZIP_DEFLATED
from openpyxl import load_workbook
from lxml import etree

work=Path(__file__).resolve().parent
root=work.parents[2]
sys.path.insert(0,str(root/'scripts'))
from verify_install_return_sku_20261101 import visual_style,preserve_native_filter

sys.stdout.reconfigure(encoding='utf-8')
p=json.loads((work/'prepared.json').read_text(encoding='utf-8'))
target,backup,staged=Path(p['target']),Path(p['backup']),Path(p['staged'])
sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest()
assert sha(target)==sha(backup)==p['before_sha256']
filter_restored=preserve_native_filter(backup,staged)
# Preserve the user's exact pane/scroll/selection state; artifact normalizes it.
with ZipFile(backup) as old,ZipFile(staged) as new:
    part='xl/worksheets/sheet1.xml'
    before_xml=etree.fromstring(old.read(part));after_xml=etree.fromstring(new.read(part))
    ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    saved=before_xml.find('s:sheetViews',ns);found=after_xml.find('s:sheetViews',ns)
    assert saved is not None and found is not None
    after_xml.replace(found,deepcopy(saved))
    output=BytesIO()
    with ZipFile(output,'w',compression=ZIP_DEFLATED) as preserved:
        for info in new.infolist():
            preserved.writestr(info,etree.tostring(after_xml,xml_declaration=True,encoding='UTF-8',standalone=True) if info.filename==part else new.read(info.filename))
staged.write_bytes(output.getvalue())
bw=load_workbook(backup,read_only=False,data_only=False)
aw=load_workbook(staged,read_only=False,data_only=False)
assert bw.sheetnames==aw.sheetnames==['Import dieu chinh']
b,a=bw.active,aw.active
changes={r['row']:r['number'] for r in p['changes']}
changed=0
for rn in range(1,229):
    for cn in range(1,27):
        bc,ac=b.cell(rn,cn),a.cell(rn,cn)
        expected=changes[rn] if cn==11 and rn in changes else bc.value
        assert ac.value==expected,(rn,cn,bc.value,ac.value,expected)
        assert visual_style(ac)==visual_style(bc),(rn,cn,'Style changed')
        if cn==11 and rn in changes:
            assert ac.data_type=='s' and len(ac.value)==8 and bc.value is None
            changed+=1
        else:
            assert ac.data_type==bc.data_type,(rn,cn,'Cell type changed')
assert changed==56
assert a.freeze_panes==b.freeze_panes and a.auto_filter.ref==b.auto_filter.ref
assert set(map(str,a.merged_cells.ranges))==set(map(str,b.merged_cells.ranges))
assert len(a.data_validations.dataValidation)==len(b.data_validations.dataValidation)
assert len(a.conditional_formatting)==len(b.conditional_formatting) and list(a.tables)==list(b.tables)
assert a.max_row==b.max_row==228 and a.max_column==b.max_column==26
by_sequence={};body=list(a.iter_rows(min_row=10,max_row=228,max_col=26,values_only=True))
for r in body:
    assert r[10] and isinstance(r[10],str)
    if r[0] not in by_sequence:by_sequence[r[0]]=r[10]
    assert r[10]==by_sequence[r[0]]
assert len(by_sequence)==163 and len(body)==219
money=[sum(r[i] or 0 for r in body) for i in [18,21,22]]
before_money=[sum(r[i] or 0 for r in b.iter_rows(min_row=10,max_row=228,max_col=26,values_only=True)) for i in [18,21,22]]
assert money==before_money
bw.close();aw.close()
staged_sha=sha(staged)
assert sha(target)==p['before_sha256'],'User file changed before saving'
os.replace(staged,target)
assert sha(target)==staged_sha
inspection=Path(str(staged)+'.inspect.ndjson')
if inspection.exists():shutil.move(str(inspection),str(work/'artifact-export-inspection.ndjson'))
v={'passed':True,'target':str(target),'filled_cells':56,'column':'K','multiline_invoices':43,
   'invoices':163,'rows_with_adjusted_invoice_number':219,'numbers_preserved_as_eight_digit_text':True,
   'all_other_values_styles_preserved':True,'filter_freeze_merges_preserved':True,'native_filter_restored':filter_restored,
   'money':money,'sha256':staged_sha}
(work/'verification.json').write_text(json.dumps(v,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(v,ensure_ascii=False))
