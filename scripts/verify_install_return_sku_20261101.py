"""Check the added SKU column and preserve the user's existing import exactly."""
from collections import defaultdict, Counter
from pathlib import Path
import hashlib
import json
import os
import posixpath
import shutil
import sys
from copy import deepcopy
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl import load_workbook
from lxml import etree

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/add-return-sku-20261101'


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def color(v):
    if v is None:return None
    if v.type=='rgb':return ('rgb',v.rgb[-6:])
    return (v.type,v.indexed if v.type=='indexed' else v.theme,v.tint)


def visual_style(c):
    font=c.font
    alignment=c.alignment
    border=[]
    for name in ['left','right','top','bottom','diagonal']:
        side=getattr(c.border,name)
        border.append((side.style,color(side.color)) if side is not None and side.style else None)
    return {'font':(font.name,font.sz,bool(font.b),bool(font.i),font.u,bool(font.strike),color(font.color)),
            'fill':(c.fill.patternType,color(c.fill.fgColor),color(c.fill.bgColor)),
            'alignment':(alignment.horizontal or 'general',alignment.vertical or 'bottom',bool(alignment.wrap_text),bool(alignment.shrink_to_fit),alignment.text_rotation or 0,alignment.indent or 0),
            'number_format':c.number_format,'borders':border,'comment':c.comment.text if c.comment else None}


def preserve_native_filter(before_path,staged):
    """Artifact export omits sheet autoFilter; restore only that existing XML node."""
    ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rel_ns='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    def sheet_part(archive):
        wb=etree.fromstring(archive.read('xl/workbook.xml'))
        sheet=next(x for x in wb.find('s:sheets',ns) if x.get('name')=='Import dieu chinh')
        rid=sheet.get('{'+rel_ns+'}id')
        rels=etree.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        target=next(x.get('Target') for x in rels if x.get('Id')==rid)
        return target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
    with ZipFile(before_path) as original,ZipFile(staged) as exported:
        before_xml=etree.fromstring(original.read(sheet_part(original)))
        part=sheet_part(exported)
        after_xml=etree.fromstring(exported.read(part))
        saved=before_xml.find('s:autoFilter',ns)
        found=after_xml.find('s:autoFilter',ns)
        if saved is None:return False
        if found is not None:
            if etree.tostring(found)==etree.tostring(saved):return False
            after_xml.remove(found)
        following={'sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'}
        position=next((i for i,x in enumerate(after_xml) if etree.QName(x).localname in following),len(after_xml))
        after_xml.insert(position,deepcopy(saved))
        payload=etree.tostring(after_xml,xml_declaration=True,encoding='UTF-8',standalone=True)
        output=BytesIO()
        with ZipFile(output,'w',compression=ZIP_DEFLATED) as packaged:
            for info in exported.infolist():
                packaged.writestr(info,payload if info.filename==part else exported.read(info.filename))
    staged.write_bytes(output.getvalue())
    return True


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    p=json.loads((WORK/'prepared.json').read_text(encoding='utf-8'))
    before_path,staged,target=Path(p['backup']),Path(p['staged']),Path(p['target'])
    assert sha(before_path)==p['before_sha256'] and sha(target)==p['before_sha256']
    assert sha(Path(p['source_audit']))==p['source_audit_sha256']
    assert sha(Path(p['source_references']))==p['source_references_sha256']
    native_filter_restored=preserve_native_filter(before_path,staged)
    before=load_workbook(before_path,read_only=False,data_only=False)
    after=load_workbook(staged,read_only=False,data_only=False)
    assert before.sheetnames==after.sheetnames==['Import dieu chinh']
    b,a=before.active,after.active
    assert (a.max_row,a.max_column)==(228,26)
    style_differences=[]
    for row in range(1,229):
        for col in range(1,26):
            bc,ac=b.cell(row,col),a.cell(row,col)
            assert bc.value==ac.value,(row,col,'Existing value changed',bc.value,ac.value)
            assert bc.data_type==ac.data_type,(row,col,'Existing cell type changed')
            if visual_style(bc)!=visual_style(ac):
                style_differences.append({'cell':bc.coordinate,'before':visual_style(bc),'after':visual_style(ac)})
    assert not style_differences, style_differences[:3]
    assert set(map(str,b.merged_cells.ranges))==set(map(str,a.merged_cells.ranges))
    assert b.freeze_panes==a.freeze_panes and b.auto_filter.ref==a.auto_filter.ref
    assert len(b.data_validations.dataValidation)==len(a.data_validations.dataValidation)
    assert len(b.conditional_formatting)==len(a.conditional_formatting)
    assert list(b.tables)==list(a.tables)
    assert a['Z9'].value=='Mã hàng hóa/dịch vụ'
    assert all(a.cell(r,26).value is None for r in range(1,9))
    refs=json.loads(Path(p['source_references']).read_text(encoding='utf-8'))
    docs={x['invoice']:x['doc'] for x in refs['audit']}
    audit=json.loads(Path(p['source_audit']).read_text(encoding='utf-8'))
    index=defaultdict(set)
    for r,source in zip(audit['rows'],audit['audit']):
        index[(r[25],r[14],r[15],r[16],r[17],r[18])].add(source['source'][4])
    current=None
    skus=[]
    for rn,row in enumerate(a.iter_rows(min_row=10,max_row=228,max_col=26,values_only=True),10):
        if row[9]:current=(row[0],docs[row[9]+'/'+row[10]])
        assert current and row[0]==current[0]
        codes=index[(current[1],row[14],row[15],row[16],row[17],row[18])]
        assert len(codes)==1 and row[25]==next(iter(codes))
        assert row[25]==p['mapping'][rn-10]['code'] and a.cell(rn,26).data_type=='s'
        skus.append(row[25])
    assert dict(Counter(skus))==p['sku_counts'] and len(skus)==219
    body=list(a.iter_rows(min_row=10,max_row=228,max_col=25,values_only=True))
    assert len({r[0] for r in body})==163
    assert [sum(r[i] or 0 for r in body) for i in [18,21,22]]==[-225865926,-18069275,-243935201]
    before.close();after.close()
    staged_hash=sha(staged)
    assert sha(target)==p['before_sha256'], 'The user file changed during verification'
    os.replace(staged,target)
    assert sha(target)==staged_hash
    inspection=Path(str(staged)+'.inspect.ndjson')
    if inspection.exists():
        shutil.move(str(inspection),str(WORK/'artifact-export-inspection.ndjson'))
    report={'passed':True,'edited_user_file_in_place':str(target),'added_column':'Z','header':p['header'],
            'invoices':163,'rows_filled':219,'unique_product_codes':len(set(skus)),
            'mapping':'Exact original invoice -> AMIS document + item name/unit/quantity/unit price/amount',
            'existing_cells_unchanged':228*25,'existing_visual_styles_preserved':True,'merges_filter_freeze_preserved':True,
            'native_filter_restored_after_artifact_export':native_filter_restored,
            'money':[-225865926,-18069275,-243935201],'source_sha256':p['before_sha256'],'sha256':staged_hash}
    (WORK/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False))


if __name__=='__main__':
    main()
