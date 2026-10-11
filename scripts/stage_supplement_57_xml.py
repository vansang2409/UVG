"""Preservation fallback after artifact-tool full-workbook export exhausted Node heap.

Patch only 11 numeric cells and their number format in OOXML; do not save through
openpyxl or regenerate sheets. All other ZIP entries retain their original bytes.
"""
from pathlib import Path
from zipfile import ZipFile
from copy import deepcopy
from lxml import etree as E
import json
import hashlib

ROOT = Path(__file__).resolve().parents[1]
H = ROOT / 'scripts/data/check-review'
p = ROOT / 'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx'
d = json.loads((H / 'supplement-57-price-changes.json').read_text(encoding='utf-8'))
assert hashlib.sha256(p.read_bytes()).hexdigest() == d['base_sha256']
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
tag = lambda s: '{' + ns['m'] + '}' + s
with ZipFile(p) as z:
    workbook = E.fromstring(z.read('xl/workbook.xml'))
    sheet = next(s for s in workbook.find('m:sheets', ns) if s.get('name') == 'Import dieu chinh')
    rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    rels = E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    target = next(r.get('Target') for r in rels if r.get('Id') == rid)
    sheet_path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
    s = E.fromstring(z.read(sheet_path))
    styles = E.fromstring(z.read('xl/styles.xml'))
    formats = styles.find('m:numFmts', ns)
    if formats is None:
        formats = E.Element(tag('numFmts'), count='0')
        styles.insert(0, formats)
    format_id = max([163] + [int(x.get('numFmtId')) for x in formats]) + 1
    E.SubElement(formats, tag('numFmt'), numFmtId=str(format_id), formatCode='#,##0.0;[Red](#,##0.0);0')
    formats.set('count', str(len(formats)))
    xfs = styles.find('m:cellXfs', ns)
    style_map = {}
    for c in d['changes']:
        cell = s.find("m:sheetData/m:row/m:c[@r='" + c['cell'] + "']", ns)
        assert cell is not None and cell.find('m:f', ns) is None
        v = cell.find('m:v', ns)
        assert float(v.text) == c['old']
        old_style = int(cell.get('s', '0'))
        if old_style not in style_map:
            xf = deepcopy(xfs[old_style])
            xf.set('numFmtId', str(format_id))
            xf.set('applyNumberFormat', '1')
            style_map[old_style] = len(xfs)
            xfs.append(xf)
        cell.set('s', str(style_map[old_style]))
        v.text = str(c['new'])
    xfs.set('count', str(len(xfs)))
    replacements = {sheet_path: E.tostring(s, encoding='utf-8', xml_declaration=True), 'xl/styles.xml': E.tostring(styles, encoding='utf-8', xml_declaration=True)}
    staged = H / 'supplement-57-staged.xlsx'
    with ZipFile(staged, 'w') as out:
        for info in z.infolist():
            out.writestr(info, replacements.get(info.filename, z.read(info.filename)))
with ZipFile(p) as a, ZipFile(staged) as b:
    assert a.namelist() == b.namelist()
    changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
    assert set(changed) == set(replacements)
print('Staged 11 numeric cells; changed ZIP entries:', changed)
# Small read-only render subset, avoiding the full-workbook import/export memory limit.
with ZipFile(staged) as z:
    wb = E.fromstring(z.read('xl/workbook.xml'))
    sheets = wb.find('m:sheets', ns)
    for x in list(sheets):
        if x.get('name') != 'Import dieu chinh':
            sheets.remove(x)
    sh = E.fromstring(z.read(sheet_path))
    data = sh.find('m:sheetData', ns)
    row = int(d['changes'][0]['cell'][1:])
    for x in list(data):
        if int(x.get('r')) not in [row, row + 1]:
            data.remove(x)
        else:
            for cell in list(x):
                col = ''.join(c for c in cell.get('r') if c.isalpha())
                if col not in ['O', 'P', 'Q', 'R', 'S']:
                    x.remove(cell)
    for t in ['mergeCells', 'dataValidations', 'tableParts', 'autoFilter']:
        elem = sh.find('m:' + t, ns)
        if elem is not None:
            sh.remove(elem)
    dim = sh.find('m:dimension', ns)
    if dim is not None:
        dim.set('ref', f'O{row}:S{row+1}')
    preview_changes = {'xl/workbook.xml': E.tostring(wb), sheet_path: E.tostring(sh)}
    with ZipFile(H / 'supplement-57-preview.xlsx', 'w') as out:
        for info in z.infolist():
            out.writestr(info, preview_changes.get(info.filename, z.read(info.filename)))
