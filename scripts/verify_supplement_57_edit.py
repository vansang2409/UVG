"""Verify staged unit-price edits before replacing the delivered workbook."""
from pathlib import Path
from itertools import zip_longest
import hashlib
import json
import shutil
import openpyxl
from lxml import etree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
H = ROOT / 'scripts/data/check-review'
p = ROOT / 'outputs/HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx'
staged = H / 'supplement-57-staged.xlsx'
d = json.loads((H / 'supplement-57-price-changes.json').read_text(encoding='utf-8'))
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
assert sha(p) == d['base_sha256'], 'Source workbook changed'
allowed = {(x['sheet'], x['cell']): x for x in d['changes']}
a = openpyxl.load_workbook(p, read_only=True, data_only=False)
b = openpyxl.load_workbook(staged, read_only=True, data_only=False)
assert a.sheetnames == b.sheetnames
changed, count = [], 0
for sn in a.sheetnames:
    for ar, br in zip_longest(a[sn].iter_rows(), b[sn].iter_rows(), fillvalue=()):
        for x, y in zip_longest(ar, br):
            ov, nv = getattr(x, 'value', None), getattr(y, 'value', None)
            if ov is None and nv is None:
                continue
            count += 1
            coord = getattr(x, 'coordinate', None) or y.coordinate
            if ov != nv:
                change = allowed[(sn, coord)]
                assert ov == change['old'] and nv == change['new'], (sn, coord)
                changed.append((sn, coord))
            if (sn, coord) in allowed:
                assert y.number_format == '#,##0.0;[Red](#,##0.0);0'
            else:
                assert x.number_format == y.number_format, (sn, coord, 'number_format')
            for attr in ['font', 'fill', 'border', 'alignment', 'protection']:
                assert getattr(x, attr) == getattr(y, attr), (sn, coord, attr)
a.close()
b.close()
assert set(changed) == set(allowed)
# Sheet structure and control settings outside cell data must survive export.
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(p) as za, ZipFile(staged) as zb:
    for name in [n for n in za.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]:
        ea, eb = ET.fromstring(za.read(name)), ET.fromstring(zb.read(name))
        for tag in ['mergeCells', 'sheetViews', 'cols', 'dataValidations', 'autoFilter', 'tableParts']:
            xa, xb = ea.find('m:' + tag, ns), eb.find('m:' + tag, ns)
            def canonical(x):
                if x is None:
                    return None
                return (ET.QName(x).localname, dict(x.attrib), [canonical(c) for c in x])
            assert canonical(xa) == canonical(xb), (name, tag)
        old_rows = [(r.get('r'), r.get('ht'), r.get('customHeight'), r.get('hidden')) for r in ea.findall('m:sheetData/m:row', ns)]
        new_rows = [(r.get('r'), r.get('ht'), r.get('customHeight'), r.get('hidden')) for r in eb.findall('m:sheetData/m:row', ns)]
        assert old_rows == new_rows, (name, 'row dimensions')
result = {'checked_nonempty_cells': count, 'changed_cells': d['changes'], 'only_authorized_cells_changed': True, 'styles_and_structure_preserved': True, 'before_sha256': sha(p), 'after_sha256': sha(staged)}
shutil.copy2(p, H / 'supplement-57-before.xlsx')
shutil.copy2(staged, p)
(H / 'supplement-57-edit-verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'checked_cells': count, 'changed_cells': len(changed), 'replaced_after_verification': True}))
