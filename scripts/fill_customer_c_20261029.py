"""Chỉ bổ sung ô C trống; giữ nguyên byte mọi nội dung Excel ngoài các ô đích."""
import io, json, re, hashlib, posixpath
from pathlib import Path
from zipfile import ZipFile
from xml.sax.saxutils import escape
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/fill-customer-c-20261029'
report = json.loads((ROOT / 'outputs/missing-customer-untreated-20261029/audit.json').read_text(encoding='utf-8'))
for path, digest in report['source_hashes'].items():
    assert hashlib.sha256((ROOT / path).read_bytes()).hexdigest() == digest
plan = {}
for x in report['records']:
    if x['root_class'] != 'customer_only':
        continue
    names = {e['customer'] for e in x['root_evidence']}
    assert len(names) == 1
    name = next(iter(names))
    for row in x['rows']:
        assert row not in plan
        plan[row] = dict(name=name, seq=x['seq'], key=x['root'])
assert len(plan) == 3913
before = OUT / 'untreated-before.xlsx'
after = OUT / 'untreated-after.xlsx'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with ZipFile(before) as z:
    wb = etree.fromstring(z.read('xl/workbook.xml'))
    sh = next(s for s in wb.find(NS+'sheets') if s.get('name') == 'Import dieu chinh')
    rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    rel = next(r for r in etree.fromstring(z.read('xl/_rels/workbook.xml.rels')) if r.get('Id') == rid)
    target = rel.get('Target')
    target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        strings = [''.join(t.text or '' for t in s.iter(NS+'t')) for s in etree.fromstring(z.read('xl/sharedStrings.xml'))]
    xml = z.read(target)
    tree = etree.fromstring(xml)
    cells = {c.get('r'):c for c in tree.iter(NS+'c')}
    def value(c):
        if c is None: return ''
        v = c.find(NS+'v')
        if c.get('t') == 's' and v is not None: return strings[int(v.text)]
        if c.get('t') == 'inlineStr': return ''.join(t.text or '' for t in c.iter(NS+'t'))
        return v.text if v is not None and v.text else ''
    changes = []
    matches_by_row = {int(m.group(1)):m for m in re.finditer(rb'<c\b(?=[^>]*\br="C(\d+)")(?:[^>]*?/>|[^>]*>.*?</c>)',xml,re.S)}
    replacements=[]
    for row,p in sorted(plan.items()):
        addr = 'C'+str(row)
        c = cells.get(addr)
        assert c is not None, addr
        assert not value(c).strip() and c.find(NS+'f') is None, addr
        assert value(cells.get('A'+str(row))).removesuffix('.0') == p['seq'].removesuffix('.0'), addr
        assert value(cells.get('J'+str(row))) == p['key'][0], addr
        assert value(cells.get('K'+str(row))).removesuffix('.0').zfill(8) == p['key'][1], addr
        match = matches_by_row[row]
        old = match.group()
        opening = old.split(b'>',1)[0].rstrip(b'/')
        opening = re.sub(rb'\s+t="[^"]*"', b'', opening)
        new = opening+b' t="inlineStr"><is><t>'+escape(p['name']).encode('utf-8')+b'</t></is></c>'
        changes.append(dict(cell=addr, name=p['name'], old=old.decode(), new=new.decode()))
        replacements.append((match.start(),match.end(),new))
    # Ghép ngược theo vị trí để giữ nguyên byte mọi phần không được phép thay đổi.
    revised=xml
    for start,end,new in sorted(replacements,reverse=True): revised=revised[:start]+new+revised[end:]
    with ZipFile(after,'w') as dest:
        for info in z.infolist(): dest.writestr(info,revised if info.filename==target else z.read(info.filename))
with ZipFile(before) as a, ZipFile(after) as b:
    assert a.namelist()==b.namelist()
    changed_entries=[n for n in a.namelist() if a.read(n)!=b.read(n)]
    assert changed_entries==[target],changed_entries
    original_cells={c.get('r'):c for c in etree.fromstring(a.read(target)).iter(NS+'c')}
    result_cells={c.get('r'):c for c in etree.fromstring(b.read(target)).iter(NS+'c')}
    assert original_cells.keys()==result_cells.keys()
    differences=[]
    for addr,c in original_cells.items():
        other=result_cells[addr]
        if etree.tostring(c)!=etree.tostring(other):
            differences.append(addr)
            assert addr.startswith('C') and int(addr[1:]) in plan
            assert value(other)==plan[int(addr[1:])]['name']
            assert {k:v for k,v in c.attrib.items() if k!='t'}=={k:v for k,v in other.attrib.items() if k!='t'}
    assert len(differences)==3913
audit=dict(changed_cells=len(changes),changed_zip_entries=changed_entries,verified_cells=len(original_cells),other_zip_entries_unchanged=True,all_other_cells_unchanged=True,before_sha256=hashlib.sha256(before.read_bytes()).hexdigest(),after_sha256=hashlib.sha256(after.read_bytes()).hexdigest(),changes=changes)
(OUT/'untreated-verification.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in audit.items() if k!='changes'},ensure_ascii=False))
