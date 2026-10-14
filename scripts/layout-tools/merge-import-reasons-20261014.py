from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
from copy import deepcopy
import json, re, sys

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'scripts/data/import-reasons-20261014'
plan = json.loads((OUT / 'plan.json').read_text(encoding='utf-8'))
idx = int(sys.argv[1]); f = plan['files'][idx]
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
tag = lambda x: '{' + NS + '}' + x
with ZipFile(OUT / 'before' / f['name']) as original, ZipFile(OUT / f'authored-{idx}.xlsx') as authored:
    wb = E.fromstring(original.read('xl/workbook.xml'))
    rels = E.fromstring(original.read('xl/_rels/workbook.xml.rels'))
    rid = next(s for s in wb.find(tag('sheets')) if s.get('name') == 'Import dieu chinh').get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    target = next(r.get('Target') for r in rels if r.get('Id') == rid)
    target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
    source = E.fromstring(authored.read('xl/worksheets/sheet1.xml'))
    newcells = {c.get('r'): c for c in source.iter(tag('c'))}
    oldsst = E.fromstring(original.read('xl/sharedStrings.xml'))
    author_sst = E.fromstring(authored.read('xl/sharedStrings.xml'))
    offset = len(oldsst)
    for si in author_sst:
        oldsst.append(deepcopy(si))
    oldsst.set('uniqueCount', str(len(oldsst)))
    oldsst.attrib.pop('count', None)
    raw = original.read(target)
    expected = {'N' + str(c['row']) for c in f['changes']}
    done = set()
    def replace(m):
        cell = E.fromstring(m.group(0), parser=E.XMLParser(recover=False)) if b'xmlns' in m.group(0) else E.fromstring(b'<root xmlns="' + NS.encode() + b'">' + m.group(0) + b'</root>')[0]
        coord = cell.get('r')
        if coord not in expected:
            return m.group(0)
        generated = deepcopy(newcells[coord])
        assert generated.get('t') in ['s', 'str', 'inlineStr'], (coord, generated.get('t'))
        if generated.get('t') == 's':
            generated.find(tag('v')).text = str(int(generated.find(tag('v')).text) + offset)
        cell.set('t', generated.get('t'))
        for ch in list(cell): cell.remove(ch)
        for ch in generated: cell.append(deepcopy(ch))
        done.add(coord)
        return E.tostring(cell, encoding='utf-8')
    raw = re.sub(rb'<c\b(?=[^>]*\br="N\d+")[^>]*?(?:/>|>.*?</c>)', replace, raw, flags=re.S)
    assert done == expected, (len(done), len(expected))
    modified = {target: raw, 'xl/sharedStrings.xml': E.tostring(oldsst, xml_declaration=True, encoding='utf-8', standalone=True)}
    with ZipFile(OUT / 'staged' / f['name'], 'w') as dest:
        for item in original.infolist(): dest.writestr(item, modified.get(item.filename, original.read(item.filename)))
    print(f['name'], 'authored cells merged', len(done), 'all other package entries preserved')
