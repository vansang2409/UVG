from pathlib import Path
import zipfile, xml.etree.ElementTree as ET, json, unicodedata, re, collections

ROOT = Path(r'C:\Users\thuong ngo\Downloads\ĐỐI SOÁT CẦN GỘP')
OUT = Path(__file__).parent
NS = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
def norm(x):
    return re.sub(r'\s+', ' ', unicodedata.normalize('NFC', str(x or '')).strip()).casefold()
def plain(x):
    return ''.join(c for c in unicodedata.normalize('NFD', norm(x).replace('đ','d')) if unicodedata.category(c) != 'Mn')
def is_target(x):
    p = plain(x)
    return p in {'hoa don 1 san pham','hoa don nhieu san pham','don hang 1sp','don hang nhieu sp','chi tiet xuat hoa don 1 san pha','chi tiet xuat hoa don nhieu san'}
def book_info(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rel = {e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    sheets=[]
    for e in wb.find('m:sheets',NS):
        t=rel[e.attrib[RNS]]
        sheets.append((e.attrib['name'], t.lstrip('/') if t.startswith('/') else 'xl/'+t))
    return sheets
def strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist(): return []
    return [''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml'))]
def cell_value(c, ss):
    v=c.find('m:v',NS)
    if c.attrib.get('t')=='s': return ss[int(v.text)] if v is not None else None
    if c.attrib.get('t')=='inlineStr': return ''.join(c.find('m:is',NS).itertext())
    return v.text if v is not None else None

def main():
    counts=collections.Counter(); selected=[]; missing=[]; signatures=collections.defaultdict(list)
    files=sorted(p for p in ROOT.rglob('*.xlsx') if not p.name.startswith('~$'))
    for i,p in enumerate(files):
        with zipfile.ZipFile(p) as z:
            sheets=book_info(z)
            counts.update(n for n,_ in sheets)
            targets=[(n,t) for n,t in sheets if is_target(n)]
            if len(targets)!=2: missing.append({'file':str(p.relative_to(ROOT)), 'sheets':[n for n,_ in sheets], 'targets':[n for n,_ in targets]})
            if targets:
                ss=strings(z)
                for n,t in targets:
                    rows=[]; dimension=None
                    with z.open(t) as f:
                        for ev,e in ET.iterparse(f,events=('end',)):
                            if e.tag.endswith('}dimension'): dimension=e.attrib.get('ref')
                            if e.tag.endswith('}row'):
                                vals={c.attrib['r']:cell_value(c,ss) for c in e if c.tag.endswith('}c')}
                                rows.append(vals)
                                if len(rows)>=8: break
                    rec={'file':str(p.relative_to(ROOT)), 'sheet':n, 'xml':t, 'dimension':dimension, 'rows':rows}
                    selected.append(rec)
                    signature=json.dumps(rows[:2],ensure_ascii=False,sort_keys=True)
                    signatures[signature].append(rec)
        if (i+1)%50==0: print(f'Inspected {i+1}/{len(files)}',flush=True)
    report={'files':len(files),'sheet_names':dict(counts),'missing':missing,'selected':selected}
    (OUT/'inspection.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    headers=collections.defaultdict(list)
    for rec in selected:
        hs=next((r for r in rec['rows'] if any(plain(v)=='ma don hang' for v in r.values())),{})
        headers[json.dumps(hs,ensure_ascii=False)].append(rec)
    print(json.dumps({'files':len(files),'target_sheets':len(selected),'missing':missing,'header_variants':[{'count':len(v),'example':v[0]['file'],'sheet':v[0]['sheet'],'headers':json.loads(k)} for k,v in headers.items()]},ensure_ascii=False,indent=2))

if __name__=='__main__': main()
