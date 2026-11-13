"""Đọc lại chi tiết 10 HĐ và 5 BHH; không sửa nguồn hoặc liên kết."""
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from collections import defaultdict,Counter
from decimal import Decimal
import io,json,sys,re,hashlib
sys.stdout.reconfigure(encoding='utf-8')
R=Path('D:/UVG');O=Path(__file__).parent;sys.path.insert(0,str(R/'scripts'))
from prepare_invoice_reports import rows,invoice_key
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def dump(n,v):(O/n).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf-8')
report=read(O/'reverse-report-data.json');cases=[];keys=set();documents=set();wanted_amis=defaultdict(dict)
amis=read(O/'amis.json');mapping={r['id']:r for r in read(O/'mapping.json')};ledger=read(O/'ledger.json')
for r in report['main']:
    a,b,doc=r[0],r[7],r[6];keys.update([a,b]);documents.add(doc)
    m=mapping['2025/'+doc];d=amis['2025/'+doc]
    cases.append({'orphan':a,'kept':b,'doc':doc,'doc_day':m['day'],'amis_money':m['amis_money'],'order':m['order'],'current_mapping':m,'header_orphan':ledger[a],'header_kept':ledger[b]})
    for pos in d['rows']:
        sn,rn=pos.rsplit('!',1);wanted_amis[sn][int(rn)]=doc
headers={};details=defaultdict(list);source_manifest=[]
for p in sorted((R/'data/1311/meinvoice').glob('*.zip')):
    if 'chi_tiet' not in p.name and 'hoa_don_da_su_dung_' not in p.name:continue
    source_manifest.append({'file':str(p.relative_to(R)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
    with ZipFile(p) as z:
        is_detail='chi_tiet' in p.name
        members=z.namelist()
        if not is_detail:
            selected={ledger[k]['source'].rsplit('!',1)[0] for k in keys};members=[x for x in members if x in selected]
        for member in members:
            if member.endswith('/'):continue
            for rn,r in rows(io.BytesIO(z.read(member))):
                if rn<6:continue
                k=invoice_key(r.get('B'),r.get('C'));k='/'.join(k) if k else None
                if k not in keys:continue
                if is_detail:
                    details[k].append({'row':rn,'file':member,'date':r.get('D'),'customer':r.get('E'),'tax':r.get('G'),'buyer':r.get('H'),'kind':r.get('K'),'name':r.get('L'),'unit':r.get('M'),'qty':r.get('N'),'price':r.get('O'),'goods':r.get('P'),'rate':r.get('Q'),'vat':r.get('R'),'payment':r.get('S')})
                else:headers[k]={'row':rn,'file':member,'values':r}
    print('READ',p.name,len(headers),sum(map(len,details.values())),flush=True)
assert set(headers)==keys and set(details)==keys
# Sổ AMIS thật đang dùng, đọc đúng các dòng BHH đã xác định.
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';amis_lines=defaultdict(list)
p=R/'data/1311/amis/So_chi_tiet_ban_hang (2025).xlsx';source_manifest.append({'file':str(p.relative_to(R)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
with ZipFile(p) as z:
    ss=[''.join(t.text or '' for t in s.iter(NS+'t')) for s in etree.fromstring(z.read('xl/sharedStrings.xml'))]
    wb=etree.fromstring(z.read('xl/workbook.xml'));rels={e.get('Id'):e.get('Target').lstrip('/') for e in etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    for sh in wb.find(NS+'sheets'):
        name=sh.get('name')
        if name not in wanted_amis:continue
        f=rels[sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')];f=f if f.startswith('xl/') else 'xl/'+f
        with z.open(f) as stream:
            for _,row in etree.iterparse(stream,events=('end',),tag=NS+'row'):
                rn=int(row.get('r'))
                if rn in wanted_amis[name]:
                    rr={}
                    for cell in row:
                        c=re.sub(r'\d','',cell.get('r'));v=cell.find(NS+'v')
                        if cell.get('t')=='s':value=ss[int(v.text)] if v is not None else None
                        elif cell.get('t')=='inlineStr':value=''.join(t.text or '' for t in cell.iter(NS+'t'))
                        else:value=v.text if v is not None else None
                        rr[c]=value
                    doc=wanted_amis[name][rn];assert rr['C']==doc
                    amis_lines[doc].append({'sheet':name,'row':rn,'values':rr})
                row.clear()
                while row.getprevious() is not None:del row.getparent()[0]
kt=[]
for p in sorted((R/'outputs/import-2025-20261109/source').glob('kt_*.json')):
    d=read(p);start=int(p.stem.split('_')[1])
    for i,r in enumerate(d['values']):
        if len(r)>10:
            k=invoice_key(r[9],r[10]);k='/'.join(k) if k else None
            if k in keys:kt.append({'key':k,'row':start+i,'values':r})
nb=read(R/'outputs/phan-loai-2025-20261105/source/plan25NB.json');orders={x for c in cases for x in c['order']};nb_matches=[]
for rn,r in enumerate(nb,2):
    if len(r)>8 and (str(r[8]) in orders):nb_matches.append({'row':rn,'values':r})
for c in cases:
    c['raw_headers']={k:headers[k] for k in [c['orphan'],c['kept']]};c['invoice_lines']={k:details[k] for k in [c['orphan'],c['kept']]};c['amis_lines']=amis_lines[c['doc']]
    c['cached_kt']=[r for r in kt if r['key'] in [c['orphan'],c['kept']]];c['cached_nb']=[r for r in nb_matches if str(r['values'][8]) in c['order']]
dump('five-invoice-deep-evidence.json',{'cases':cases,'sources':source_manifest,'kt_rows':[r['row'] for r in kt],'nb_rows':[r['row'] for r in nb_matches]})
print('SUMMARY',json.dumps([{'orphan':c['orphan'],'kept':c['kept'],'doc':c['doc'],'day':c['doc_day'],'order':c['order'],'orphan_header':{k:c['raw_headers'][c['orphan']]['values'].get(k) for k in ['D','F','H','J','L','P','Q','R','U']},'kept_header':{k:c['raw_headers'][c['kept']]['values'].get(k) for k in ['D','F','H','J','L','P','Q','R','U']},'orphan_items':c['invoice_lines'][c['orphan']],'kept_items':c['invoice_lines'][c['kept']],'amis_items':[{k:x['values'].get(k) for k in ['C','P','Q','W','X','Y','AP','AR']} for x in c['amis_lines']],'nb':c['cached_nb']} for c in cases],ensure_ascii=False),flush=True)
