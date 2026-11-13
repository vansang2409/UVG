"""Đọc dữ liệu nghiệm thu mới, giữ nguyên file nguồn và mã định danh."""
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from collections import defaultdict, Counter
from decimal import Decimal
from datetime import date, datetime, timedelta
import json, io, re, hashlib, sys
sys.stdout.reconfigure(encoding='utf-8')
ROOT=Path('D:/UVG'); SRC=ROOT/'data/1311'; OUT=Path(__file__).parent
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def txt(x): return '' if x is None else str(x).strip()
def num(x): return None if x in (None,'') else Decimal(str(x))
def native(x):
    if isinstance(x,Decimal): return int(x) if x==int(x) else float(x)
    if isinstance(x,set): return sorted(x)
    raise TypeError(type(x).__name__)
def dump(n,x): (OUT/n).write_text(json.dumps(x,ensure_ascii=False,separators=(',',':'),default=native),encoding='utf-8')
def day(v):
    if not txt(v): return None
    for f in ('%d/%m/%Y','%Y-%m-%d'):
        try: return datetime.strptime(txt(v),f).date().isoformat()
        except ValueError: pass
    try:return (datetime(1899,12,30)+timedelta(days=float(v))).date().isoformat()
    except (ValueError,OverflowError): return None
def key(s,n):
    s,n=txt(s).upper(),txt(n)
    if n.endswith('.0'): n=n[:-2]
    return s+'/'+n.zfill(8) if s and n.isdigit() else None
def sheets(content):
    with ZipFile(content) as z:
        ss=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            ss=[''.join(e.itertext()) for e in etree.fromstring(z.read('xl/sharedStrings.xml'))]
        wb=etree.fromstring(z.read('xl/workbook.xml'))
        rel=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets={e.get('Id'):e.get('Target').lstrip('/') for e in rel}
        for sh in wb.find(NS+'sheets'):
            target=targets[sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
            if not target.startswith('xl/'): target='xl/'+target
            def gen(target=target):
                with z.open(target) as f:
                    for _,e in etree.iterparse(f,events=('end',),tag=NS+'row'):
                        r={}
                        for c in e:
                            col=re.sub(r'\d','',c.get('r',''));v=c.find(NS+'v');t=c.get('t')
                            if t=='s' and v is not None: val=ss[int(v.text)]
                            elif t=='inlineStr':val=''.join(x.text or '' for x in c.iter(NS+'t'))
                            else:val=v.text if v is not None else None
                            r[col]=val
                        yield int(e.get('r')),r
                        e.clear()
                        while e.getprevious() is not None:del e.getparent()[0]
            yield sh.get('name'),gen()
manifest=[]; docs={}; amis_controls=[]
for p in sorted((SRC/'amis').glob('*.xlsx')):
    manifest.append({'file':str(p.relative_to(ROOT)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size})
    for sheet,rr in sheets(p):
        counts=0; sums=[Decimal(0)]*3;foot=[];headers={}
        for rn,r in rr:
            if rn==4: headers={txt(v):k for k,v in r.items() if v}
            if rn<5:continue
            get=lambda label:r.get(headers[label])
            doc=txt(get('Số chứng từ'));dt=day(get('Ngày hạch toán'))
            if not doc or not dt:
                if any('Tổng cộng' in txt(v) for v in r.values()):foot.append({'row':rn,'values':{k:v for k,v in r.items() if v}})
                continue
            year=int(dt[:4]);dk=str(year)+'/'+doc
            d=docs.setdefault(dk,{'year':year,'doc':doc,'dates':set(),'invoice_dates':set(),'customer':txt(get('Tên khách hàng')),'memo':txt(get('Diễn giải chung')),'keys':set(),'partial_keys':set(),'goods':Decimal(0),'vat':Decimal(0),'pay':Decimal(0),'missing':Counter(),'rows':[],'lines':0,'source':str(p.relative_to(ROOT))})
            d['dates'].add(dt);d['invoice_dates'].add(day(get('Ngày hóa đơn')) or '')
            sig,inv=get('Ký hiệu hóa đơn'),get('Số hóa đơn');k=key(sig,inv)
            if k:d['keys'].add(k)
            elif txt(sig) or txt(inv):d['partial_keys'].add(txt(sig)+'/'+txt(inv))
            amounts={l:num(get(l)) for l in ['Doanh số bán','Chiết khấu','Giá trị trả lại','Giá trị giảm giá','Thuế GTGT','Tổng thanh toán']}
            values=[None if any(amounts[l] is None for l in ['Doanh số bán','Chiết khấu','Giá trị trả lại','Giá trị giảm giá']) else amounts['Doanh số bán']-amounts['Chiết khấu']-amounts['Giá trị trả lại']-amounts['Giá trị giảm giá'],amounts['Thuế GTGT'],amounts['Tổng thanh toán']]
            for i,f in enumerate(['goods','vat','pay']):
                if values[i] is None:d['missing'][f]+=1
                else:d[f]+=values[i];sums[i]+=values[i]
            d['rows'].append(sheet+'!'+str(rn));d['lines']+=1;counts+=1
        amis_controls.append({'file':p.name,'sheet':sheet,'lines':counts,'totals':sums,'footers':foot})
        print('AMIS',p.name,sheet,counts,flush=True)
dump('amis.json',docs);dump('amis-controls.json',amis_controls)
ledger={}; detail={}; edges={}; file_controls=[]; issues=[]
for p in sorted((SRC/'meinvoice').glob('*.zip')):
    manifest.append({'file':str(p.relative_to(ROOT)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size})
    kind='detail' if 'chi_tiet' in p.name else 'adjustment' if 'dieu_chinh' in p.name else 'replacement' if 'thay_the' in p.name else 'header'
    with ZipFile(p) as z:
        assert z.testzip() is None
        for member in z.infolist():
            if member.is_dir():continue
            data=z.read(member)
            manifest.append({'file':str(p.relative_to(ROOT)),'member':member.filename,'sha256':hashlib.sha256(data).hexdigest(),'bytes':len(data)})
            for sheet,rr in sheets(io.BytesIO(data)):
                count=0;period=None;foot=[]
                for rn,r in rr:
                    if rn==3:period=r.get('A')
                    if rn<6 or not txt(r.get('A')).isdigit():
                        if rn>6 and any('Tổng' in txt(v) for v in r.values()):foot.append({'row':rn,'values':r})
                        continue
                    source=member.filename+'!'+str(rn)
                    if kind=='header':
                        k=key(r.get('B'),r.get('C'));assert k
                        rec={'key':k,'day':day(r.get('D')),'order':txt(r.get('F')),'buyer':txt(r.get('H') or r.get('L')),'money':[num(r.get(c)) for c in ['P','Q','R']],'status':txt(r.get('U')),'cqt_status':txt(r.get('S')),'source':source}
                        if k in ledger: issues.append({'kind':'duplicate_header','key':k,'old':ledger[k],'new':rec})
                        else:ledger[k]=rec
                    elif kind=='detail':
                        k=key(r.get('B'),r.get('C'));assert k
                        d=detail.setdefault(k,{'lines':0,'sums':[Decimal(0)]*3,'known':[0]*3,'source_first':source,'days':set()})
                        for i,c in enumerate(['P','R','S']):
                            v=num(r.get(c))
                            if v is not None:d['sums'][i]+=v;d['known'][i]+=1
                        d['lines']+=1;d['days'].add(day(r.get('D')))
                    else:
                        a,b=key(r.get('B'),r.get('D')),key(r.get('M'),r.get('O'));assert a and b
                        e={'old':a,'new':b,'kind':kind,'old_money':[num(r.get(c)) for c in ['J','K','L']],'new_money':[num(r.get(c)) for c in ['U','V','W']],'reason':txt(r.get('X')),'source':source}
                        ek=a+'>'+b+'>'+kind
                        if ek in edges and edges[ek]!=e:issues.append({'kind':'duplicate_edge','edge':e})
                        edges[ek]=e
                    count+=1
                file_controls.append({'member':member.filename,'sheet':sheet,'kind':kind,'rows':count,'period':period,'footers':foot})
        print('MEINVOICE',kind,'ledger',len(ledger),'detail',len(detail),'edges',len(edges),flush=True)
dump('ledger.json',ledger);dump('detail.json',detail);dump('edges.json',list(edges.values()));dump('manifest.json',manifest);dump('meinvoice-controls.json',file_controls);dump('source-issues.json',issues)
print(json.dumps({'docs':len(docs),'amis_lines':sum(d['lines'] for d in docs.values()),'direct_docs':sum(bool(d['keys']) for d in docs.values()),'partial_docs':sum(bool(d['partial_keys']) for d in docs.values()),'ledger':len(ledger),'details':len(detail),'edges':len(edges),'issues':len(issues),'doc_prefixes':dict(Counter(str(d['year'])+'/'+re.sub(r'\d.*','',d['doc']) for d in docs.values()))},ensure_ascii=False),flush=True)
