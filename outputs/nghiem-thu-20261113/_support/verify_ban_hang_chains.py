"""Đọc độc lập file đã lưu, kiểm tiền/công thức và lấy mẫu định dạng thật để xem."""
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from collections import defaultdict,Counter
from decimal import Decimal
from datetime import datetime,timedelta
from openpyxl import load_workbook
import json,re,hashlib,sys,unicodedata

sys.stdout.reconfigure(encoding='utf-8');O=Path(__file__).resolve().parent
N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CELL=re.compile(rb'<c\b[^>]*?\br="([A-Z]+)(\d+)"[^>]*?(?:/>|>.*?</c>)',re.S)
RX=re.compile(r'1[A-Z0-9]+/\d{8}')
def read(n):return json.loads((O/n).read_text(encoding='utf-8'))
def norm(s):return ''.join(c for c in unicodedata.normalize('NFD',str(s or '').lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
def val(c,strings):
    if c is None:return None
    if c.get('t')=='s':return strings[int(c.findtext(N+'v'))]
    if c.get('t')=='inlineStr':return ''.join(c.find(N+'is').itertext())
    value=c.findtext(N+'v')
    return Decimal(value) if value is not None else None
def cs(row):return {re.sub(r'\d','',c.get('r')):c for c in row}
def scan(z,part):
    with z.open(part) as stream:
        for _,row in etree.iterparse(stream,events=('end',),tag=N+'row'):
            yield row
            row.clear()
            while row.getprevious() is not None:del row.getparent()[0]
def digest(p):
    with p.open('rb') as s:return hashlib.file_digest(s,'sha256').hexdigest()

report=read('ban-hang-chains-verification.json');assert report['passed']
ledger=read('ledger.json');mapping={r['id']:r for r in read('mapping.json')}
reverse=read('reverse-chain-audit.json');edges=read('edges.json')+reverse['recovered_edges']
resolved={i['key']:i['resolved'] for c in reverse['components'] for i in c['invoices']};del reverse
replaced={e['old'] for e in edges if e['kind']=='replacement'};parent={e['new']:e['old'] for e in edges};children=defaultdict(set)
kind={e['new']:e['kind'] for e in edges}
for e in edges:children[e['old']].add(e['new'])
views=[];results=[]
for rec in report['files']:
    path=Path(rec['file']);year=rec['year'];assert digest(path)==rec['after_sha256']
    backup=Path(rec['backup'])
    if not backup.exists():backup=O/backup.parent.name/backup.name
    with ZipFile(path) as z,ZipFile(backup) as before:
        strings=[''.join(si.itertext()) for si in etree.fromstring(z.read('xl/sharedStrings.xml'))]
        workbook=etree.fromstring(z.read('xl/workbook.xml'));names=[s.get('name') for s in workbook.find(N+'sheets')]
        assert names==rec['original_sheets']+[rec['detail_sheet']]
        assert z.testzip() is None
        oldstyles=etree.fromstring(before.read('xl/styles.xml'));newstyles=etree.fromstring(z.read('xl/styles.xml'))
        for group in ['fonts','fills','borders','cellXfs','numFmts']:
            a=oldstyles.find(N+group);b=newstyles.find(N+group)
            for i,node in enumerate(a):assert etree.tostring(node)==etree.tostring(b[i]),group
        originals={};main={};examples=[];member_by_sheet={};row_samples={};sample_types=set();type_counts=Counter()
        for index,name in enumerate(rec['original_sheets'],1):
            part=f'xl/worksheets/sheet{index}.xml';raw=z.read(part);original=before.read(part)
            assert [m[0] for m in CELL.finditer(original)]==[m[0] for m in CELL.finditer(raw) if len(m[1])==1 and m[1]<=b'L']
            root=etree.fromstring(raw)
            assert root.find(N+'autoFilter').get('ref').startswith('A3:AA') and root.find(N+'sheetViews/'+N+'sheetView/'+N+'pane') is not None
            member_by_sheet[name]=part
            for row in root.find(N+'sheetData'):
                rn=int(row.get('r'));cells=cs(row);doc=val(cells.get('C'),strings)
                if rn<4 or not doc:continue
                dk=str(year)+'/'+doc;assert dk not in main
                phases=[RX.findall(val(cells.get(c),strings) or '') for c in ['N','Q','T','W']]
                for c,ks in zip(['N','Q','T','W'],phases):assert (val(cells.get(c),strings) or '')=='\n'.join(ks),'Số HĐ còn nhãn loại/trạng thái'
                keys=[k for phase in phases for k in phase];assert len(keys)==len(set(keys))
                assert set(mapping[dk]['keys'])<=set(keys)
                assert all(children[k]<=set(keys) for k in keys)
                for phase,ks in enumerate(phases):
                    for k in ks:
                        if phase==0:assert k not in parent
                        elif phase==1:assert k in parent and ledger[k]['day'][:4]<'2026'
                        elif phase==2:assert k in parent and ledger[k]['day'].startswith('2026-')
                        else:assert k not in parent and ledger[k]['day'].startswith('2026-')
                has_tt=any(k in replaced or kind.get(k)=='replacement' or 'thay the' in norm(ledger[k]['status']) for k in keys)
                has_dc=any(kind.get(k)=='adjustment' or 'dieu chinh' in norm(ledger[k]['status']) for k in keys)
                category='Thay thế' if has_tt else 'Điều chỉnh' if has_dc else 'HĐ mới'
                assert val(cells['M'],strings)==category
                if has_tt and has_dc:assert 'cả điều chỉnh và thay thế' in val(cells['AA'],strings)
                type_counts[category]+=1
                main[dk]={'keys':set(keys),'phases':phases,'final':val(cells.get('Z'),strings),'formula':cells['Z'].findtext(N+'f'),
                          'phase_values':[val(cells.get(c),strings) for c in ['P','S','V','Y']],
                          'phase_formulas':[cells[c].findtext(N+'f') for c in ['P','S','V','Y']]}
                wanted=category not in sample_types
                if wanted:examples.append((name,rn,doc));sample_types.add(category);row_samples[(name,rn)]=etree.tostring(row)
            row_samples[(name,2)]=etree.tostring(root.find(N+'sheetData/'+N+'row[@r="2"]'))
            row_samples[(name,3)]=etree.tostring(root.find(N+'sheetData/'+N+'row[@r="3"]'))
            del root,raw,original
        totals=defaultdict(Decimal);phase_totals=defaultdict(lambda:defaultdict(Decimal));actual=defaultdict(lambda:defaultdict(list));ranges=defaultdict(dict);seen=set();detail_count=0;formulas=0
        for row in scan(z,rec['detail_member']):
            rn=int(row.get('r'))
            if rn<=3:continue
            cells=cs(row);doc=val(cells['A'],strings);dk=str(year)+'/'+doc;phase=int(val(cells['B'],strings)[1:])
            key=val(cells['C'],strings)+'/'+val(cells['D'],strings)
            assert (dk,key) not in seen;seen.add((dk,key));x=ledger[key]
            assert cells['D'].get('t')=='inlineStr' and len(val(cells['D'],strings))==8
            assert int(val(cells['E'],strings))==(datetime.fromisoformat(x['day'])-datetime(1899,12,30)).days
            raw_money=[val(cells.get(c),strings) for c in ['F','G','H']]
            assert raw_money==[None if v is None else Decimal(str(v)) for v in x['money']]
            g,v,p=resolved.get(key,x['money']);pay=x['money'][2]
            if pay is None:pay=p if p is not None else g+v if g is not None and v is not None else None
            pay=Decimal(str(pay));assert val(cells['I'],strings)==pay
            active=key not in replaced and 'bi thay the' not in norm(x['status']) and 'huy' not in norm(x['status']) and 'khong hop le' not in norm(x['cqt_status'])
            value=pay if active else Decimal(0)
            assert val(cells['K'],strings)==value
            assert cells['K'].findtext(N+'f')==f'IF(J{rn}="Còn hiệu lực",I{rn},0)'
            assert val(cells['J'],strings)==('Còn hiệu lực' if active else 'Không cộng')
            totals[dk]+=value;phase_totals[dk][phase]+=pay;actual[dk][phase].append(key)
            span=ranges[dk].setdefault(phase,[rn,rn]);span[1]=rn;detail_count+=1;formulas+=1
        assert set(main)==set(actual) and len(main)==rec['documents']
        for dk,r in main.items():
            assert r['final']==totals[dk],(dk,r['final'],totals[dk])
            first=min(span[0] for span in ranges[dk].values());last=max(span[1] for span in ranges[dk].values())
            assert r['formula']==f"SUM('{rec['detail_sheet']}'!K{first}:K{last})"
            for phase in range(4):
                assert r['phases'][phase]==actual[dk][phase]
                if r['phases'][phase]:
                    first,last=ranges[dk][phase];assert r['phase_values'][phase]==phase_totals[dk][phase]
                    assert r['phase_formulas'][phase]==f"SUM('{rec['detail_sheet']}'!I{first}:I{last})"
                else:assert r['phase_values'][phase] is None and r['phase_formulas'][phase] is None
        if year==2025:
            assert totals['2025/BHHS2025017410']==2639000 and totals['2025/BHHS2025017546']==1479000
        assert detail_count==rec['invoice_memberships']
        # Đọc bằng openpyxl một vùng thật để kiểm định dạng Excel và công thức cached.
        w=load_workbook(path,read_only=True,data_only=True)
        s=w[examples[0][0]];rn=examples[0][1];row=next(s.iter_rows(min_row=rn,max_row=rn,max_col=27))
        assert isinstance(row[25].value,(int,float)) and row[25].font.bold
        w.close()
        fills=newstyles.find(N+'fills');fonts=newstyles.find(N+'fonts');xfs=newstyles.find(N+'cellXfs');formats={int(n.get('numFmtId')):n.get('formatCode') for n in newstyles.find(N+'numFmts')}
        for label,columns in [('F0 F1',[3,13,14,15,16,17,18,19]),('F2 F3 cuối',[3,13,9,20,21,22,23,24,25,26,27])]:
            sheet=examples[0][0];chosen=[e for e in examples if e[0]==sheet][:3];source_rows=[2,3]+[e[1] for e in chosen]
            rows=[];heights=[]
            for rn in source_rows:
                node=etree.fromstring(row_samples[(sheet,rn)]);cells=cs(node);output=[]
                heights.append(float(node.get('ht','24')))
                for c in columns:
                    address='';cc=c
                    while cc:cc,d=divmod(cc-1,26);address=chr(65+d)+address
                    cell=cells.get(address)
                    if cell is None:output.append({'value':None});continue
                    xf=xfs[int(cell.get('s','0'))];font=fonts[int(xf.get('fontId','0'))];fill=fills[int(xf.get('fillId','0'))];alignment=xf.find(N+'alignment');fmt=formats.get(int(xf.get('numFmtId','0')),'General')
                    value=val(cell,strings);isdate='dd/mm/yyyy'==fmt and isinstance(value,Decimal)
                    if isdate:value=(datetime(1899,12,30)+timedelta(days=int(value))).isoformat()
                    elif isinstance(value,Decimal):value=int(value) if value==int(value) else float(value)
                    fg=fill.find(N+'patternFill/'+N+'fgColor');color=font.find(N+'color')
                    output.append({'value':value,'isDate':isdate,'font':{'name':font.find(N+'name').get('val'),'size':float(font.find(N+'sz').get('val')),'bold':font.find(N+'b') is not None,'italic':False,'color':'#'+color.get('rgb')[-6:] if color is not None and color.get('rgb') else '#233348'},'fill':'#'+fg.get('rgb')[-6:] if fg is not None and fg.get('rgb') else None,'format':fmt,'wrap':True,'align':alignment.get('horizontal','left') if alignment is not None else 'left','borders':{}})
                rows.append(output)
            widths_by_col={3:21,9:22,13:22,**{14+i:w for i,w in enumerate([34,15,22]*4+[23,65])}}
            merges=[]
            for phase in range(4):
                block=[14+phase*3,15+phase*3,16+phase*3]
                if set(block)<=set(columns):
                    start=columns.index(block[0])+1;merges.append(chr(64+start)+'1:'+chr(66+start)+'1')
            if 26 in columns and 27 in columns:merges.append(chr(65+columns.index(26))+'1:'+chr(65+columns.index(27))+'1')
            views.append({'name':str(year)+' '+label,'source':sheet,'rows':rows,'heights':heights,'widths':[widths_by_col[c] for c in columns],'source_columns':columns,'source_rows':source_rows,'merges':merges})
        assert dict(type_counts)==rec['chain_type_counts']
        results.append({'year':year,'documents':len(main),'invoice_memberships':detail_count,'original_A_L_cells_identical':True,'all_amounts_dates_invoice_ids_formulas_verified':True,'chain_type_column_verified':True,'chain_type_counts':dict(type_counts),'invoice_ids_without_annotations_verified':True,'unknown_final':0,'sha256':digest(path)})
        print('VERIFIED',year,len(main),detail_count,flush=True)
(O/'ban-hang-preview-cells.json').write_text(json.dumps(views,ensure_ascii=False),encoding='utf-8')
(O/'ban-hang-chains-independent-verification.json').write_text(json.dumps({'passed':True,'files':results},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'passed':True,'documents':sum(r['documents'] for r in results),'files':2},ensure_ascii=False),flush=True)
