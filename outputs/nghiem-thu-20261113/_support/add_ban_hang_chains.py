"""Bổ sung chuỗi vào danh sách bán hàng, bảo toàn nguyên ô A:L bằng OOXML.

Fallback giữ byte nguồn: Artifact Tool đã hết bộ nhớ ở bộ báo cáo lớn cùng đợt.
Không nạp/xuất lại toàn workbook qua trình đọc có thể đổi dữ liệu AMIS.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from collections import defaultdict, Counter
from functools import lru_cache
from datetime import datetime
from decimal import Decimal
from xml.sax.saxutils import escape
from lxml import etree
import hashlib, json, re, shutil, os, sys, io

sys.stdout.reconfigure(encoding='utf-8')
OUT = Path(__file__).resolve().parent
SRC = (OUT.parents[2] if OUT.name=='_support' else OUT.parents[1]) / 'data/1311/amis'
BACKUP = OUT / 'backup-ban-hang-before-chains-20261113'
REPORT = OUT / 'ban-hang-chains-verification.json'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N = '{' + NS + '}'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
DETAIL = 'Chi tiết chuỗi HĐ'
ROW = re.compile(rb'<row\b[^>]*?\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
CELL = re.compile(rb'<c\b[^>]*?\br="([A-Z]+)(\d+)"[^>]*?(?:/>|>.*?</c>)', re.S)
PALETTE = [('#305E8C','#EAF2FA'),('#9A6700','#FFF5DF'),('#6B4C9A','#F4EEFA'),('#2F6F46','#EAF5EC')]
TITLES = ['F0 - HĐ gốc', 'F1 - ĐC/TT trước 2026', 'F2 - Xử lý năm 2026', 'F3 - Xuất mới / xuất lại 2026']

def read(name): return json.loads((OUT / name).read_text(encoding='utf-8'))
def sha(path):
    with path.open('rb') as stream: return hashlib.file_digest(stream, 'sha256').hexdigest()
def norm(value):
    import unicodedata
    return ''.join(c for c in unicodedata.normalize('NFD',str(value or '').lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
def col(index):
    result=''
    while index: index, digit=divmod(index-1,26);result=chr(65+digit)+result
    return result
def serial(day): return (datetime.fromisoformat(day)-datetime(1899,12,30)).days
def xmlcell(address, value, style, formula=None):
    prefix=f'<c r="{address}" s="{style}"'
    if formula is not None:
        return (prefix+'><f>'+escape(formula.lstrip('='))+'</f><v>'+str(value)+'</v></c>').encode()
    if value is None or value=='': return (prefix+'/>').encode()
    if isinstance(value,(int,float,Decimal)): return (prefix+'><v>'+str(value)+'</v></c>').encode()
    assert len(str(value))<=32767
    return (prefix+' t="inlineStr"><is><t xml:space="preserve">'+escape(str(value))+'</t></is></c>').encode('utf-8')
def decode(raw, strings):
    c=etree.fromstring(raw)
    if c.get('t')=='s': return strings[int(c.findtext('v'))]
    if c.get('t')=='inlineStr': return ''.join(c.itertext())
    return c.findtext('v')
def bindings(z):
    wb=etree.fromstring(z.read('xl/workbook.xml'))
    rels={r.get('Id'):r.get('Target').lstrip('/') for r in etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    out=[]
    for sheet in wb.find(N+'sheets'):
        target=rels[sheet.get('{'+REL+'}id')]
        out.append((sheet.get('name'),target if target.startswith('xl/') else 'xl/'+target))
    return out

def inputs():
    ledger=read('ledger.json');mapping={r['id']:r for r in read('mapping.json')}
    prior=read('issued-2026-chain-view-audit.json');existing=defaultdict(list)
    for r in prior['records']:
        if r['doc']: existing[r['year']+'/'+r['doc']].append(r)
    reverse=read('reverse-chain-audit.json')
    resolved={r['key']:r['resolved'] for c in reverse['components'] for r in c['invoices']}
    edges=read('edges.json')+reverse['recovered_edges'];del reverse,prior
    parent={};children=defaultdict(set);replaced=set();kind={}
    for e in edges:
        a,b=e['old'],e['new'];assert b not in parent or parent[b]==a
        parent[b]=a;children[a].add(b);kind[b]=e['kind']
        if e['kind']=='replacement':replaced.add(a)
    @lru_cache(None)
    def root(k): return root(parent[k]) if k in parent else k
    @lru_cache(None)
    def members(r):
        seen=set();todo=[r]
        while todo:
            k=todo.pop()
            if k in seen:continue
            seen.add(k);todo.extend(children[k]-seen)
        assert seen<=ledger.keys()
        return frozenset(seen)
    @lru_cache(None)
    def model(dk):
        m=mapping[dk];phases=[set() for _ in range(4)]
        if dk in existing:
            for r in existing[dk]:
                for phase in range(4):phases[phase].update(r['phases'][str(phase)])
        else:
            seeds=set(m['keys'])|set(m['history'])
            assert seeds and seeds<=ledger.keys(),dk
            roots={root(k) for k in seeds}
            processed={r for r in roots if len(members(r))>1 or ledger[r]['day'][:4]<'2026'}
            standalone=roots-processed
            phases[0].update(processed if processed else standalone)
            if processed:phases[3].update(standalone)
            for r in processed:
                for k in members(r)-{r}:phases[1 if ledger[k]['day'][:4]<'2026' else 2].add(k)
        nodes=set().union(*phases)
        assert sum(map(len,phases))==len(nodes),dk
        assert set(m['keys'])<=nodes,(dk,'Thiếu HĐ đã chốt',set(m['keys'])-nodes)
        assert all(children[k]<=nodes for k in nodes),dk
        notes=[];amounts={};active={};fallback=[];missing=[]
        has_replacement=any(k in replaced or kind.get(k)=='replacement' or 'thay the' in norm(ledger[k]['status']) for k in nodes)
        has_adjustment=any(kind.get(k)=='adjustment' or 'dieu chinh' in norm(ledger[k]['status']) for k in nodes)
        chain_type='Thay thế' if has_replacement else 'Điều chỉnh' if has_adjustment else 'HĐ mới'
        if has_replacement and has_adjustment:notes.append('Chuỗi có cả điều chỉnh và thay thế; xếp loại Thay thế')
        if any('huy' in norm(ledger[k]['status']) for k in nodes):notes.append('Có HĐ đã hủy; xem trạng thái tại tab chi tiết')
        if len(phases[0])>1:notes.append(f'{len(phases[0])} HĐ gốc; mỗi HĐ chỉ tính một lần')
        if len(m['group_docs'])>1:notes.append('HĐ dùng chung '+', '.join(m['group_docs'])+'; chưa phân bổ tiền theo BHH')
        if m['issues']:notes.extend(m['issues'])
        for k in nodes:
            x=ledger[k];g,v,p=resolved.get(k,x['money'])
            pay=x['money'][2]
            if pay is None:
                if p is not None:pay=p
                elif g is not None and v is not None:pay=g+v
                if pay is not None:fallback.append(k)
            amounts[k]=pay
            excluded=k in replaced or 'bi thay the' in norm(x['status']) or 'huy' in norm(x['status']) or 'khong hop le' in norm(x['cqt_status'])
            active[k]=not excluded
            if not excluded and pay is None:missing.append(k)
        if fallback:notes.append(f'{len(fallback)} HĐ trống Thanh toán nguồn; dùng số xác nhận từ nguồn liên kết/chi tiết hoặc Tiền hàng + VAT')
        if any('khong hop le' in norm(ledger[k]['cqt_status']) for k in nodes):notes.append('Không cộng HĐ CQT không hợp lệ')
        if missing:notes.append('Chưa đủ số Thanh toán để tính giá trị cuối')
        final=None if missing else sum(Decimal(str(amounts[k])) for k in nodes if active[k])
        if dk in existing and len(existing[dk])==1:assert final==existing[dk][0]['final_value'],dk
        return {'doc':m['doc'],'phases':[sorted(ks,key=lambda k:(ledger[k]['day'],k)) for ks in phases],
                'amounts':amounts,'active':active,'final':final,'notes':'; '.join(dict.fromkeys(notes)),
                'current':m['symbol']+'/'+m['invoice'],'fallback':fallback,'chain_type':chain_type}
    return ledger,model,kind

def source_rows(path,year):
    with ZipFile(path) as z:
        ss=[''.join(si.itertext()) for si in etree.fromstring(z.read('xl/sharedStrings.xml'))]
        result=[];seen=set()
        for name,member in bindings(z):
            rows={};raw=z.read(member)
            for match in ROW.finditer(raw):
                rn=int(match[1])
                if rn<4:continue
                c=next((c for c in CELL.finditer(match[0]) if c[1]==b'C'),None)
                doc=str(decode(c[0],ss) or '').strip() if c else ''
                if not doc:continue
                dk=str(year)+'/'+doc
                assert dk not in seen,(path.name,dk);seen.add(dk);rows[rn]=dk
            result.append((name,member,rows))
        return result

def styles(raw):
    x=etree.fromstring(raw);fonts=x.find(N+'fonts');fills=x.find(N+'fills');xfs=x.find(N+'cellXfs');nums=x.find(N+'numFmts')
    def font(color,bold=False):
        idx=len(fonts);f=etree.SubElement(fonts,N+'font');etree.SubElement(f,N+'sz',val='10');etree.SubElement(f,N+'name',val='Arial');etree.SubElement(f,N+'color',rgb='FF'+color.lstrip('#'))
        if bold:etree.SubElement(f,N+'b')
        return idx
    dark=font('#233348');white=font('#FFFFFF',True);bold=font('#233348',True)
    fmt=max([163]+[int(n.get('numFmtId')) for n in nums])+1
    for idx,code in [(fmt,'#,##0;[Red](#,##0);0'),(fmt+1,'dd/mm/yyyy')]:etree.SubElement(nums,N+'numFmt',numFmtId=str(idx),formatCode=code)
    def style(fill,typ='text',header=False):
        fi=len(fills);f=etree.SubElement(fills,N+'fill');pf=etree.SubElement(f,N+'patternFill',patternType='solid');etree.SubElement(pf,N+'fgColor',rgb='FF'+fill.lstrip('#'));etree.SubElement(pf,N+'bgColor',indexed='64')
        idx=len(xfs);xf=etree.SubElement(xfs,N+'xf',numFmtId=str(fmt if typ=='money' else fmt+1 if typ=='date' else 0),fontId=str(white if header else bold if typ=='final' else dark),fillId=str(fi),borderId='0',xfId='0',applyFont='1',applyFill='1',applyAlignment='1',applyNumberFormat='1')
        if typ=='final':xf.set('numFmtId',str(fmt))
        etree.SubElement(xf,N+'alignment',horizontal='center' if header or typ=='date' else 'right' if typ in ['money','final'] else 'left',vertical='center',wrapText='1')
        return idx
    ids={'header':[],'body':[]}
    for heading,body in PALETTE:
        ids['header'].append(style(heading,header=True));ids['body'].append({t:style(body,t) for t in ['text','date','money']})
    ids.update(meta=style('#F3F6F9'),head=style('#253B53',header=True),final=style('#E4F0EF','final'),missing=style('#FFF1D6','money'))
    for group in [fonts,fills,xfs,nums]:group.set('count',str(len(group)))
    return etree.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True),ids

def rowappend(raw,extra,height=None):
    head=re.match(rb'<row\b[^>]*',raw).group()
    updated=re.sub(rb'\s+spans="[^"]*"',b'',head)
    if height is not None:
        updated=re.sub(rb'\s+(?:ht|customHeight)="[^"]*"',b'',updated)+f' ht="{height}" customHeight="1"'.encode()
    raw=updated+raw[len(head):]
    if raw.endswith(b'/>'):return raw[:-2]+b'>'+extra+b'</row>'
    return raw[:-6]+extra+b'</row>'

def detail_xml(ordered,model,ledger,st):
    out=io.BytesIO();ranges={};rn=3;phase_counts=Counter();missing=0
    out.write(('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="'+NS+'"><dimension ref="A1:L999999"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="24"/><cols>').encode())
    for c,width in enumerate([22,12,15,16,15,21,19,23,24,24,24,65],1):out.write(f'<col min="{c}" max="{c}" width="{width}" customWidth="1"/>'.encode())
    out.write(b'</cols><sheetData>')
    out.write(b'<row r="1" ht="24">'+xmlcell('A1','Chi tiết chuỗi HĐ - tiền đã gồm VAT',st['meta'])+b'</row>')
    out.write(b'<row r="2" ht="32">'+xmlcell('A2','Mỗi HĐ chỉ xuất hiện một lần trong một BHH. HĐ dùng chung nhiều BHH có ghi chú tại tab Bán hàng.',st['meta'])+b'</row>')
    heads=['BHH / chứng từ','Giai đoạn','Ký hiệu HĐ','Số HĐ','Ngày HĐ','Tiền hàng nguồn (đ)','VAT nguồn (đ)','Thanh toán nguồn (đ)','Thanh toán dùng tính (đ)','Tình trạng tính','Giá trị còn hiệu lực (đ)','Trạng thái / căn cứ']
    out.write(b'<row r="3" ht="40">'+b''.join(xmlcell(col(c)+'3',v,st['head']) for c,v in enumerate(heads,1))+b'</row>')
    for dk in ordered:
        r=model(dk);start=rn+1;pr={}
        for phase,keys in enumerate(r['phases']):
            ps=rn+1;total=None if any(r['amounts'][k] is None for k in keys) else sum(Decimal(str(r['amounts'][k])) for k in keys)
            for k in keys:
                rn+=1;x=ledger[k];pay=r['amounts'][k];active=r['active'][k];phase_counts[phase]+=1
                note=x['status']+(' | '+x['cqt_status'] if x['cqt_status'] else '')
                if k in r['fallback']:note+=' | Thanh toán nguồn trống; số dùng tính lấy từ liên kết/chi tiết hoặc Tiền hàng + VAT'
                values=[r['doc'],'F'+str(phase),*k.split('/'),serial(x['day']),*x['money'],pay,'Còn hiệu lực' if active else 'Không cộng',pay if active else 0,note]
                cells=[]
                for c,value in enumerate(values,1):
                    typ='date' if c==5 else 'money' if c in [6,7,8,9,11] else 'text';fmt=st['body'][phase][typ]
                    if c in [6,7,8,9] and value is None:fmt=st['missing'];missing+=int(c==8)
                    formula=f'IF(J{rn}="Còn hiệu lực",I{rn},0)' if c==11 and value is not None else None
                    cells.append(xmlcell(col(c)+str(rn),value,fmt,formula))
                out.write(f'<row r="{rn}" ht="30">'.encode()+b''.join(cells)+b'</row>')
            pr[phase]={'start':ps,'end':rn,'sum':total,'count':len(keys)}
        ranges[dk]={'start':start,'end':rn,'phases':pr}
    out.write(f'</sheetData><autoFilter ref="A3:L{rn}"/><mergeCells count="2"><mergeCell ref="A1:L1"/><mergeCell ref="A2:L2"/></mergeCells><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>'.encode())
    data=out.getvalue().replace(b'A1:L999999',f'A1:L{rn}'.encode(),1)
    return data,ranges,{'invoice_memberships':rn-3,'phase_counts':dict(phase_counts),'raw_blank_payments':missing}

def patch_sheet(raw,rows,model,ledger,ranges,st,kind):
    def patch(match):
        rn=int(match[1]);extra=[]
        if rn==1:extra.append(xmlcell('M1','F0: gốc; F1: ĐC/TT trước 2026; F2: xử lý năm 2026; F3: xuất mới/xuất lại 2026. Tiền gồm VAT; từng HĐ ở tab Chi tiết chuỗi HĐ.',st['meta']))
        elif rn==2:
            extra.append(xmlcell('M2','Phân loại',st['head']))
            for phase,label in enumerate(TITLES):
                for c in range(14+phase*3,17+phase*3):extra.append(xmlcell(col(c)+'2',label if c==14+phase*3 else None,st['header'][phase]))
            extra.extend([xmlcell('Z2','Kết quả sau xử lý',st['head']),xmlcell('AA2',None,st['head'])])
        elif rn==3:
            extra.append(xmlcell('M3','Loại chuỗi',st['head']))
            for phase in range(4):
                for off,label in enumerate(['Ký hiệu / số HĐ',f'Ngày F{phase}',f'Tổng thanh toán F{phase} (đ)']):extra.append(xmlcell(col(14+phase*3+off)+'3',label,st['header'][phase]))
            extra.extend([xmlcell('Z3','Giá trị sau cùng (đ)',st['head']),xmlcell('AA3','Ghi chú chuỗi',st['head'])])
        elif rn in rows:
            dk=rows[rn];r=model(dk);rr=ranges[dk]
            extra.append(xmlcell('M'+str(rn),r['chain_type'],st['meta']))
            for phase,keys in enumerate(r['phases']):
                c=14+phase*3;fmt=st['body'][phase];p=rr['phases'][phase]
                extra.append(xmlcell(col(c)+str(rn),'\n'.join(keys),fmt['text']))
                days=[ledger[k]['day'] for k in keys]
                date=serial(days[0]) if len(days)==1 else '\n'.join(datetime.fromisoformat(d).strftime('%d/%m/%Y') for d in days)
                extra.append(xmlcell(col(c+1)+str(rn),date,fmt['date'] if len(days)==1 else fmt['text']))
                formula=f"SUM('{DETAIL}'!I{p['start']}:I{p['end']})" if keys and p['sum'] is not None else None
                extra.append(xmlcell(col(c+2)+str(rn),p['sum'] if keys else None,fmt['money'] if p['sum'] is not None or not keys else st['missing'],formula))
            formula=f"SUM('{DETAIL}'!K{rr['start']}:K{rr['end']})" if r['final'] is not None else None
            extra.extend([xmlcell('Z'+str(rn),r['final'],st['final'] if r['final'] is not None else st['missing'],formula),xmlcell('AA'+str(rn),r['notes'],st['meta'])])
        if not extra:return match[0]
        height=44 if rn==3 else max(24,16*max(len(p) for p in model(rows[rn])['phases'])+8,15*((len(model(rows[rn])['notes'])+64)//65)+8) if rn in rows else None
        return rowappend(match[0],b''.join(extra),height)
    data=ROW.sub(patch,raw)
    data=re.sub(rb'(<dimension\b[^>]*ref="[^"]*:)[A-Z]+(\d+")',rb'\g<1>AA\2',data,count=1)
    added=b''.join(f'<col min="{c}" max="{c}" width="{width}" customWidth="1"/>'.encode() for c,width in enumerate([22]+[34,15,22]*4+[23,65],13))
    data=data.replace(b'</cols>',added+b'</cols>',1)
    groups=['M1:AA1']+[f'{col(14+i*3)}2:{col(16+i*3)}2' for i in range(4)]+['Z2:AA2']
    def merges(m):
        count=int(re.search(rb'count="(\d+)"',m[0])[1])+len(groups)
        return re.sub(rb'count="\d+"',f'count="{count}"'.encode(),m[0][:-13],count=1)+b''.join(f'<mergeCell ref="{r}"/>'.encode() for r in groups)+b'</mergeCells>'
    data=re.sub(rb'<mergeCells\b.*?</mergeCells>',merges,data,count=1,flags=re.S)
    data=data.replace(b'<mergeCells',f'<autoFilter ref="A3:AA{max(rows)}"/><mergeCells'.encode(),1)
    def views(m):
        v=m[0];v=re.sub(rb'\s+showZeros="[^"]*"',b'',v)
        if v.endswith(b'/>'):return v[:-2]+b' showZeros="1"><pane xSplit="3" ySplit="3" topLeftCell="D4" activePane="bottomRight" state="frozen"/></sheetView>'
        raise AssertionError('Sheet view cần kiểm tra trước khi sửa')
    data=re.sub(rb'<sheetView\s[^>]*?/>',views,data,count=1)
    assert len(ROW.findall(raw))==len(ROW.findall(data))
    # Bảo toàn byte mọi ô gốc A:L, kể cả công thức và ô tổng cộng.
    old=[m[0] for m in CELL.finditer(raw)];new=[m[0] for m in CELL.finditer(data) if len(m[1])==1 and m[1]<=b'L']
    assert old==new,'Có thay đổi ô gốc A:L'
    etree.fromstring(data)
    return data

def main():
    ledger,model,kind=inputs();plans=[]
    for year in [2024,2025]:
        path=SRC/f'Ban_hang ({year}).xlsx';backup=BACKUP/path.name
        original=backup if backup.exists() else path
        sheets=source_rows(original,year);ordered=[dk for _,_,rs in sheets for dk in rs.values()]
        for dk in ordered:model(dk)
        print(json.dumps({'year':year,'documents':len(ordered),'unknown_final':sum(model(dk)['final'] is None for dk in ordered),'multiple_invoices':sum(any(len(ks)>1 for ks in model(dk)['phases']) for dk in ordered)},ensure_ascii=False),flush=True)
        plans.append((year,path,original,sheets,ordered,sha(path)))
    if '--plan' in sys.argv:return
    prior=read(REPORT.name) if REPORT.exists() else None
    BACKUP.mkdir(exist_ok=True)
    staged=[];report={'passed':False,'writer':'OOXML preservation fallback after previous Artifact Tool out-of-memory','files':[]}
    for year,path,original,sheets,ordered,before in plans:
        backup=BACKUP/path.name
        if original==path:shutil.copy2(path,backup)
        else:assert before in [sha(backup),next(r['after_sha256'] for r in prior['files'] if r['year']==year)],'Nguồn đã sửa đồng thời'
        target=OUT/f'ban-hang-{year}.xlsx.writing'
        with ZipFile(backup) as zi:
            newstyles,st=styles(zi.read('xl/styles.xml'))
            detail,ranges,counts=detail_xml(ordered,model,ledger,st)
            wb=etree.fromstring(zi.read('xl/workbook.xml'));assert DETAIL not in [s.get('name') for s in wb.find(N+'sheets')]
            rels=etree.fromstring(zi.read('xl/_rels/workbook.xml.rels'));rid='rId'+str(max(int(r.get('Id')[3:]) for r in rels)+1)
            number=max(int(s.get('sheetId')) for s in wb.find(N+'sheets'))+1;member=f'xl/worksheets/sheet{number}.xml'
            etree.SubElement(wb.find(N+'sheets'),N+'sheet',name=DETAIL,sheetId=str(number),attrib={'{'+REL+'}id':rid})
            etree.SubElement(rels,'{'+PKG+'}Relationship',Id=rid,Type=REL+'/worksheet',Target=f'worksheets/sheet{number}.xml')
            types=etree.fromstring(zi.read('[Content_Types].xml'));etree.SubElement(types,'{'+etree.QName(types).namespace+'}Override',PartName='/'+member,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
            changed={'xl/styles.xml':newstyles,'xl/workbook.xml':etree.tostring(wb,xml_declaration=True,encoding='UTF-8',standalone=True),'xl/_rels/workbook.xml.rels':etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True),'[Content_Types].xml':etree.tostring(types,xml_declaration=True,encoding='UTF-8',standalone=True)}
            for name,part,rs in sheets:changed[part]=patch_sheet(zi.read(part),rs,model,ledger,ranges,st,kind)
            with ZipFile(target,'w',compression=ZIP_DEFLATED,compresslevel=6) as zo:
                for info in zi.infolist():zo.writestr(info,changed.get(info.filename,zi.read(info.filename)))
                zo.writestr(member,detail)
        with ZipFile(backup) as zi,ZipFile(target) as zo:
            assert zo.testzip() is None
            for part in zi.namelist():
                if part not in changed:assert zi.read(part)==zo.read(part)
            # Cấu trúc styles, số lượng ô và bộ lọc hợp lệ trước khi thay nguồn.
            for part in [*changed,member]:etree.fromstring(zo.read(part))
        result={'year':year,'file':str(path),'backup':str(backup),'before_sha256':sha(backup),'after_sha256':sha(target),'documents':len(ordered),'original_sheets':[name for name,_,_ in sheets],'original_A_L_cells_unchanged':True,'new_columns':'M:AA','chain_type_column':'M','chain_type_counts':dict(Counter(model(dk)['chain_type'] for dk in ordered)),'chain_type_rule':'Thay thế nếu có thay thế; nếu không có thì Điều chỉnh nếu có điều chỉnh; còn lại HĐ mới. Chuỗi có cả hai ghi rõ tại AA.','detail_sheet':DETAIL,'detail_member':member,'original_members_unchanged_except':list(changed),**counts}
        report['files'].append(result);staged.append((path,target,before))
        print('STAGED',year,len(ordered),counts,flush=True)
    for path,target,before in staged:assert sha(path)==before,'Nguồn thay đổi trong lúc xuất'
    for path,target,before in staged:os.replace(target,path)
    report['passed']=True
    REPORT.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'passed':True,'updated_in_place':2,'documents':sum(r['documents'] for r in report['files'])},ensure_ascii=False),flush=True)

if __name__=='__main__':main()
