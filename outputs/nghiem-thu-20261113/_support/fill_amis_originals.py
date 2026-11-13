"""Chỉ cập nhật D/F trong hai file AMIS, giữ nguyên các phần OOXML khác."""
from pathlib import Path
from zipfile import ZipFile
from collections import defaultdict,Counter
from datetime import datetime,timezone,timedelta
from lxml import etree
from openpyxl import load_workbook
import json,re,hashlib,shutil,os,sys
sys.stdout.reconfigure(encoding='utf-8')
ROOT=Path('D:/UVG');OUT=Path(__file__).parent;SRC=ROOT/'data/1311/amis'
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def read(n):return json.loads((OUT/n).read_text(encoding='utf-8'))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
amis=read('amis.json');mapping={r['id']:r for r in read('mapping.json')};manifest=read('manifest.json')
expected_hash={str((ROOT/r['file']).resolve()):r['sha256'] for r in manifest if 'member' not in r}
updates=defaultdict(lambda:defaultdict(dict));unmapped=[]
for dk,d in amis.items():
    m=mapping[dk]
    if not m['invoice']:unmapped.append({'id':dk,'rows':d['rows'],'candidates':m['candidates']});continue
    assert re.fullmatch(r'\d{8}',m['invoice']) and re.fullmatch(r'1C\d{2}[A-Z]+',m['symbol'])
    for loc in d['rows']:
        sheet,rn=loc.rsplit('!',1);rn=int(rn)
        assert rn not in updates[d['source']][sheet]
        updates[d['source']][sheet][rn]={'doc':d['doc'],'symbol':m['symbol'],'invoice':m['invoice']}
stamp=datetime.now(timezone(timedelta(hours=7))).strftime('%Y%m%d_%H%M%S')
backup=OUT/('backup-amis-before-fill-'+stamp);backup.mkdir()
stage=OUT/('staged-amis-fill-'+stamp);stage.mkdir()
files=[]
for rel in updates:
    p=(ROOT/rel).resolve();assert p.parent==SRC.resolve()
    assert sha(p)==expected_hash[str(p)],'Nguồn đã thay đổi, cần đọc lại trước khi điền: '+str(p)
    b=backup/p.name;shutil.copy2(p,b);assert sha(b)==sha(p)
    files.append((rel,p,b,stage/p.name))
CELL=re.compile(rb'<c\b[^>]*\br="([DF])(\d+)"[^>]*(?:/>|>.*?</c>)',re.S)
def decode(c,strings):
    if c is None:return None
    if c.get('t')=='s':return strings[int(c.find(NS+'v').text)]
    if c.get('t')=='inlineStr':return ''.join(t.text or '' for t in c.iter(NS+'t'))
    v=c.find(NS+'v');return v.text if v is not None else None
report={'backup_directory':str(backup),'files':[],'unmapped':unmapped,'unmapped_documents':len(unmapped)}
for rel,p,b,target in files:
    with ZipFile(b) as zin:
        strings=[''.join(t.text or '' for t in si.iter(NS+'t')) for si in etree.fromstring(zin.read('xl/sharedStrings.xml'))]
        wb=etree.fromstring(zin.read('xl/workbook.xml'));rels=etree.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
        targets={x.get('Id'):x.get('Target').lstrip('/') for x in rels};members={}
        for sh in wb.find(NS+'sheets'):
            t=targets[sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
            members[sh.get('name')]=t if t.startswith('xl/') else 'xl/'+t
        changed={};sheet_stats=[]
        for name,rows in updates[rel].items():
            member=members[name];original=zin.read(member);seen=set();old_values=[]
            def replace(m):
                col=m.group(1).decode();rn=int(m.group(2))
                if rn not in rows:return m.group(0)
                raw=m.group(0);r=rows[rn];v=r['symbol'] if col=='D' else r['invoice']
                old=etree.fromstring(raw,parser=etree.XMLParser()) if b'xmlns' in raw else etree.fromstring(b'<root xmlns="'+NS[1:-1].encode()+b'">'+raw+b'</root>')[0]
                previous=decode(old,strings)
                if previous not in (None,''):old_values.append({'row':rn,'column':col,'before':previous,'after':v})
                attrs=re.match(rb'<c\b([^>]*?)(?:\s*/>|>)',raw,re.S).group(1).rstrip()
                attrs=re.sub(rb'\s+t="[^"]*"',b'',attrs)
                seen.add((col,rn))
                return b'<c'+attrs+b' t="inlineStr"><is><t>'+v.encode()+b'</t></is></c>'
            modified=CELL.sub(replace,original)
            assert len(seen)==len(rows)*2,(name,len(seen),len(rows))
            def mask(m):
                return b'__TARGET_'+m.group(1)+m.group(2)+b'__' if int(m.group(2)) in rows else m.group(0)
            assert CELL.sub(mask,original)==CELL.sub(mask,modified),'Có thay đổi ngoài ô D/F'
            # Kiểm đủ khóa chứng từ và giá trị mới trực tiếp từ XML đã sửa.
            count=0
            import io
            for _,row in etree.iterparse(io.BytesIO(modified),events=('end',),tag=NS+'row'):
                rn=int(row.get('r'))
                if rn in rows:
                    cells={c.get('r'):c for c in row};rec=rows[rn]
                    assert decode(cells.get('C'+str(rn)),strings)==rec['doc']
                    for col,field in [('D','symbol'),('F','invoice')]:
                        cell=cells[col+str(rn)];assert cell.get('t')=='inlineStr' and decode(cell,strings)==rec[field]
                    count+=1
                row.clear()
                while row.getprevious() is not None:del row.getparent()[0]
            assert count==len(rows)
            changed[member]=modified
            sheet_stats.append({'sheet':name,'filled_detail_rows':len(rows),'filled_cells':len(seen),'prior_nonblank_target_cells':old_values})
        with ZipFile(target,'w') as zout:
            zout.comment=zin.comment
            for info in zin.infolist():zout.writestr(info,changed.get(info.filename,zin.read(info.filename)))
    with ZipFile(b) as old,ZipFile(target) as new:
        assert new.testzip() is None and old.namelist()==new.namelist()
        for n in old.namelist():
            if n not in changed:assert old.read(n)==new.read(n),n
    # Đọc độc lập các dòng đầu của từng sheet bằng openpyxl.
    w=load_workbook(target,read_only=True,data_only=True)
    for st in sheet_stats:
        rn=min(updates[rel][st['sheet']]);r=next(w[st['sheet']].iter_rows(min_row=rn,max_row=rn,max_col=6))
        expected=updates[rel][st['sheet']][rn]
        assert r[3].value==expected['symbol'] and r[5].value==expected['invoice'] and r[5].data_type=='s'
    w.close()
    report['files'].append({'file':str(p),'backup':str(b),'before_sha256':sha(b),'after_sha256':sha(target),'sheets':sheet_stats,'only_D_F_cells_changed':True})
    print('STAGED AND VERIFIED',p.name,sum(s['filled_detail_rows'] for s in sheet_stats),flush=True)
# Chỉ thay file nguồn sau khi cả hai bản đã qua kiểm tra.
for rel,p,b,target in files:assert sha(p)==sha(b),'File nguồn đang được sửa đồng thời: '+str(p)
for rel,p,b,target in files:
    os.replace(target,p)
    assert sha(p)==next(x['after_sha256'] for x in report['files'] if x['file']==str(p))
    print('UPDATED',p,flush=True)
report['passed']=True;report['filled_documents']=sum(bool(m['invoice']) for m in mapping.values());report['filled_detail_rows']=sum(s['filled_detail_rows'] for f in report['files'] for s in f['sheets']);report['filled_cells']=report['filled_detail_rows']*2
(OUT/'fill-amis-originals-verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in report.items() if k!='files'},ensure_ascii=False),flush=True)
