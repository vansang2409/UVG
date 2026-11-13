"""Sửa đúng D/F theo liên kết chốt live, không gộp ứng viên lịch sử."""
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from collections import defaultdict,Counter
from datetime import datetime,timezone,timedelta
import json,re,shutil,hashlib,io,os,sys
sys.stdout.reconfigure(encoding='utf-8')
O=Path(__file__).parent;P=Path('D:/UVG/data/1311/amis/So_chi_tiet_ban_hang (2025).xlsx');NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def read(n):return json.loads((O/n).read_text(encoding='utf-8'))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
a=read('final-links-audit.json');records=read('mapping.json');bydoc={r['doc']:r for r in records if r['year']==2025};amis=read('amis.json');ledger=read('ledger.json')
assert not a['conflicts'] and not a['malformed'] and not a['missing_ledger']
# Cặp 11927/11929 còn mâu thuẫn giữa tab giữ lại và AMIS/Ktoan; giữ số hiện có.
held=[r for r in a['differences'] if r['doc']=='BHHS2025010461']
changes=[r for r in a['differences'] if r['doc']!='BHHS2025010461']
assert len(changes)==9
rows=defaultdict(dict)
for r in changes:
    assert r['money']==r['invoice_money'] and r['status']=='Hoá đơn mới'
    m=bydoc[r['doc']];assert (m['symbol']+'/'+m['invoice'] if m['invoice'] else None)==r['prior']
    for loc in amis['2025/'+r['doc']]['rows']:
        sn,rn=loc.rsplit('!',1);rows[sn][int(rn)]=r
stamp=datetime.now(timezone(timedelta(hours=7))).strftime('%Y%m%d_%H%M%S')
B=O/('backup-before-final-link-correction-'+stamp);B.mkdir();backup=B/P.name;shutil.copy2(P,backup);shutil.copy2(O/'mapping.json',B/'mapping.json')
previous=read('fill-amis-originals-verification.json');assert sha(P)==next(r['after_sha256'] for r in previous['files'] if Path(r['file'])==P)
CELL=re.compile(rb'<c\b[^>]*\br="([DF])(\d+)"[^>]*(?:/>|>.*?</c>)',re.S)
def val(c,strings):
    if c.get('t')=='s':return strings[int(c.find(NS+'v').text)]
    if c.get('t')=='inlineStr':return ''.join(t.text or '' for t in c.iter(NS+'t'))
    x=c.find(NS+'v');return x.text if x is not None else ''
stage=O/('corrected-amis-'+stamp+'.xlsx');stats=[]
with ZipFile(backup) as zi:
    ss=[''.join(t.text or '' for t in s.iter(NS+'t')) for s in etree.fromstring(zi.read('xl/sharedStrings.xml'))]
    wb=etree.fromstring(zi.read('xl/workbook.xml'));rels={x.get('Id'):x.get('Target').lstrip('/') for x in etree.fromstring(zi.read('xl/_rels/workbook.xml.rels'))}
    patches={}
    for sh in wb.find(NS+'sheets'):
        name=sh.get('name')
        if name not in rows:continue
        target=rels[sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')];target=target if target.startswith('xl/') else 'xl/'+target
        original=zi.read(target);seen=set()
        def patch(m):
            col=m.group(1).decode();rn=int(m.group(2))
            if rn not in rows[name]:return m.group(0)
            r=rows[name][rn];sig,number=r['final'].split('/');value=sig if col=='D' else number
            c=etree.fromstring(b'<root xmlns="'+NS[1:-1].encode()+b'">'+m.group(0)+b'</root>')[0]
            before=val(c,ss);prior=r['prior'].split('/')[0 if col=='D' else 1] if r['prior'] else ''
            assert before==prior,(name,rn,col,before,prior)
            attrs=re.match(rb'<c\b([^>]*?)(?:\s*/>|>)',m.group(0),re.S).group(1).rstrip();attrs=re.sub(rb'\s+t="[^"]*"',b'',attrs)
            seen.add((rn,col));return b'<c'+attrs+b' t="inlineStr"><is><t>'+value.encode()+b'</t></is></c>'
        updated=CELL.sub(patch,original);assert len(seen)==len(rows[name])*2
        def mask(m):return b'__TARGET_'+m.group(1)+m.group(2)+b'__' if int(m.group(2)) in rows[name] else m.group(0)
        assert CELL.sub(mask,original)==CELL.sub(mask,updated)
        checked=0
        for _,r in etree.iterparse(io.BytesIO(updated),events=('end',),tag=NS+'row'):
            rn=int(r.get('r'))
            if rn in rows[name]:
                cells={c.get('r'):c for c in r};e=rows[name][rn];sig,number=e['final'].split('/')
                assert val(cells['C'+str(rn)],ss)==e['doc']
                assert val(cells['D'+str(rn)],ss)==sig and val(cells['F'+str(rn)],ss)==number
                assert cells['F'+str(rn)].get('t')=='inlineStr';checked+=1
            r.clear()
            while r.getprevious() is not None:del r.getparent()[0]
        assert checked==len(rows[name]);patches[target]=updated;stats.append({'sheet':name,'rows':checked})
    with ZipFile(stage,'w') as zo:
        zo.comment=zi.comment
        for info in zi.infolist():zo.writestr(info,patches.get(info.filename,zi.read(info.filename)))
with ZipFile(backup) as old,ZipFile(stage) as new:
    assert new.testzip() is None and old.namelist()==new.namelist()
    for n in old.namelist():
        if n not in patches:assert old.read(n)==new.read(n)
assert sha(P)==sha(backup);os.replace(stage,P)
for r in changes:
    m=bydoc[r['doc']];sig,number=r['final'].split('/');m['previous_mapping']={'key':r['prior'],'candidates':m['candidates'],'issues':m['issues']}
    m.update(invoice=number,symbol=sig,invoice_day=ledger[r['final']]['day'],status='Khớp tiền và VAT',mi_money=r['invoice_money'],group_amis=m['amis_money'],delta=[0,0,0],group_docs=[r['doc']],keys=[r['final']],related=[],candidates=[r['final']],issues=[],evidence=['Liên kết chốt live: HĐ meInvoice giữ lại 2025; đối chiếu bảng kê 13/11'],source_invoice=ledger[r['final']]['source'])
(O/'mapping.json').write_text(json.dumps(records,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
summary=read('summary.json');summary['mapped']=sum(bool(r['invoice']) for r in records);summary['unmapped']=sum(not r['invoice'] for r in records);summary['by_year']={str(y):dict(Counter(r['status'] for r in records if r['year']==y)) for y in [2024,2025]}
(O/'summary.json').write_text(json.dumps(summary,ensure_ascii=False),encoding='utf-8')
result={'passed':True,'file':str(P),'backup':str(backup),'after_sha256':sha(P),'filled_blank_documents':sum(r['prior'] is None for r in changes),'corrected_existing_documents':sum(r['prior'] is not None for r in changes),'changed_rows':sum(r['rows'] for r in stats),'sheets':stats,'changes':changes,'source_conflict_kept_unchanged':held,'unmapped':[r['doc'] for r in records if not r['invoice']],'only_D_F_changed':True,'live_final_links_audited':a['docs_in_new_amis']}
(O/'final-link-correction-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ['changes','source_conflict_kept_unchanged']},ensure_ascii=False),flush=True)
