from pathlib import Path
from collections import defaultdict,Counter
import json,re,sys
sys.stdout.reconfigure(encoding='utf-8')
O=Path(__file__).parent
def read(p):return json.loads(p.read_text(encoding='utf-8'))
old={x['doc']:x for x in read(O/'mapping.json') if x['year']==2025};ledger=read(O/'ledger.json');amis=read(O/'amis.json')
replacement=defaultdict(set)
for e in read(O/'edges.json'):
    if e['kind']=='replacement':replacement[e['old']].add(e['new'])
for r in read(Path('D:/UVG/outputs/btl-replacement-audit-20261111/result.json'))['replacements']:
    for e in r['incoming']:
        if e['old'] in ledger and e['new'] in ledger and not replacement[e['old']]:replacement[e['old']].add(e['new'])
def current(k):
    seen=set()
    while replacement[k]:
        assert k not in seen and len(replacement[k])==1,(k,replacement[k])
        seen.add(k);k=next(iter(replacement[k]))
    return k
links=defaultdict(set);source=defaultdict(list);issues=[]
for p in sorted(O.glob('live-retained-*.json')):
    x=read(p);start=int(re.search(r'!(?:B|\$B)(\d+)',x['range'])[1])
    for rn,r in enumerate(x['values'],start):
        if not r or not r[0]:continue
        doc=str(r[0]);k=str(r[1] or '') if len(r)>1 else ''
        if re.fullmatch(r'1C\d{2}[A-Z]+/\d{8}',k):links[doc].add(k);source[doc].append({'range':x['range'],'row':rn,'key':k})
        else:issues.append({'doc':doc,'value':k,'row':rn})
for p in sorted(O.glob('live-import-*.json')):
    x=read(p);start=int(re.search(r'!U(\d+)',x['range'])[1])
    for rn,r in enumerate(x['values'],start):
        if not r or not r[0]:continue
        r=r+[None]*(4-len(r));doc=str(r[0]);n=str(r[2] or '').removesuffix('.0').zfill(8);k=str(r[3] or '')+'/'+n
        if re.fullmatch(r'1C\d{2}[A-Z]+/\d{8}',k):links[doc].add(k);source[doc].append({'range':x['range'],'row':rn,'key':k})
        else:issues.append({'doc':doc,'value':k,'row':rn})
diff=[];conflicts=[];missing=[]
for doc,ks in links.items():
    if doc not in old:continue
    if len(ks)!=1:conflicts.append({'doc':doc,'keys':sorted(ks)});continue
    stored=next(iter(ks));k=current(stored);r=old[doc];prior=r['symbol']+'/'+r['invoice'] if r['invoice'] else None
    if k not in ledger:missing.append({'doc':doc,'key':k});continue
    if k!=prior:diff.append({'doc':doc,'prior':prior,'final':k,'stored_link':stored,'source':source[doc],'money':r['amis_money'],'invoice_money':ledger[k]['money'],'status':ledger[k]['status']})
result={'source_docs':len(links),'docs_in_new_amis':len(set(links)&set(old)),'differences':diff,'conflicts':conflicts,'malformed':issues,'missing_ledger':missing,'split_docs_in_new_amis':{d:amis.get('2025/'+d) for d in ['BHHN2025000166.1','BHHN000165']},'final_links':{d:[current(k) for k in sorted(v)] for d,v in links.items()}}
(O/'final-links-audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'source_docs':len(links),'differences':diff,'conflicts':conflicts,'malformed':issues,'missing_ledger':missing},ensure_ascii=False),flush=True)
