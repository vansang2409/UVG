"""Rà liên kết mã đơn ghi chú; chỉ đọc snapshot Google Sheets, không sửa nguồn/import."""
from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import json, unicodedata, hashlib, sys

sys.stdout.reconfigure(encoding='utf-8')
ROOT=Path(__file__).resolve().parents[1]
HERE=ROOT/'scripts/data/link-audit-20261021'
CUT='2025-07-01'
raw=defaultdict(list)
manifest=json.loads((HERE/'source-manifest.json').read_text(encoding='utf-8'))
hashes={}
for f in manifest['files']:
    p=ROOT/f['file'];part=json.loads(p.read_text(encoding='utf-8'))
    raw[part['sheet']].extend(part['rows'])
    hashes[f['file']]=hashlib.sha256(p.read_bytes()).hexdigest()
def at(r,i): return r[i] if i<len(r) else None
def text(v):
    return '' if v is None or str(v).startswith('#') else unicodedata.normalize('NFC',str(v)).strip()
def name(v): return ' '.join(text(v).lower().split())
def code(v):
    if isinstance(v,(int,float)) and abs(v)>2**53-1:return ''
    return text(v)
def ino(v):
    s=text(v)
    if isinstance(v,float) and v.is_integer():s=str(int(v))
    return s.zfill(8) if s.isdigit() and len(s)<=8 else ''
def day(v):
    if isinstance(v,(int,float)) and 30000<v<60000:return (datetime(1899,12,30)+timedelta(days=int(v))).strftime('%Y-%m-%d')
    token=text(v).split(' ')[0].split('T')[0]
    for fmt in ['%d/%m/%Y','%Y-%m-%d','%d-%b-%Y','%d-%b-%y','%d-%m-%Y','%d.%m.%Y']:
        try:return datetime.strptime(token,fmt).strftime('%Y-%m-%d')
        except ValueError:pass
    return ''
def money(v):
    if v is None or str(v).startswith('#') or text(v)=='':return None
    if isinstance(v,(int,float)):return Decimal(str(v))
    s=text(v).replace(' ','')
    if ',' in s:s=s.replace('.','').replace(',','.')
    elif '.' in s and all(len(p)==3 for p in s.lstrip('-').split('.')[1:]):s=s.replace('.','')
    try:return Decimal(s)
    except InvalidOperation:return None
def rnd(v):return v.quantize(Decimal(1),rounding=ROUND_HALF_UP) if v is not None else None
def near(a,b):return a is not None and b is not None and abs(rnd(a)-rnd(b))<=1
def add(m,k,v):
    if k:m[k].append(v)
def count(xs,k):return dict(Counter(x[k] for x in xs))

A=[];N=[];M=[];G=[]
for x in raw['AMIS']:
    r=x['v']
    if x['row']==1:continue
    A.append(dict(row=x['row'],date=day(at(r,0)),doc=text(at(r,2)),buyer=text(at(r,4)),tax=money(at(r,6)),total=money(at(r,7)),code=code(at(r,8)),derivedInvoice=ino(at(r,9))))
for x in raw['Đối soát NB']:
    r=x['v']
    if x['row']==1:continue
    N.append(dict(row=x['row'],date=day(at(r,0)),invoice=ino(at(r,2)),buyer=text(at(r,7)),code=code(at(r,8)),ordered=day(at(r,9)),received=day(at(r,10)),paid=day(at(r,11)),goods=money(at(r,17)),tax=money(at(r,18)),total=money(at(r,19))))
for x in raw['Meinvoice 07 -> 12 / 2025']:
    r=x['v']
    if x['row']==1:continue
    M.append(dict(row=x['row'],series=text(at(r,1)),invoice=ino(at(r,2)),date=day(at(r,3)),code=code(at(r,5)),buyer=text(at(r,11)),goods=money(at(r,15)),tax=money(at(r,16)),total=money(at(r,17)),status=text(at(r,20)),derivedDoc=text(at(r,23)),sourceNote=text(at(r,25))))
for x in raw['MISA_HD_GOC']:
    if x['row']>1:G.append(dict(row=x['row'],doc=text(at(x['v'],0)),code=code(at(x['v'],1))))
ao=defaultdict(list);mo=defaultdict(list);no=defaultdict(list);nd=defaultdict(list);nn=defaultdict(list);md=defaultdict(list);go=defaultdict(list)
for a in A:add(ao,a['code'],a)
for n in N:
    add(no,n['code'],n);add(nn,n['invoice'],n)
    if n['date'] and n['invoice']:add(nd,(n['invoice'],n['date']),n)
for m in M:
    add(mo,m['code'],m)
    if m['date'] and m['invoice']:add(md,(m['invoice'],m['date']),m)
for g in G:add(go,g['code'],g)
review=json.loads((HERE/'source/review-0.json').read_text(encoding='utf-8'))
kt={}
m_by_row={m['row']:m for m in M}
for rr,r in enumerate(review['values'][1:],2):
    if at(r,4) is not None:
        source_row=int(r[4]);actual=m_by_row[source_row]
        assert (text(at(r,5)),ino(at(r,6)))==(actual['series'],actual['invoice']), ('KT note reference changed',rr)
        kt[source_row]=dict(row=rr,note=text(at(r,15)),selected=ino(at(r,14)))

results=[]
for m in M:
    candidates=nd[(m['invoice'],m['date'])]
    nms=[n for n in candidates if name(n['buyer']) and name(n['buyer'])==name(m['buyer']) and set(n['buyer'].replace(' ',''))!={'*'}]
    full=[n for n in nms if all(near(n[k],m[k]) for k in ['goods','tax','total'])]
    codes=sorted({n['code'] for n in full if n['code']})
    if not m['invoice'] or not m['date']:cat='missing_invoice_info'
    elif not nn[m['invoice']]:cat='no_nb_invoice'
    elif not candidates:cat='invoice_date_conflict'
    elif not nms:cat='buyer_conflict'
    elif not full:cat='money_conflict'
    elif not codes:cat='nb_code_missing'
    elif len(codes)>1:cat='multiple_nb_orders'
    else:cat='fields_match'
    n=next((n for n in full if len(codes)==1 and n['code']==codes[0]),None)
    am=ao[n['code']] if n else []
    amatch=[a for a in am if near(a['tax'],m['tax']) and near(a['total'],m['total'])]
    docs=sorted({a['doc'] for a in amatch if a['doc']})
    dates=sorted({a['date'] for a in amatch if a['date']})
    formula_docs=sorted({g['doc'] for g in go[m['code']] if g['doc']})
    if cat=='fields_match':
        other=[mm for mm in md[(m['invoice'],m['date'])] if name(mm['buyer'])==name(n['buyer']) and all(near(mm[k],n[k]) for k in ['goods','tax','total'])]
        if len(other)>1:cat='ambiguous_invoice_series'
        elif m['code'] and m['code']!=n['code']:cat='order_note_conflict'
        elif not am:cat='no_amis_order'
        elif not amatch:cat='amis_money_conflict'
        elif len(docs)!=1 or len(dates)!=1:cat='multiple_amis_documents'
        elif len(formula_docs)>1:cat='ambiguous_formula_lookup'
        elif m['derivedDoc'] and m['derivedDoc']!=docs[0]:cat='derived_amis_doc_conflict'
        else:cat='corroborated'
    direct=ao[m['code']]
    direct_matches=[a for a in direct if near(a['tax'],m['tax']) and near(a['total'],m['total'])]
    directcat='missing_code' if not m['code'] else 'no_amis_code' if not direct else 'amount_mismatch' if not direct_matches else 'multiple_docs' if len({a['doc'] for a in direct_matches})>1 else 'amount_match'
    r=dict(m,category=cat,directCategory=directcat,nbCandidateRows=[n['row'] for n in candidates],nbSameNumberRows=[n['row'] for n in nn[m['invoice']]],nbMatchingRows=[n['row'] for n in full],nbCodes=codes,nbCode=n['code'] if n else '',ordered=n['ordered'] if n else '',received=n['received'] if n else '',paid=n['paid'] if n else '',nbBuyer=n['buyer'] if n else '',amisMatchingRows=[a['row'] for a in amatch],amisDocs=docs,amisDates=dates,directAmisRows=[a['row'] for a in direct],directAmisDocs=[a['doc'] for a in direct],directAmisDates=[a['date'] for a in direct],directAmisTax=[a['tax'] for a in direct],directAmisTotal=[a['total'] for a in direct],formulaDocCandidates=formula_docs,maskedBuyer=bool(n and '*' in n['buyer']),orderInvoices=[f"{mm['series']}/{mm['invoice']}" for mm in mo[m['code']]],kt=kt.get(m['row'],{}))
    r['oldReceive']=bool(n and n['received'] and n['received']<CUT<=m['date'])
    r['oldAmis']=bool(len(dates)==1 and dates[0]<CUT<=m['date'])
    r['oldAmisViaNotedCode']=bool(m['date']>=CUT and m['status']=='Hoá đơn mới' and any(a['date'] and a['date']<CUT for a in direct_matches))
    r['duplicateOrder']=bool(m['code'] and len(set(r['orderInvoices']))>1)
    results.append(r)

old=[r for r in results if r['oldReceive']]
corroborated_old=[r for r in old if r['category']=='corroborated']
new_status='Hoá đơn mới'
summary=dict(readAt='2026-10-21',spreadsheetId=manifest['spreadsheetId'],sourceCounts=dict(AMIS=len(A),Meinvoice=len(M),NB=len(N),MISA_HD_GOC=len(G)),categories=count(results,'category'),directCategories=count(results,'directCategory'),statuses=count(results,'status'),directMismatchNew=sum(r['directCategory']=='amount_mismatch' and r['status']==new_status for r in results),duplicateOrderCodes=sum(len({(m['series'],m['invoice']) for m in ms})>1 for k,ms in mo.items()),duplicateNBOrderCodes=sum(len({(n['invoice'],n['date']) for n in ns})>1 for k,ns in no.items()),ktNotedRows=sum(bool(v['note']) for v in kt.values()),ktNotedCodes=len({m['code'] for m in M if kt.get(m['row'],{}).get('note')}),oldReceiveCategories=count(old,'category'),oldReceiveTotal=len(old),oldReceiveCorroborated=len(corroborated_old),oldReceiveRepeatedCodes=sum(r['duplicateOrder'] for r in corroborated_old),oldReceiveMaskedNames=sum(r['maskedBuyer'] for r in corroborated_old),moneyRule='So sánh từng chỉ tiêu sau làm tròn đồng (ROUND_HALF_UP); dung sai tối đa1đồng/chỉ tiêu. Đây là quy tắc phân loại đối soát, không sửa số nguồn.',limits=['Mã đơn meInvoice/NB là ghi chú; khớp ba nguồn không chứng minh tuyệt đối mã note đúng đơn thực tế.','NB không có ký hiệu HĐ; chỉ nhận khớp khi số/ngày/tên/bộ tiền chọn duy nhất một hóa đơn.','Không gộp hoặc tự cộng các dòng NB/AMIS vì chưa chứng minh là chi tiết hay ghi lặp.','AMIS hiện tại chỉ có ngày hạch toán2025; không chứng minh bao phủ toàn bộ đơn2024.','Các mã trùng có thể liên quan điều chỉnh/thay thế; giữ trạng thái và ghi chú KT, không tự loại hóa đơn.','Đọc từng vùng Google Sheets, không phải snapshot nguyên tử; có snapshot giá trị nguồn để tái kiểm.'],sourceHashes=hashes)
old_by_amis=[r for r in results if r['oldAmisViaNotedCode']]
extra_amis=[r for r in old_by_amis if not (r['category']=='corroborated' and r['oldReceive'])]
extra_nb=[r for r in old if r['category']!='corroborated']
summary['oldCandidateGroups']={'three_way_corroborated':len(corroborated_old),'noted_code_amis_match_without_nb_corroboration':len(extra_amis),'nb_before_cut_without_amis_match':len(extra_nb),'union':len({(r['series'],r['invoice']) for r in corroborated_old+extra_amis+extra_nb})}
summary['extraAmisCategories']=count(extra_amis,'category')
summary['nbConflictTotal']=sum(summary['categories'].get(c,0) for c in ['invoice_date_conflict','buyer_conflict','money_conflict','multiple_nb_orders'])
assert summary['sourceCounts']=={'AMIS':65495,'Meinvoice':37768,'NB':26550,'MISA_HD_GOC':65436}
assert sum(summary['categories'].values())==37768
assert len({(r['series'],r['invoice']) for r in results})==37768
assert summary['oldCandidateGroups']['union']==117, summary['oldCandidateGroups']
assert all(r['code'] for r in results if r['oldAmisViaNotedCode'])
def default(x):
    if isinstance(x,Decimal):return float(x)
    raise TypeError(type(x).__name__)
payload=dict(summary=summary,results=results,amis=A,nb=N,ktReview=review)
(HERE/'audit.json').write_text(json.dumps(payload,ensure_ascii=False,default=default),encoding='utf-8')
(HERE/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in summary.items() if k!='sourceHashes'},ensure_ascii=False,indent=2))
