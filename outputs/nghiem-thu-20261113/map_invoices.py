"""Map số/ký hiệu HĐ theo bằng chứng cũ, kiểm lại trên nguồn 13/11."""
from pathlib import Path
from collections import defaultdict,Counter
from functools import lru_cache
from decimal import Decimal
import json,csv,re,sys,hashlib,unicodedata
from openpyxl import load_workbook
sys.stdout.reconfigure(encoding='utf-8')
ROOT=Path('D:/UVG');OUT=Path(__file__).parent
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def dump(n,v):(OUT/n).write_text(json.dumps(v,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
def norm(x):return ''.join(c for c in unicodedata.normalize('NFD',str(x or '').lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
def key(s,n):return str(s).strip().upper()+'/'+str(n).removesuffix('.0').zfill(8) if s and n else None
docs=read(OUT/'amis.json');ledger=read(OUT/'ledger.json');detail=read(OUT/'detail.json');edges=read(OUT/'edges.json')
parents=defaultdict(set);children=defaultdict(set);rep=defaultdict(set);edgekind={};linkedmoney=defaultdict(lambda:defaultdict(set))
for e in edges:
    a,b=e['old'],e['new'];parents[b].add(a);children[a].add(b);edgekind[a,b]=e['kind']
    if e['kind']=='replacement':rep[a].add(b)
    for k,vs in [(a,e['old_money']),(b,e['new_money'])]:
        for i,v in enumerate(vs):
            if v is not None:linkedmoney[k][i].add(v)
money_issues=[]
for k,r in ledger.items():
    r['resolved']=r['money'][:];r['fallback']=[];r['issues']=[]
    for i,v in enumerate(r['money']):
        vals=linkedmoney[k][i]
        if v is None and len(vals)==1:r['resolved'][i]=next(iter(vals));r['fallback'].append('Bảng liên kết')
        elif v is not None and vals and (len(vals)>1 or v not in vals):r['issues'].append('Tiền bảng kê khác bảng liên kết')
        # Bổ sung từ chi tiết cùng bộ 13/11 khi tổng chi tiết xác nhận được giá trị.
        if r['resolved'][i] is None and k in detail and detail[k]['known'][i]>0:
            r['resolved'][i]=detail[k]['sums'][i];r['fallback'].append('Bảng chi tiết 13/11')
    if r['issues']:money_issues.append({'key':k,'issues':r['issues']})
@lru_cache(None)
def root(k):
    todo=[k];seen=set();rs=set()
    while todo:
        x=todo.pop()
        if x in seen:continue
        seen.add(x)
        if parents[x]:todo.extend(parents[x])
        else:rs.add(x)
    return tuple(sorted(rs))
@lru_cache(None)
def chain(k):
    todo=[k];seen=set();active=[];issues=[]
    while todo:
        x=todo.pop()
        if x in seen:continue
        seen.add(x);todo.extend(children[x]);r=ledger.get(x)
        if not r:issues.append('Thiếu HĐ trong bảng kê '+x);continue
        st=norm(r['status'])
        if rep[x] or 'bi thay the' in st or 'huy' in st:
            if 'bi thay the' in st and not rep[x]:issues.append('Thiếu liên kết thay thế '+x)
            continue
        active.append(x)
        if len(parents[x])>1:issues.append('HĐ có nhiều gốc '+x)
    return {'keys':sorted(seen),'active':sorted(active),'issues':issues}
def sums(keys):
    out=[]
    for i in range(3):
        vs=[ledger.get(k,{}).get('resolved',[None]*3)[i] for k in keys]
        out.append(sum(vs) if vs and all(v is not None for v in vs) else None)
    return out
history=defaultdict(lambda:defaultdict(set));orders=defaultdict(set);evidence_manifest=[];return_direct={};return_parent={};preferred={}
def histpath(rel):
    p=ROOT/rel;evidence_manifest.append({'file':rel,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()});return p
def add(year,doc,k,why):
    dk=str(year)+'/'+doc
    if dk in docs and k:history[dk][k].add(why)
p=histpath('scripts/data/new-sales-issued-20261101/resolved-coverage.json');old=read(p)
for r in old['coverage']+old['blank_payments']:
    dk='2024/'+r['doc'];orders[dk].add(r['order']) if r['order'] else None
    for i in r['invoices']:add(2024,r['doc'],i['key'],'Đối chiếu phát hành 2024 ngày 01/11')
p=histpath('scripts/data/new-amis-2024-20261026/prepared.json');old=read(p)
for r in old['exclusions']:
    for doc in re.findall(r'BH[A-Z]*\d+(?:\.\d+)?',str(r['values'][12] or '')):add(2024,doc,r['key'],'Cột chứng từ của bảng HĐ công ty 2024')
del old
p=histpath('scripts/data/post-final-detail-20261101/replacements.json')
for r in read(p):
    if '2024/'+r['doc'] in docs:
        # HĐ giữ lại đã được xác minh độc lập trong lô trùng.
        history['2024/'+r['doc']].clear();add(2024,r['doc'],r['retained'],'HĐ giữ lại của lô trùng 2.000 HĐ')
p=histpath('outputs/import-2025-20261109/DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv')
for r in csv.DictReader(p.open(encoding='utf-8-sig')):
    keys=re.findall(r'1C\d{2}[A-Z]+/\d{8}',r['Chuỗi liên quan'])
    for k in keys:add(2025,r['Chứng từ AMIS'],k,'Bảng đối chiếu chứng từ 2025 ngày 09/11')
    if len(keys)==1:preferred['2025/'+r['Chứng từ AMIS']]=keys[0]
p=histpath('outputs/import-2025-20261109/prepared.json');old=read(p)
for doc,r in old['docs'].items():
    if r.get('order'):orders['2025/'+doc].add(str(r['order']))
for r in old['company_rows']:
    for doc in r['docs']:add(2025,doc,r['key'],'Bảng HĐ công ty 2025')
del old
for p in sorted((ROOT/'outputs/audit-new-sales-detail-20261111/live').glob('K*.json')):
    histpath(str(p.relative_to(ROOT)).replace('\\','/'))
    for r in read(p)['values']:
        r=r+[None]*(14-len(r));doc=str(r[10] or '');add(2025,doc,key(r[13],r[12]),'Số HĐ phát hành Import HD moi đã đọc lại 11/11')
        preferred['2025/'+doc]=key(r[13],r[12])
        if r[8]:orders['2025/'+doc].add(str(r[8]))
p=histpath('outputs/import-2025-20261109/shared-invoice-map-google-verification.json')
for r in read(p)['records']:
    history['2025/'+r['doc']].clear();add(2025,r['doc'],r['actual'],'Sửa liên kết 2 đơn dùng chung HĐ đã kiểm 09/11')
    preferred['2025/'+r['doc']]=r['actual']
    history['2025/'+r['kept_doc']].clear();add(2025,r['kept_doc'],r['kept_invoice'],'HĐ giữ lại sau sửa liên kết dùng chung')
    preferred['2025/'+r['kept_doc']]=r['kept_invoice']
# Ghép riêng các phiếu trả hàng vào HĐ điều chỉnh, tránh cộng cả HĐ bán vào phiếu trả.
p=histpath('outputs/btl-bhh-20261026/LIEN_KET_BTL_BHH_2024.xlsx');w=load_workbook(p,read_only=True,data_only=True)
sales_returns={}
for r in w['Phieu BTL'].iter_rows(min_row=2,values_only=True):
    if r[0] and r[4]:sales_returns[str(r[4])]=str(r[0])
w.close()
p=histpath('scripts/data/post-final-detail-20261101/adjustments.json')
for r in read(p):
    v=sales_returns.get(r['doc'])
    if v:return_direct['2024/'+v]=r['new'];return_parent['2024/'+v]=r['old']
p=histpath('outputs/btl-bgg-issued-20261111/mapping.json')
for tab in read(p)['sheets']:
    for r in tab['mapped']:
        return_direct['2025/'+r['voucher']]=key(r['symbol'],r['invoice']);return_parent['2025/'+r['voucher']]=r['key']
# Quan hệ 19974 -> 21507 bị thiếu trong bảng liên kết mới, đã có nguồn lịch sử rõ ràng.
p=histpath('outputs/btl-replacement-audit-20261111/result.json')
for r in read(p)['replacements']:
    for e in r['incoming']:
        a,b=e['old'],e['new']
        if a in ledger and b in ledger and not rep[a] and 'bi thay the' in norm(ledger[a]['status']):
            parents[b].add(a);children[a].add(b);rep[a].add(b);edgekind[a,b]='replacement'
    for ref in r['references']:
        add(2025,ref['doc'],r['key'],'HĐ thay thế 21507 và chứng từ BTL00475 đã kiểm 11/11')
        if rep[r['key']]:
            leaf=list(rep[r['key']])
            if len(leaf)==1 and ledger[leaf[0]]['money'][:2]==[0,0]:
                return_direct['2025/'+ref['btl']]=leaf[0]
                return_parent['2025/'+ref['btl']]=r['key']
return_keys={k for k in return_direct.values() if norm(ledger.get(k,{}).get('status'))=='hoa don dieu chinh'}
# Ưu tiên các liên kết chốt đã đối chiếu live; không mở lại các ứng viên lịch sử đã loại.
final_correction_path=OUT/'final-link-correction-verification.json'
if final_correction_path.exists():
    final_corrections=read(final_correction_path)
    if final_corrections.get('passed'):
        for r in final_corrections['changes']:
            dk='2025/'+r['doc'];history[dk].clear()
            add(2025,r['doc'],r['final'],'Liên kết chốt live đã kiểm lại 13/11')
            preferred[dk]=r['final']
order_index=defaultdict(set)
for k,r in ledger.items():
    if r['order']:order_index[r['order']].update(root(k))
mapping={};ambiguous=[]
for dk,d in docs.items():
    if dk in return_direct:
        k=return_direct[dk];mapping[dk]={'roots':[],'selected':[k] if k in ledger else [],'candidates':[k],'evidence':['Liên kết phiếu trả/giảm giá và số HĐ phát hành đã kiểm'],'issues':[] if k in ledger else ['Không tìm thấy HĐ đã map trong bảng kê mới'],'history':[],'return':True};continue
    cands=defaultdict(set)
    for k,ev in history[dk].items():
        for rt in root(k):cands[rt].update(ev)
    for k in d['keys']:
        for rt in root(k):cands[rt].add('Ký hiệu và số HĐ AMIS mới')
    # Mã đơn chỉ khớp đầy đủ, không sửa đuôi 18 chữ số đã mất độ chính xác.
    for order in orders[dk]:
        for rt in order_index.get(order,[]):
            if not cands or rt in cands:cands[rt].add('Mã đơn khớp nguyên vẹn bảng kê 13/11')
    existing=[k for k in cands if k in ledger]
    live=[];zero=[]
    for rt in existing:
        active=chain(rt)['active']
        if not d['doc'].startswith(('BTL','BGG')):active=[k for k in active if k not in return_keys]
        ss=sums(active)
        # Giữ ô thanh toán trống, nhưng HĐ đã về 0 tiền hàng/VAT không cạnh tranh với HĐ bán còn giá trị.
        if ss[0]==0 and ss[1]==0 and ss[2] in (0,None):zero.append(rt)
        else:live.append(rt)
    chosen=live or zero if len(zero)==1 else live
    selected=[];iss=[];ev=[];hist=[]
    pref=preferred.get(dk);pr=root(pref) if pref else ()
    if not live and len(pr)==1 and pr[0] in zero:
        chosen=[pr[0]]
        hist+=sum([chain(k)['keys'] for k in zero if k!=pr[0]],[])
    if len(chosen)>1 and len(pr)==1 and pr[0] in chosen:
        iss.append('Có liên kết khác còn giá trị; ưu tiên liên kết chứng từ đã chốt gần nhất')
        hist+=sum([chain(k)['keys'] for k in chosen if k!=pr[0]],[])
        chosen=[pr[0]]
    if len(chosen)==1:
        rt=chosen[0];selected=chain(rt)['active'];iss+=chain(rt)['issues'];ev=sorted(cands[rt]);hist+=[k for k in chain(rt)['keys'] if k not in selected]
    elif len(chosen)>1:
        iss.append('Có nhiều chuỗi ứng viên còn giá trị; chưa chọn HĐ');ambiguous.append(dk)
    else:iss.append('Chưa tìm thấy HĐ từ mã chứng từ/liên kết đã có')
    mapping[dk]={'roots':chosen if len(chosen)==1 else [],'selected':selected,'candidates':sorted(cands),'evidence':ev,'issues':iss,'history':sorted(set(hist+sum([chain(k)['keys'] for k in zero if k not in chosen],[]))),'return':d['doc'].startswith(('BTL','BGG'))}
# HĐ điều chỉnh đã map phiếu trả được tách khỏi giá trị HĐ bán.
for dk,m in mapping.items():
    if not m['return']:m['selected']=[k for k in m['selected'] if k not in return_keys]
    m['selected']=sorted(set(m['selected']))
groups=defaultdict(list)
for dk,m in mapping.items():
    if m['selected']:groups[tuple(m['selected'])].append(dk)
records=[];status_counts=defaultdict(Counter)
for dk,d in sorted(docs.items(),key=lambda kv:(kv[1]['year'],min(kv[1]['dates']),kv[1]['doc'])):
    m=mapping[dk];ks=m['selected'];gg=groups.get(tuple(ks),[]) if ks else []
    am=[sum(docs[x][f] for x in gg) for f in ['goods','vat','pay']] if gg else [None]*3
    mi=sums(ks);delta=[a-b if a is not None and b is not None else None for a,b in zip(am,mi)]
    iss=m['issues'][:]
    for k in ks:
        if ledger[k]['fallback']:iss.append('Số tiền bổ sung từ '+', '.join(sorted(set(ledger[k]['fallback']))))
    if any(d['missing'].values()):iss.append('Thiếu số tiền ở dòng AMIS')
    if len(d['dates'])>1:iss.append('Chứng từ có nhiều ngày hạch toán')
    for k in ks:
        iss+=ledger.get(k,{}).get('issues',[])
        if k not in detail:iss.append('Thiếu chi tiết HĐ '+k)
        if 'tu choi' in norm(ledger.get(k,{}).get('cqt_status')) or 'khong hop le' in norm(ledger.get(k,{}).get('cqt_status')):iss.append('Trạng thái CQT cần kiểm tra '+k)
    if not ks:status='Chưa map' if len(m['candidates'])<=1 else 'Nhiều ứng viên'
    elif iss:status='Đã map - cần kiểm tra nguồn'
    elif any(x is None for x in delta):status='Đã map - thiếu số tiền'
    elif all(x==0 for x in delta):status='Khớp tiền và VAT' if len(gg)==1 else 'Khớp theo nhóm chứng từ'
    elif all(abs(x)<=2 for x in delta):status='Lệch làm tròn 1-2đ'
    else:status='Đã map - lệch tiền/VAT'
    # HĐ chính là HĐ bán/thay thế hiện hành. Các điều chỉnh đi kèm ghi riêng.
    primary=next((k for k in ks if 'hoa don dieu chinh'!=norm(ledger[k]['status'])),ks[0] if ks else None)
    if m['return']:primary=ks[0] if ks else None
    rec={'id':dk,'year':d['year'],'doc':d['doc'],'day':min(d['dates']),'customer':d['customer'],'invoice':primary.split('/')[1] if primary else None,'symbol':primary.split('/')[0] if primary else None,'invoice_day':ledger[primary]['day'] if primary else None,'status':status,'amis_money':[d['goods'],d['vat'],d['pay']],'group_amis':am,'mi_money':mi,'delta':delta,'group_docs':[docs[x]['doc'] for x in gg],'related':[k for k in ks if k!=primary],'history':m['history'],'candidates':m['candidates'],'evidence':m['evidence'],'issues':sorted(set(iss)),'source_amis':d['source']+' | '+', '.join(d['rows']),'source_invoice':ledger[primary]['source'] if primary else None,'keys':ks,'order':sorted(orders[dk])}
    records.append(rec);status_counts[str(d['year'])][status]+=1
summary={'as_of':'2026-11-13','docs':len(records),'mapped':sum(bool(r['invoice']) for r in records),'unmapped':sum(not r['invoice'] for r in records),'by_year':{y:dict(c) for y,c in status_counts.items()},'amis_totals':{str(y):[sum(d[f] for d in docs.values() if d['year']==y) for f in ['goods','vat','pay']] for y in (2024,2025)},'ledger':len(ledger),'detail_keys':len(detail),'edges':len(edges),'money_source_conflicts':len(money_issues)}
dump('mapping.json',records);dump('summary.json',summary);dump('history-manifest.json',evidence_manifest);dump('money-source-conflicts.json',money_issues)
print(json.dumps(summary,ensure_ascii=False),flush=True)
print('EXCEPTIONS',json.dumps([{k:r[k] for k in ['id','status','candidates','delta','issues']} for r in records if r['status'] not in ['Khớp tiền và VAT','Khớp theo nhóm chứng từ']][:25],ensure_ascii=False),flush=True)
