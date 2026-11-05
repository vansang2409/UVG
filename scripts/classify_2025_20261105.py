"""Phân loại chứng từ AMIS2025; không sửa nguồn, không tạo lệnh phát hành."""
from pathlib import Path
from collections import defaultdict,Counter
from functools import lru_cache
from datetime import date,datetime,timedelta
import json,re,unicodedata,hashlib
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/phan-loai-2025-20261105'
SRC=OUT/'source'
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def dump(p,x):p.write_text(json.dumps(x,ensure_ascii=False,indent=2,default=str)+'\n',encoding='utf-8')
def txt(x):return '' if x is None else str(x).strip()
def norm(x):return ''.join(c for c in unicodedata.normalize('NFD',txt(x).lower()) if unicodedata.category(c)!='Mn').replace('đ','d')
def num(x):return float(x) if isinstance(x,(int,float)) else None
def kval(s,n):
    n=txt(n)
    if not txt(s) or not re.fullmatch(r'\d+(?:\.0)?',n):return None
    k=invoice_key(s,n)
    return '/'.join(k) if k else None
def ksrc(k):return '/'.join(k) if isinstance(k,list) else k
def day(x):
    if isinstance(x,(int,float)):return (datetime(1899,12,30)+timedelta(days=x)).date()
    d=invoice_day(x)
    if d:return d
    for f in ('%d-%b-%Y','%d-%B-%Y'):
        try:return datetime.strptime(txt(x),f).date()
        except ValueError:pass
    return None
def iso(x):return str(day(x)) if day(x) else None
def money_match(a,b):return a is not None and b is not None and abs(a-b)<=2

audit=read(ROOT/'scripts/data/post-issue-20261029/audit.json')
ledger={ksrc(r['key']):dict(r,key=ksrc(r['key']),order='',buyer='') for r in audit['invoices']}
children=defaultdict(set);parents=defaultdict(set);replacement=defaultdict(set);edge_evidence={}
for e in audit['edges']:
    a,b=ksrc(e['old']),ksrc(e['new']);children[a].add(b);parents[b].add(a)
    if e['kind']=='replacement':replacement[a].add(b)
    edge_evidence[(a,b)]=e['sources']
file_controls=[]
raw_exact=defaultdict(set)
raw2025=sorted((ROOT/'data/Meinv/2025').glob('*.xls'))
raw2026=sorted((ROOT/'data/Meinv/2026').glob('*.xls'))
latest=sorted((ROOT/'data/Meinv_SAU_DIEU_CHINH_THAY_THE_20261101').glob('Bang_ke_hoa_don_da_su_dung_*.xls'))
cache_path=OUT/'source-index.json'
for p in ([] if cache_path.exists() else raw2025+raw2026+latest):
    count=0
    for rn,r in rows(p):
        if rn<6 or not txt(r.get('A')).isdigit():continue
        count+=1;k=kval(r.get('B'),r.get('C'))
        if not k:continue
        order=txt(r.get('F'));raw_exact[order].add(k) if order else None
        if k not in ledger or p in latest:
            ms=[amount(r.get(c)) for c in ('P','Q','R')]
            ledger[k]={'key':k,'date':iso(r.get('D')),'money':ms,'resolved_money':ms,'status':r.get('U'),'cqt_status':r.get('S'),'cqt_code_present':bool(r.get('E')),'source':{'file':str(p.relative_to(ROOT)),'row':rn},'issues':[],'order':order,'buyer':txt(r.get('L'))}
        else:
            ledger[k]['order']=order;ledger[k]['buyer']=txt(r.get('L'))
    file_controls.append({'file':str(p.relative_to(ROOT)),'rows':count})
# Bổ sung quan hệ phát hành mới nhất. Không thay số tiền trống bằng0.
for name,kind in [('replacements.json','replacement'),('adjustments.json','adjustment')]:
    p=ROOT/'scripts/data/post-final-issue-20261101'/name
    if not p.exists():continue
    for e in read(p):
        a,b=e.get('old') or e.get('root'),e.get('new') or e.get('invoice')
        if not isinstance(a,str) or not isinstance(b,str):continue
        children[a].add(b);parents[b].add(a)
        if kind=='replacement':replacement[a].add(b)

if cache_path.exists():
    cached=read(cache_path);ledger=cached['ledger'];file_controls=cached['files']
    for dest,keyname in [(children,'children'),(parents,'parents'),(replacement,'replacement'),(raw_exact,'raw_exact')]:
        dest.clear();dest.update({k:set(v) for k,v in cached[keyname].items()})
else:
    dump(cache_path,{'ledger':ledger,'files':file_controls,'children':{k:sorted(v) for k,v in children.items()},'parents':{k:sorted(v) for k,v in parents.items()},'replacement':{k:sorted(v) for k,v in replacement.items()},'raw_exact':{k:sorted(v) for k,v in raw_exact.items()}})

# Bảng đối soát cũ tập trung HĐ2024–T6/2025. Nối thêm đủ quan hệ H2/2025 và2026.
relation_files=[ROOT/'data/Du lieu Meinvoice/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_2025.xls',ROOT/'data/Du lieu Meinvoice/Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_2024_2025.xls']
relation_files+=list((ROOT/'data/Du lieu Meinvoice/Xuat_lai_2026-10-22').glob('*.xls'))
link_amounts=defaultdict(lambda:defaultdict(set))
for p in relation_files:
    if not p.exists():continue
    kind='replacement' if 'thay_the' in p.name else 'adjustment';count=0
    for rn,r in rows(p):
        if rn<7 or not txt(r.get('A')).isdigit():continue
        a,b=kval(r.get('B'),r.get('D')),kval(r.get('M'),r.get('O'))
        if not a or not b:continue
        count+=1;children[a].add(b);parents[b].add(a)
        if kind=='replacement':replacement[a].add(b)
        for k,cols in [(a,('J','K','L')),(b,('U','V','W'))]:
            for i,c in enumerate(cols):
                v=amount(r.get(c))
                if v is not None:link_amounts[k][i].add(float(v))
    file_controls.append({'file':str(p.relative_to(ROOT)),'rows':count,'kind':kind})
for k,fields in link_amounts.items():
    if k not in ledger:continue
    rec=ledger[k]
    for i,vs in fields.items():
        if rec['resolved_money'][i] is None and len(vs)==1:rec['resolved_money'][i]=next(iter(vs))

@lru_cache(None)
def roots(k):
    todo=[k];seen=set();result=set()
    while todo:
        x=todo.pop()
        if x in seen:continue
        seen.add(x)
        if parents[x]:todo.extend(parents[x])
        else:result.add(x)
    return frozenset(result)
@lru_cache(None)
def chain(k):
    todo=[k];seen=set();issues=[];active=[];excluded=[];unknown=[False]*3;sums=[0.0]*3
    while todo:
        x=todo.pop()
        if x in seen:continue
        seen.add(x);todo.extend(children[x]);r=ledger.get(x)
        if not r:issues.append('Thiếu HĐ trong bảng kê: '+x);unknown=[True]*3;continue
        st=norm(r.get('status'))
        if len(parents[x])>1:issues.append('HĐ có nhiều gốc: '+x)
        if replacement[x] or 'bi thay the' in st or 'huy' in st:
            excluded.append(x)
            if 'bi thay the' in st and not replacement[x]:issues.append('Thiếu liên kết thay thế: '+x);unknown=[True]*3
            continue
        active.append(x)
        if ('thay the' in st or 'dieu chinh' in st) and not parents[x] and x==k and ('hoa don thay the'==st or 'hoa don dieu chinh'==st):issues.append('Thiếu liên kết HĐ gốc: '+x)
        if r.get('issues'):issues.append('Nguồn chưa được xác minh đầy đủ: '+x)
        if 'khong hop le' in norm(r.get('cqt_status')) or 'tu choi' in norm(r.get('cqt_status')):issues.append('Trạng thái CQT chưa hợp lệ: '+x)
        for i,v in enumerate(r.get('resolved_money',r['money'])):
            if v is None:unknown[i]=True
            else:sums[i]+=float(v)
    balance=[None if unknown[i] else sums[i] for i in range(3)]
    return {'root':k,'keys':sorted(seen),'active':sorted(active),'balance':balance,'issues':sorted(set(issues)),'excluded':excluded}

amis=read(SRC/'plan25AMIS.json');kt=read(SRC/'plan25KT.json');live=read(SRC/'plan25MI.json')
nb=read(SRC/'plan25NB.json');pairs=read(SRC/'plan25Pairs.json')
ktorders=defaultdict(set);ktbykey=defaultdict(list)
for rn,r in enumerate(kt[1:],2):
    r=r+[None]*max(0,34-len(r));k=kval(r[9],r[10]);order=txt(r[3])
    if k:
        ktbykey[k].append({'row':rn,'order':order,'buyer':r[1],'action':r[5],'note':r[33],'balance':r[29:32]})
        if order and not order.startswith('#'):ktorders[order].update(roots(k))

# NB chỉ làm chứng cứ khi số/ngày/người mua/tiền khớp duy nhất meInvoice.
by_number_day=defaultdict(list)
for k,r in ledger.items():
    if r.get('date'):by_number_day[(k.split('/')[-1].lstrip('0') or '0',r['date'])].append(k)
nborders=defaultdict(set);nbproof=defaultdict(list)
for rn,r in enumerate(nb,2):
    r=r+[None]*max(0,20-len(r));order=txt(r[8]);nd=iso(r[0]);n=txt(r[2]).split('.')[0].lstrip('0') or '0'
    if not order or not nd:continue
    matches=[k for k in by_number_day[(n,nd)] if norm(ledger[k].get('buyer'))==norm(r[7]) and norm(r[7]) and money_match(float(ledger[k]['money'][2]) if ledger[k]['money'][2] is not None else None,num(r[19]))]
    if len(matches)==1:
        rr=roots(matches[0]);nborders[order].update(rr);nbproof[order].append({'row':rn,'key':matches[0]})

docs=[];outside=[]
amis_prefix=defaultdict(set)
for v in amis[1:]:
    o=txt(v[8]) if len(v)>8 else ''
    if re.fullmatch(r'\d{18}',o):amis_prefix[o[:15]].add(o)
def equivalent_order(a,b):
    if a==b:return True
    return bool(re.fullmatch(r'\d{18}',a or '') and re.fullmatch(r'\d{18}',b or '') and a[:15]==b[:15] and (a.endswith('000') or b.endswith('000')) and len(amis_prefix[a[:15]])==1)
for rn,v in enumerate(amis[1:],2):
    v=v+[None]*max(0,13-len(v));d=day(v[0])
    if not d or d.year!=2025:outside.append(rn);continue
    o=txt(v[8]);badorder=not o or o.startswith('#');k=kval(v[10],v[9]);rs=roots(k) if k else set()
    rawroots=set().union(*(roots(x) for x in raw_exact.get(o,set()))) if not badorder else set()
    hints=rawroots|nborders.get(o,set())|ktorders.get(o,set()) if not badorder else set()
    strong=rawroots|nborders.get(o,set()) if not badorder else set()
    chosen=next(iter(rs)) if len(rs)==1 and k in ledger else None
    proposed=None;link_issue=None
    if chosen and strong and chosen not in strong:
        if len(strong)==1:proposed=next(iter(strong));link_issue='Mã đơn nguồn trỏ HĐ khác HĐ AMIS'
        else:link_issue='Mã đơn nguồn có nhiều chuỗi khác HĐ AMIS'
    elif not chosen:
        if len(hints)==1:chosen=next(iter(hints));proposed=chosen;link_issue='Thiếu/sai khóa HĐ AMIS, tìm được một chuỗi theo mã đơn'
        elif len(hints)>1:link_issue='Mã đơn có nhiều chuỗi HĐ, chưa chọn được'
    evidence=[]
    if chosen in rawroots:evidence.append('Mã đơn HĐ gốc')
    if chosen in nborders.get(o,set()):evidence.append('NB khớp số/ngày/người mua/tiền')
    if chosen in ktorders.get(o,set()):evidence.append('Mã đơn Ktoan')
    if k in ledger:evidence.append('Ký hiệu + số HĐ AMIS')
    docs.append({'row':rn,'doc':txt(v[2]),'day':str(d),'serial':v[0],'month':d.month,'order':o,'money':num(v[7]),'vat':num(v[6]),'old_invoice':txt(v[9]),'old_symbol':txt(v[10]),'key':k,'root':chosen,'proposal':proposed,'hints':sorted(hints),'strong':sorted(strong),'evidence':evidence,'link_issue':link_issue,'note':txt(v[12]),'source_L':v[11]})

byroot=defaultdict(list);byorderdocs=defaultdict(list)
for d in docs:
    if d['root']:byroot[d['root']].append(d)
    if d['order'] and not d['order'].startswith('#'):byorderdocs[d['order']].append(d)

review_decisions={};review_pairs=[]
for rn,r in enumerate(pairs[6:],7):
    if not r or not isinstance(r[0],(int,float)):continue
    r=r+[None]*max(0,17-len(r));a=kval('1C25MUV',r[3]);b=kval('1C25MUV',r[4])
    if a and b:
        rec={'row':rn,'keep':a,'remove':b,'order':txt(r[1]),'money':r[5],'decision':txt(r[9]),'proof':txt(r[10])}
        review_pairs.append(rec);review_decisions[b]=rec

summary=defaultdict(lambda:{'docs':0,'money':0.0,'vat':0.0});reasons=Counter();monthly=defaultdict(lambda:defaultdict(lambda:{'docs':0,'money':0.0}));chain_rows=[]
nb_expected_byroot=defaultdict(set)
for o,rs in nborders.items():
    for r in rs:nb_expected_byroot[r].add(o)
for root,ds in byroot.items():
    ch=chain(root);bs=ch['balance'];known_am=all(x['money'] is not None and x['vat'] is not None for x in ds)
    amsum=sum(x['money'] for x in ds) if known_am else None;vatsum=sum(x['vat'] for x in ds) if known_am else None
    rec=ledger.get(root,{});rootday=day(rec.get('date'));protected=bool(rootday and rootday.year<2025)
    active_orders={ledger[x].get('order') for x in ch['keys'] if x in ledger and ledger[x].get('order')}
    assigned_orders={x['order'] for x in ds if x['order'] and not x['order'].startswith('#')}
    foreign_orders={o for o in active_orders if not any(equivalent_order(o,a) for a in assigned_orders)}
    chain_rows.append({'root':root,'docs':[x['doc'] for x in ds],'amis_money':amsum,'amis_vat':vatsum,'balance':bs,'keys':ch['keys'],'issues':ch['issues'],'protected_2024':protected})
    nb_expected=nb_expected_byroot[root]
    for d in ds:
        d.update(group_count=len(ds),group_money=amsum,group_vat=vatsum,balance=bs,active=ch['active'],chain_keys=ch['keys'],root_date=rec.get('date'),protected=protected)
        if any(equivalent_order(d['order'],o) for o in active_orders) and not any(x in d['evidence'] for x in ['Mã đơn HĐ gốc','NB khớp số/ngày/người mua/tiền','Mã đơn Ktoan']):
            d['evidence'].append('Mã nguồn trùng15số đầu, đuôi000; ứng viên duy nhất')
        if protected:
            cat='Cần KT xác minh';reason='HĐ thuộc chuỗi trước2025 đã chốt';action='Kiểm tra liên kết; giữ nguyên chuỗi2024'
        elif d['money'] is None or d['vat'] is None:
            cat='Cần KT xác minh';reason='AMIS thiếu số tiền hoặc VAT';action='Bổ sung số liệu AMIS'
        elif root in review_decisions and not review_decisions[root]['proof']:
            cat='Xử lý HĐ sai';reason='KT đã note bỏ HĐ trùng, chưa có bằng chứng xử lý';action='Thực hiện phương án KT chốt và lấy bảng kê sau xử lý'
        elif d['link_issue']:
            cat='Sửa liên kết';reason=d['link_issue'];action='Đối chiếu HĐ đề xuất và cập nhật liên kết sau khi KT chốt'
        elif len(ds)>1 and nb_expected and not any(equivalent_order(d['order'],o) for o in nb_expected) and not money_match(amsum,bs[2]):
            cat='Sửa liên kết';reason='Chứng từ ghép thêm khác mã NB và tổng nhóm lệch HĐ';action='Kiểm tra chứng từ ghép thêm; không sửa HĐ khi chưa xác minh'
        elif ch['issues'] or any(x is None for x in bs):
            cat='Cần KT xác minh';reason='Chuỗi thiếu tiền hoặc liên kết/trạng thái nguồn';action='Lấy bảng kê/quan hệ/XML để xác minh số dư'
        elif foreign_orders:
            cat='Sửa liên kết';reason='Chuỗi HĐ có mã đơn nguồn khác nhóm AMIS';action='Kiểm tra chứng từ nào thực sự thuộc HĐ'
        elif set(d['hints'])-{root}:
            cat='Cần KT xác minh';reason='Mã đơn còn chuỗi HĐ khác cần kiểm tra';action='So tất cả HĐ cùng mã trước khi giữ hoặc xuất bổ sung'
        elif all(abs(x)<=2 for x in bs) and amsum and amsum>2:
            cat='Xuất bổ sung';reason='Chuỗi đã về0, AMIS còn doanh thu';action='Ứng viên xuất lại: xác minh nghiệp vụ, liên kết và HĐ khác trước khi lập'
        elif not d['evidence'] or not any(x in d['evidence'] for x in ['Mã đơn HĐ gốc','NB khớp số/ngày/người mua/tiền','Mã đơn Ktoan','Mã nguồn trùng15số đầu, đuôi000; ứng viên duy nhất']):
            cat='Cần KT xác minh';reason='Chỉ khớp số HĐ, chưa xác minh mã đơn';action='Đối chiếu người mua/ngày/hàng hóa trước khi chốt'
        elif money_match(amsum,bs[2]) and money_match(vatsum,bs[1]) and money_match(amsum-vatsum,bs[0]):
            cat='Giữ HĐ';reason='Khớp mã đơn, VAT và thanh toán theo nhóm';action='Giữ HĐ theo đối soát tiền; kiểm chi tiết hàng hóa'
        else:
            cat='Xử lý HĐ sai';reason='Lệch tiền/VAT sau khi gộp đủ chứng từ';action='Kiểm tra chiết khấu, trả hàng/upgrade và nội dung HĐ'
        d.update(category=cat,reason=reason,action=action)

for d in docs:
    if d['root']:continue
    d.update(group_count=1,group_money=d['money'],group_vat=d['vat'],balance=[None,None,None],active=[],chain_keys=[],root_date=None,protected=False)
    if d['link_issue']:
        cat='Cần KT xác minh';reason=d['link_issue'];action='Đối chiếu các HĐ ứng viên, không lấy kết quả đầu tiên'
    elif not d['order'] or d['order'].startswith('#'):
        cat='Cần KT xác minh';reason='Mã đơn thiếu/lỗi và chưa có HĐ xác minh';action='Bổ sung mã đơn/chứng từ nguồn'
    elif d['money']==0:
        cat='Cần KT xác minh';reason='AMIS0đ, chưa xác định HĐ/nghiệp vụ hàng tặng';action='Kiểm tra chi tiết hàng và căn cứ xuất HĐ'
    else:
        cat='Cần KT xác minh';reason='Chưa tìm được HĐ trong nguồn hiện có';action='Ứng viên xuất thiếu, cần KT xác nhận chưa xuất hoặc xuất gộp'
    d.update(category=cat,reason=reason,action=action)

for d in docs:
    cat=d['category'];summary[cat]['docs']+=1;summary[cat]['money']+=d['money'] or 0;summary[cat]['vat']+=d['vat'] or 0
    reasons[(cat,d['reason'])]+=1;monthly[d['month']][cat]['docs']+=1;monthly[d['month']][cat]['money']+=d['money'] or 0

# Danh sách HĐ cần kiểm riêng: tránh bỏ sót HĐ trùng chưa gắn chứng từ AMIS.
linked_roots=set(byroot);invoice_review=[]
for root,rec in sorted(ledger.items()):
    rd=day(rec.get('date'))
    if not rd or rd.year!=2025 or parents[root] or root in linked_roots:continue
    ch=chain(root)
    if not ch['active'] or all(x is not None and abs(x)<=2 for x in ch['balance']):continue
    decision=review_decisions.get(root)
    invoice_review.append({'root':root,'date':rec['date'],'order':rec.get('order'),'buyer':rec.get('buyer'),'balance':ch['balance'],'state':rec['status'],'reason':'KT đã note bỏ, chưa có bằng chứng xử lý' if decision and not decision['proof'] else 'HĐ2025 còn giá trị, chưa nối chứng từ AMIS','decision':decision['decision'] if decision else '', 'pair_keep':decision['keep'] if decision else '', 'source':rec['source'],'issues':ch['issues'],'active':ch['active']})
reviewed_roots={x['root'] for x in invoice_review}
for root,decision in review_decisions.items():
    if root in reviewed_roots:continue
    rec=ledger.get(root,{})
    ch=chain(root)
    invoice_review.append({'root':root,'date':rec.get('date'),'order':rec.get('order'),'buyer':rec.get('buyer'),'balance':ch['balance'],'state':rec.get('status'),'reason':'KT đã note bỏ, cần đối chiếu bằng chứng xử lý','decision':decision['decision'],'pair_keep':decision['keep'],'source':rec.get('source'),'issues':ch['issues'],'active':ch['active']})

assert len({d['doc'] for d in docs})==len(docs)
assert sum(x['docs'] for x in summary.values())==len(docs)
assert abs(sum(x['money'] for x in summary.values())-sum(d['money'] or 0 for d in docs))<0.01
result={'as_of':'2026-11-05','source_url':'https://docs.google.com/spreadsheets/d/1bhOnhnBcHaHiy06-REAdVVEk160FG94a6XgkfH0AQsw/edit','source_cutoff':'Bảng kê và quan hệ29/10/2026, bổ sung phát hành01/11/2026; Google Sheets đọc mới05/11/2026','summary':dict(summary),'monthly':dict(monthly),'reasons':[{'category':k[0],'reason':k[1],'docs':v} for k,v in reasons.items()],'docs':docs,'chains':chain_rows,'invoice_review':invoice_review,'pairs':review_pairs,'files':file_controls,'outside_2025':outside,'ledger_count':len(ledger),'edge_count':sum(len(x) for x in children.values())}
dump(OUT/'classification.json',result)
print(json.dumps({'summary':result['summary'],'reasons':result['reasons'],'outside':len(outside),'invoice_review':len(invoice_review),'ledger':len(ledger),'pairs':len(review_pairs)},ensure_ascii=False))
