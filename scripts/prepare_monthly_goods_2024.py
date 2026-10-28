"""Thống kê hàng theo tháng từ AMIS và phương án hóa đơn, giữ phần chưa ghép riêng."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal
from datetime import datetime
import json,re,unicodedata,hashlib
from prepare_invoice_reports import rows,invoice_key,amount
ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/monthly-goods-20261028'
WORK.mkdir(parents=True,exist_ok=True)
def load(p):return json.loads((ROOT/p).read_text(encoding='utf8'))
def dec(v):return Decimal(str(v)) if v not in (None,'') else Decimal(0)
def native(v):
    if isinstance(v,Decimal):return int(v) if v==int(v) else float(v)
    raise TypeError(type(v))
def norm(v):return ' '.join(re.sub(r'[^A-Z0-9]+',' ',''.join(c for c in unicodedata.normalize('NFD',str(v or '').upper().replace('Đ','D')) if not unicodedata.combining(c))).split())
def signature(name,code=''):
    s=norm(name);bases=set(re.findall(r'(?:KA|KM|KK|UK)\s*\d+',s))
    bases={x.replace(' ','') for x in bases}
    if code:bases|=set(re.findall(r'(?:KA|KM|KK|UK)\d+',code.upper()))
    if len(bases)!=1 or 'COMBO' in s:return None
    kind='filter' if ('FILTER' in code.upper() or re.match(r'^(MANG LOC|LOI LOC|FILTER|BO LOC)',s)) else 'machine'
    return (next(iter(bases)),kind)
def month(v):return datetime.fromisoformat(v[:19]).month
f=load('scripts/data/revenue-forecast-20261028/prepared.json')
p=load('scripts/data/amis-2024-linked-20261026/prepared.json')
gap=load('scripts/data/revenue-gap-20261028/prepared.json')
probe=load('scripts/data/monthly-goods-20261028/probe.json')
docs={d['doc']:d for d in p['docs']}
for m in p['manifest']:assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256'],m['file']
processed={tuple(r['root']) for r in f['adjustment_records']}
manual={tuple(r['root']) for r in f['manual_records']}
all_by_root=defaultdict(list)
for r in f['old_invoices']:all_by_root[tuple(r['root'])].append(r)
intact={k:v for k,v in all_by_root.items() if k not in processed|manual}
assert len(intact)==117
needed={tuple(r['key']):r for rs in intact.values() for r in rs}
details=probe['details']
# Bổ sung hóa đơn0đ vẫn giữ, không bỏ khỏi thống kê số lượng chỉ vì tiền bằng0.
missing={k for k in needed if '/'.join(k) not in details}
if missing:
    for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2024.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025.xls']:
        src=ROOT/'data/Du lieu Meinvoice'/name
        for rn,r in rows(src):
            k=invoice_key(r.get('D'),r.get('B'))
            if rn>=7 and k in missing:
                details.setdefault('/'.join(k),[]).append({'row':rn,'source':str(src.relative_to(ROOT)),'code':r.get('K'),'name':r.get('L'),'unit':r.get('M'),'qty':r.get('N'),'price':r.get('O'),'goods':r.get('P'),'discount':r.get('R'),'vat':r.get('T'),'payment':r.get('U')})
        missing={k for k in missing if not details.get('/'.join(k))}
        if not missing:break
ledger=[];pred=[];unclear=[];cost_pending=[];mapping=[];sku_names={};sku_units={}
def rec(m,code,name,unit,qty,revenue,cost,doc,kind,source='',cost_missing=False,note=''):
    return {'month':m,'code':code,'name':name,'unit':unit or '', 'qty':dec(qty),'revenue':dec(revenue),'cost':None if cost is None else dec(cost),'doc':doc,'kind':kind,'source':source,'cost_missing':cost_missing,'note':note}
for d in p['docs']:
    for l in d['original_lines']:
        r=l['values'];code=r[12];sku_names.setdefault(code,r[13] or r[7]);sku_units.setdefault(code,r[16])
        q=dec(r[17])-dec(r[26]);rev=dec(r[20])-dec(r[27]);m=month(d['date'])
        assert r[39] is not None,('Giá vốn nguồn trống',d['doc'],l['row'])
        x=rec(m,code,sku_names[code],r[16],q,rev,r[39],d['doc'],'AMIS',f'Sổ AMIS | dòng {l["row"]}')
        ledger.append(x)
        if d['category']=='Đưa vào import':pred.append({**x,'kind':'Bán mới'})
        elif d['doc'].startswith('BTL'):pred.append({**x,'kind':'Trả hàng'})
assert sum(r['revenue'] for r in ledger)==f['totals']['amis_net'][0]
assert sum(r['revenue'] for r in pred)==f['totals']['new_gross'][0]+f['totals']['new_returns'][0]
root_group={rt:g for g in gap['groups'] for rt in g['roots']}
for rt,inv in sorted(intact.items()):
    root='/'.join(rt);g=root_group[root]
    candidates={l['values'][12] for c in g['codes'] if c in docs for l in docs[c]['original_lines']}
    group_units={code:{norm(l['values'][16]) for c in g['codes'] if c in docs for l in docs[c]['original_lines'] if l['values'][12]==code} for code in candidates}
    root_month=month(inv[0]['root_date'])
    def match_line(l):
        if l['code'] in candidates:return l['code'],'Mã hàng trùng trực tiếp'
        exact={code for code in candidates if norm(l['name']) in {norm(sku_names[code])}|{norm(x['values'][7]) for c in g['codes'] if c in docs for x in docs[c]['original_lines'] if x['values'][12]==code}}
        if len(exact)==1:return next(iter(exact)),'Tên hàng chuẩn hóa trùng, trong BHH đã xác nhận'
        if 'COMBO' in norm(l['name']) or len(set(re.findall(r'(?:KA|KM|KK|UK)\s*\d+',norm(l['name']))))>1:return None,'Dòng combo gồm nhiều hàng, chưa tự tách'
        sig=signature(l['name'])
        compatible={c for c in candidates if signature(sku_names[c],c)==sig} if sig else set()
        if len(compatible)==1:return next(iter(compatible)),'Model và loại hàng duy nhất trong BHH đã xác nhận'
        if len(candidates)==1 and norm(l['unit']) in group_units[next(iter(candidates))]:return next(iter(candidates)),'BHH đã xác nhận chỉ có một mã hàng, cùng ĐVT'
        return None,'Chưa có mã hàng duy nhất trong BHH đã xác nhận'
    physical=[]
    for h in inv:
        for l in details.get('/'.join(h['key']),[]):
            if dec(l['qty'])==0 and dec(l['goods'])-dec(l['discount'])==0:continue
            code,basis=match_line(l);physical.append((h,l,code,basis))
    mapped_root={code for h,l,code,basis in physical if code}
    for h in inv:
        key='/'.join(h['key']);ls=details.get(key,[]);detail_total=sum(dec(l['goods'])-dec(l['discount']) for l in ls)
        for ix,l in enumerate(ls):
            qty=dec(l['qty']);rev=dec(l['goods'])-dec(l['discount'])
            if qty==0 and rev==0:continue
            code,basis=match_line(l)
            unresolved=code is None
            unit_mismatch=bool(code and norm(l['unit']) not in group_units[code])
            if unit_mismatch:basis+='; ĐVT trên HĐ khác AMIS, chưa quy đổi'
            actual_code=code or 'HD_CHUA_GHEP_'+rt[1]+'_'+str(l['row'])
            unit={'CAI':'Cái','CHIEC':'Chiếc','BO':'Bộ','SOI':'Sợi','LAN':'Lần'}.get(norm(l['unit']),l['unit'])
            x=rec(root_month,actual_code,sku_names.get(code,l['name']),unit,qty,rev,None if unresolved or unit_mismatch else 0,key,'HĐ cũ giữ',l['source']+' | dòng '+str(l['row']),cost_missing=unresolved or unit_mismatch,note=basis)
            pred.append(x)
            mapping.append({'root':root,'invoice':key,'bhh':g['codes'],'source_name':l['name'],'mapped_code':code,'basis':basis,'qty':qty,'revenue':rev})
            if unresolved:unclear.append({'kind':'Mã hàng/Combo','month':root_month,'invoice':key,'root':root,'bhh':g['codes'],'item':l['name'],'qty':qty,'revenue':rev,'reason':basis,'source':x['source']})
            elif unit_mismatch:unclear.append({'kind':'ĐVT khác AMIS','month':root_month,'invoice':key,'root':root,'bhh':g['codes'],'item':l['name'],'qty':qty,'revenue':rev,'reason':'ĐVT HĐ: '+str(l['unit'])+'; ĐVT BHH: '+', '.join(sorted(group_units[code]))+'. Chưa tự quy đổi hoặc gán giá vốn vào ĐVT khác.','source':x['source']})
        delta=dec(h['money'][0])-detail_total
        if delta:
            code=next(iter(mapped_root)) if len(mapped_root)==1 else next(iter(candidates)) if len(candidates)==1 else None
            actual_code=code or 'DIEU_CHINH_TIEN_CHUA_PHAN_BO'
            x=rec(root_month,actual_code,sku_names.get(code,'Điều chỉnh tiền chưa phân bổ theo mã hàng'),sku_units.get(code,''),0,delta,0,key,'Điều chỉnh tiền HĐ cũ',h['source']+' | dòng '+str(h.get('row','')),note='Theo tổng tiền HĐ; chi tiết nguồn không có khoản tiền này')
            pred.append(x)
            if code is None:unclear.append({'kind':'Phân bổ tiền','month':root_month,'invoice':key,'root':root,'bhh':g['codes'],'item':x['name'],'qty':None,'revenue':delta,'reason':'Hóa đơn điều chỉnh chỉ có tổng tiền; gốc có nhiều mã hàng, chưa tự phân bổ','source':x['source']})
    assert sum(r['revenue'] for r in pred if r['doc'] in {'/'.join(h['key']) for h in inv})==sum(dec(h['money'][0]) for h in inv),root
# Giá vốn giữ lại lấy từng dòng BHH AMIS, mỗi BHH chỉ dùng một lần trong một nhóm liên kết.
used_cost_docs=set()
for g in gap['groups']:
    roots=[rt for rt in g['roots'] if tuple(rt.split('/')) in intact]
    if not roots:continue
    months={month(intact[tuple(rt.split('/'))][0]['root_date']) for rt in roots}
    assert len(months)==1,('Nhóm có nhiều tháng gốc, chưa phân bổ GV',roots)
    m=next(iter(months));cost_docs=[c for c in g['codes'] if c in docs]
    if not cost_docs:
        for rt in roots:
            cost_pending.append({'kind':'Thiếu giá vốn','month':m,'invoice':rt,'root':rt,'bhh':g['codes'],'item':'Giá vốn hàng thuộc HĐ cũ giữ lại','qty':None,'revenue':sum(dec(h['money'][0]) for h in intact[tuple(rt.split('/'))]),'reason':'Chưa có BHH năm 2024 để lấy giá vốn; không mặc định bằng 0','source':'data/HD Xuất CTY 2024.xlsx | HD cty 2024 | dòng '+','.join(map(str,g['exclusion_rows']))})
        continue
    for code in cost_docs:
        assert code not in used_cost_docs,('Trùng giá vốn BHH',code)
        used_cost_docs.add(code)
        for l in docs[code]['original_lines']:
            r=l['values'];sku=r[12]
            pred.append(rec(m,sku,sku_names[sku],r[16],0,0,r[39],code,'Giá vốn HĐ cũ giữ',f'Sổ AMIS | dòng {l["row"]}',note='Giá vốn BHH đã xác nhận; không phải lấy từ giá bán HĐ'))
assert len(used_cost_docs)==115
assert sum(r['revenue'] for r in pred)==f['totals']['forecast_after_manual'][0]
assert sum((r['cost'] or 0) for r in ledger)==24542992457
assert sum((r['cost'] or 0) for r in pred)==24542992457
unknown=unclear+cost_pending
stats=defaultdict(lambda:{'qty_amis':Decimal(0),'rev_amis':Decimal(0),'cost_amis':Decimal(0),'qty_invoice':Decimal(0),'rev_invoice':Decimal(0),'cost_invoice':Decimal(0),'cost_invoice_complete':True,'notes':set(),'name':'','unit':''})
for side,rs in [('amis',ledger),('invoice',pred)]:
    for r in rs:
        v=stats[(r['month'],r['code'],r['unit'])];v['name']=r['name'];v['unit']=r['unit'];v['qty_'+side]+=r['qty'];v['rev_'+side]+=r['revenue'];v['cost_'+side]+=r['cost'] or 0
        if side=='invoice' and r['cost_missing']:v['cost_invoice_complete']=False
        if r['note'] and (r['code'].startswith('HD_') or r['code'].startswith('DIEU_') or 'ĐVT' in r['note']):v['notes'].add(r['note'])
monthly=[]
for m in range(1,13):
    a=[r for r in ledger if r['month']==m];b=[r for r in pred if r['month']==m];pending=[r for r in cost_pending if r['month']==m]
    monthly.append({'month':m,'rev_amis':sum(r['revenue'] for r in a),'rev_invoice':sum(r['revenue'] for r in b),'cost_amis':sum(r['cost'] or 0 for r in a),'cost_invoice_known':sum(r['cost'] or 0 for r in b),'missing_cost_invoices':len(pending),'unmapped_rows':sum(r['month']==m for r in unclear)})
result={'basis':'Bán mới/trả theo ngày chứng từ AMIS; hóa đơn cũ theo tháng HĐ gốc 2024. Không phải báo cáo ngày phát hành thực tế.','forecast_revenue':f['totals']['forecast_after_manual'][0],'amis_revenue':f['totals']['amis_net'][0],'known_cost':24542992457,'monthly':monthly,'stats':[{'month':m,'code':code,**{k:sorted(v) if isinstance(v,set) else v for k,v in x.items()}} for (m,code,unit),x in sorted(stats.items())],'pending':unknown,'cost_pending':cost_pending,'mapping':mapping,'sources':f['manifest'],'counts':{'amis_lines':len(ledger),'predicted_lines':len(pred),'intact_roots':len(intact),'intact_invoices':len(needed),'cost_mapped_bhh':len(used_cost_docs),'unmapped_item_rows':sum(r['kind']=='Mã hàng/Combo' for r in unclear),'unallocated_money_rows':sum(r['kind']=='Phân bổ tiền' for r in unclear),'unit_mismatch_rows':sum(r['kind']=='ĐVT khác AMIS' for r in unclear)}}
(WORK/'prepared.json').write_text(json.dumps(result,ensure_ascii=False,default=native),encoding='utf8')
print(json.dumps({'counts':result['counts'],'stats':len(result['stats']),'pending':len(unknown),'cost_pending':cost_pending,'unallocated_goods':sum(r['revenue'] for r in unclear if r['kind']=='Phân bổ tiền'),'forecast_revenue':result['forecast_revenue'],'known_cost':result['known_cost']},ensure_ascii=False,default=native))
