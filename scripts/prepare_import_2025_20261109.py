"""Prepare 2025 import decisions from fresh AMIS and read-only Sheets snapshots."""
from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
import json, sys, hashlib, re
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/import-2025-20261109'; SRC=OUT/'source'
sys.path.insert(0,str(ROOT/'scripts'))
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def save(p,v):p.write_text(json.dumps(v,ensure_ascii=False,default=str),encoding='utf-8')
def txt(v):return '' if v is None else str(v).strip()
def num(v):return Decimal(str(v or 0))
def import_quantity(line):return num(line.get('import_quantity',line['v'][22]))
def chain_is_verified_zero(ch):
    balance=ch.get('balance')
    return not ch.get('issues') and balance is not None and len(balance)==3 and all(v is not None and txt(v)!='' and num(v)==0 for v in balance)
def import_rate_candidates(goods,vat,promotion=False,confirmed_rate=None):
    if promotion and goods==0 and vat==0 and confirmed_rate in [0,5,8,10]:return [confirmed_rate]
    return [rate for rate in [0,5,8,10] if abs((goods*num(rate)/100).quantize(Decimal(1),rounding=ROUND_HALF_UP)-vat)<=1]
def native(v):
    if isinstance(v,Decimal):return int(v) if v==int(v) else float(v)
    if isinstance(v,(date,datetime)):return v.isoformat()[:10]
    if isinstance(v,dict):return {k:native(x) for k,x in v.items()}
    if isinstance(v,(list,tuple)):return [native(x) for x in v]
    return v

def main():
    # Reuse the existing read-only chain parser, with a dedicated output/cache directory.
    original=ROOT/'scripts/classify_2025_20261105.py'
    prefix=original.read_text(encoding='utf-8').split("amis=read(SRC/'plan25AMIS.json')",1)[0]
    prefix=prefix.replace("OUT=ROOT/'outputs/phan-loai-2025-20261105'",'OUT=Path('+repr(str(OUT))+')')
    ns={'__file__':str(original),'__name__':'chain_reader'}
    exec(compile(prefix,str(original),'exec'),ns)
    ledger,roots,chain,kval=ns['ledger'],ns['roots'],ns['chain'],ns['kval']
    print(json.dumps({'stage':'chains','invoices':len(ledger)},ensure_ascii=False),flush=True)
    amis=read(SRC/'amis.json')['values']; live={txt(r[2]):{'row':i,'v':r+[None]*max(0,16-len(r))} for i,r in enumerate(amis[1:],2) if len(r)>2 and r[2]}
    original_codes=read(SRC/'original_0.json')['values'];original_links=read(SRC/'original_1.json')['values']
    original_map={txt(r[0]):(original_links[i]+[None]*3)[:3] for i,r in enumerate(original_codes) if r and r[0] and i<len(original_links)}
    kt=[]
    for p in sorted(SRC.glob('kt_*.json'),key=lambda p:int(p.stem.split('_')[1])):kt+=read(p)['values']
    kt_byorder=defaultdict(set); kt_bykey={}
    for i,r in enumerate(kt[1:],2):
        r=r+[None]*max(0,34-len(r));k=kval(r[9],r[10])
        if k:
            kt_bykey[k]={'row':i,'values':r}
            if r[3] and not txt(r[3]).startswith('#'):kt_byorder[txt(r[3])].update(roots(k))
    raw_orders=defaultdict(set)
    for k,r in ledger.items():
        if r.get('order'):raw_orders[txt(r['order'])].update(roots(k))
    company_docs=set(); company_keys=set(); company_rows=[]
    for i,r in enumerate(read(SRC/'company.json')['values'],1):
        if i<5 or len(r)<4 or not r[3]:continue
        r=r+[None]*max(0,28-len(r));k=kval(txt(r[1])+txt(r[2]),r[3])
        if k:company_keys.update(roots(k))
        codes=re.findall(r'(?:BHH[A-Z]*|BHDV)\d+',txt(r[16]))
        company_docs.update(codes);company_rows.append({'row':i,'key':k,'docs':codes})
    docs={};manifest=[]
    for name in ['Ban_hang.xlsx','So_chi_tiet_ban_hang.xlsx']:
        p=ROOT/'data/0811'/name;manifest.append({'file':str(p.relative_to(ROOT)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
    wb=load_workbook(ROOT/'data/0811/Ban_hang.xlsx',read_only=True,data_only=True)
    for ws in wb:
        for rn,r in enumerate(ws.iter_rows(min_row=4,values_only=True),4):
            if not isinstance(r[1],(date,datetime)) or not r[2]:continue
            code=txt(r[2]);assert code not in docs,code
            docs[code]={'doc':code,'date':r[1].date().isoformat() if isinstance(r[1],datetime) else r[1].isoformat(),'customer':txt(r[3]),'buyer':txt(r[4]),'vat':num(r[6]),'pay':num(r[7]),'source':ws.title+'!'+str(rn),'lines':[]}
    wb.close()
    returns=[];headers=None
    wb=load_workbook(ROOT/'data/0811/So_chi_tiet_ban_hang.xlsx',read_only=True,data_only=True)
    for ws in wb:
        h=list(next(ws.iter_rows(min_row=4,max_row=4,values_only=True)));assert h[15]=='Mã hàng' and h[41]=='Thuế GTGT';headers=h
        for rn,r in enumerate(ws.iter_rows(min_row=5,values_only=True),5):
            if not r[2] or not isinstance(r[0],(date,datetime)):continue
            code=txt(r[2]);line={'source':ws.title+'!'+str(rn),'v':list(r)}
            if code in docs:docs[code]['lines'].append(line)
            elif code.startswith('BTL'):returns.append({'doc':code,**line})
    wb.close()
    print(json.dumps({'stage':'amis','docs':len(docs),'lines':sum(len(d['lines']) for d in docs.values()),'returns':len(returns)}),flush=True)
    service_path=OUT/'service-quantity-confirmations.json'
    service_corrections={r['doc']:r for r in read(service_path)['records']} if service_path.exists() else {}
    rounding_path=OUT/'rounding-confirmation.json'
    rounding_policy=read(rounding_path) if rounding_path.exists() else None
    amount_tolerance=num(rounding_policy['max_abs_delta_vnd']) if rounding_policy else Decimal(1)
    assert Decimal(1)<=amount_tolerance<=Decimal(2)
    if rounding_policy:manifest.append({'file':str(rounding_path.relative_to(ROOT)),'sha256':hashlib.sha256(rounding_path.read_bytes()).hexdigest()})
    if service_corrections:
        manifest.append({'file':str(service_path.relative_to(ROOT)),'sha256':hashlib.sha256(service_path.read_bytes()).hexdigest()})
        for code,d in docs.items():
            if code.startswith('BHDV'):d['business_type']='Dịch vụ'
        for code,c in service_corrections.items():
            d=docs[code];assert code.startswith('BHDV') and len(d['lines'])==1 and d['business_type']=='Dịch vụ'
            line=d['lines'][0];r=line['v']
            assert line['source']==c['source'] and txt(r[15])==c['sku'] and txt(r[19])==c['unit']
            assert num(r[22])==num(c['source_quantity']) and num(r[23])==num(c['price']) and num(r[24])==num(c['goods'])
            assert num(r[41])==num(c['vat']) and num(r[43])==num(c['pay'])
            line['import_quantity']=num(c['import_quantity'])
            line['quantity_confirmation']={'source_quantity':num(r[22]),'import_quantity':num(c['import_quantity']),'file':str(service_path.relative_to(ROOT))}
    # Read original 2025 invoice detail, keeping exact item names and units.
    detail=defaultdict(list)
    from prepare_invoice_reports import rows
    for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
        mt='_MTT' in name;p=ROOT/'data/Du lieu Meinvoice'/name
        for rn,r in rows(p):
            if rn<7:continue
            k=kval(r.get('F' if mt else 'D'),r.get('B'))
            if not k:continue
            detail[k].append({'name':txt(r.get('S' if mt else 'L')),'unit':txt(r.get('T' if mt else 'M')),'qty':r.get('U' if mt else 'N'),'price':r.get('V' if mt else 'O'),'goods':r.get('W' if mt else 'P'),'vat':r.get('AA' if mt else 'T'),'discount':r.get('Y' if mt else 'R'),'rate':r.get('Z' if mt else 'S'),'source':name+'!'+str(rn)})
    def items_amis(d):
        z=defaultdict(Decimal)
        for l in d['lines']:
            r=l['v'];z[(ns['norm'](r[7] or r[16]),ns['norm'](r[19]))]+=import_quantity(l)
        return dict(z)
    def items_mi(ch):
        z=defaultdict(Decimal)
        for k in ch['active']:
            if k not in detail:return None
            for l in detail[k]:
                if l['qty'] is None:return None
                z[(ns['norm'](l['name']),ns['norm'](l['unit']))]+=num(l['qty'])
        return {k:v for k,v in z.items() if v}
    review_path=OUT/'h1-12-invoice-link-review.json'
    reviewed={r['doc']:r for r in read(review_path)['records']} if review_path.exists() else {}
    if reviewed:manifest.append({'file':str(review_path.relative_to(ROOT)),'sha256':hashlib.sha256(review_path.read_bytes()).hexdigest()})
    decisions=[];chains={};doc_roots=defaultdict(set);groups=defaultdict(list)
    for code,d in docs.items():
        lv=live.get(code);v=lv['v'] if lv else [None]*16
        if not lv and code in original_map:
            v[8:11]=original_map[code]
        k=kval(v[10],v[9]);rs=set(roots(k)) if k else set()
        if not rs:
            for l in d['lines']:
                r=l['v'];lk=kval(r[3],r[5])
                if lk:rs.update(roots(lk))
        order=txt(v[8])
        if order and not order.startswith('#'):rs.update(raw_orders.get(order,set()) or kt_byorder.get(order,set()))
        if code in reviewed:
            evidence=reviewed[code]
            assert d['date']<'2025-07-01' and order==evidence['order']
            assert evidence['confirmed_invoice'] is None and evidence['amis_symbol'] is None and evidence['amis_invoice_number'] is None
            assert not ns['raw_exact'].get(order), 'New invoice evidence requires another review'
            d['unverified_snapshot_roots']=sorted(rs)
            # These links were checked on AMIS and against source order IDs.
            # Preserve them as unverified evidence, not actionable invoice roots.
            rs=set()
        for root in rs:groups[root].append(code)
        doc_roots[code]=rs
        d['order']=order if order and not order.startswith('#') else None
    for root in set().union(*doc_roots.values()):
        ch=chain(root)
        kr=kt_bykey.get(root)
        if kr and any(v is None for v in ch['balance']):
            kb=kr['values'][29:32]
            if all(isinstance(x,(int,float)) for x in kb) and all(a is None or num(a)==num(b) for a,b in zip(ch['balance'],kb)):
                ch['balance']=kb;ch['balance_evidence']='Ktoan bản sao dòng '+str(kr['row'])+' AD:AF; khớp các khoản nguồn có giá trị'
        chains[root]=ch
    for code,d in docs.items():
        lv=live.get(code);v=lv['v'] if lv else [None]*16;rs=doc_roots[code];reasons=[]
        goods=sum((num(l['v'][24])-num(l['v'][29])-num(l['v'][36])-num(l['v'][37]) for l in d['lines']),Decimal(0))
        assert sum((num(l['v'][41]) for l in d['lines']),Decimal(0))==d['vat'],code
        assert sum((num(l['v'][43]) for l in d['lines']),Decimal(0))==d['pay'],code
        d['goods']=goods
        is_company=code in company_docs or bool(rs&company_keys)
        h1=d['date']<'2025-07-01';action='new' if h1 else 'review';content='not_checked'
        if is_company:action='company'
        elif h1:reasons.append('Tháng 1–6: tạo mới theo yêu cầu')
        elif not lv:reasons.append('Chứng từ không có trên AMIS bản sao')
        elif len(rs)==1:
            ch=chains[next(iter(rs))];bal=ch['balance']
            if not ch['issues'] and all(x is not None for x in bal):
                same=all(num(a)==num(b) for a,b in zip(bal,[goods,d['vat'],d['pay']]))
                if same:
                    mi_items=items_mi(ch);content='matched' if mi_items==items_amis(d) else 'different' if mi_items is not None else 'missing_detail'
                    # Existing sheet match is preserved; content discrepancies are explicit, not silently ignored.
                    action='keep'
                    if content!='matched':reasons.append('Giữ đơn đã khớp đủ tiền theo yêu cầu; chi tiết '+content+' cần kiểm riêng, không tự xuất lại')
                else:action='new';reasons.append('Số dư chuỗi khác AMIS')
            else:reasons+=ch['issues'] or ['Thiếu tiền chuỗi']
        elif not rs:
            action='new';reasons.append('Chưa xác định được hóa đơn liên quan')
        else:reasons.append('Có nhiều chuỗi hóa đơn ứng viên')
        if code in reviewed:reasons.append('Đã rà AMIS live và mã đơn: chưa xác định HĐ thực tế; không dùng liên kết Sheet chưa xác minh để xử lý HĐ cũ. T1–T6 vẫn tạo mới theo yêu cầu')
        if code in service_corrections:reasons.append('Dịch vụ BHDV; SL import '+str(service_corrections[code]['import_quantity'])+' theo xác nhận người dùng, giữ nguyên ĐVT/đơn giá/tiền hàng/VAT/thanh toán nguồn')
        rec={'doc':code,'date':d['date'],'action':action,'roots':sorted(rs),'live_row':lv['row'] if lv else None,'sheet_delta':v[12],'content':content,'goods':goods,'vat':d['vat'],'pay':d['pay'],'reasons':reasons}
        decisions.append(rec)
    # Never zero a shared chain while retaining any constituent invoice amount.
    bydoc={x['doc']:x for x in decisions};zero=[]
    for root,codes in groups.items():
        ch=chains[root];selected=[c for c in codes if bydoc[c]['action']=='new']
        if not selected:continue
        # A zeroed chain needs no further mutation. Shared-document guards only
        # apply when we would actually issue another invoice to zero the chain.
        if chain_is_verified_zero(ch):continue
        reason=[]
        if root in company_keys:reason.append('Chuỗi thuộc danh sách công ty')
        if ch['issues']:reason+=ch['issues']
        if any(b is None for b in ch['balance']):reason.append('Số dư chuỗi còn ô trống')
        if any(bydoc[c]['action']!='new' for c in codes):reason.append('Chuỗi dùng chung với chứng từ giữ lại/chờ/loại')
        for c in selected:
            if len(doc_roots[c])!=1:reason.append('Chứng từ có nhiều chuỗi')
        kind='adjustment' if any(ns['parents'][k] and k not in set().union(*ns['replacement'].values()) for k in ch['keys']) else 'replacement' if len(ch['keys'])==1 else 'review'
        if kind=='review':reason.append('Chuỗi thay thế phức tạp')
        if reason:
            for c in selected:bydoc[c]['action']='review';bydoc[c]['reasons']+=list(dict.fromkeys(reason))
        zero.append({'root':root,'docs':sorted(set(codes)),'kind':kind,'balance':ch['balance'],'keys':ch['keys'],'active':ch['active'],'reasons':list(dict.fromkeys(reason))})
    # Eligible new invoices must have a representable rate, no unexplained amounts, and full source lines.
    promotion_path=OUT/'promotion-confirmations.json'
    promotions={r['doc']:r for r in read(promotion_path)['records']} if promotion_path.exists() else {}
    if promotions:manifest.append({'file':str(promotion_path.relative_to(ROOT)),'sha256':hashlib.sha256(promotion_path.read_bytes()).hexdigest()})
    newrows=[];new_audit=[]
    for rec in decisions:
        if rec['action']!='new':continue
        d=docs[rec['doc']];rates=set();reasons=[];confirmed=promotions.get(rec['doc'])
        if confirmed:
            assert len(d['lines'])==1 and rec['goods']==rec['vat']==rec['pay']==0
            sr=d['lines'][0]['v']
            assert txt(sr[18])=='✓' and txt(sr[15])==confirmed['sku'] and num(sr[22])==num(confirmed['qty'])
            assert num(sr[23])==num(sr[24])==num(sr[41])==0 and confirmed['promotion'] is True
            rec['reasons'].append('Hàng khuyến mại được người dùng xác nhận; cờ KM và thuế suất '+str(confirmed['rate'])+'% đã đọc trên AMIS live; giữ tiền bằng 0')
        if not d['lines']:reasons.append('Thiếu dòng hàng AMIS')
        for l in d['lines']:
            r=l['v'];g=num(r[24])-num(r[29]);vat=num(r[41]);possible=import_rate_candidates(g,vat,txt(r[18])=='✓',confirmed['rate'] if confirmed else None)
            if len(possible)!=1:reasons.append('Chưa xác định duy nhất thuế suất')
            else:rates.add(possible[0])
            if num(r[36]) or num(r[37]):reasons.append('Có trả lại/giảm giá trên dòng bán')
            if num(r[29]):reasons.append('Có chiết khấu cần mẫu phù hợp')
            if not r[7] and not r[16]:reasons.append('Thiếu tên hàng')
            if import_quantity(l)<=0:reasons.append('Số lượng không dương')
            if abs((import_quantity(l)*num(r[23])).quantize(Decimal(1),rounding=ROUND_HALF_UP)-g)>amount_tolerance:reasons.append('Đơn giá × số lượng khác thành tiền quá '+str(amount_tolerance)+'đ')
        if len(rates)!=1:reasons.append('Thuế suất chứng từ không duy nhất')
        if rec['goods']+rec['vat']!=rec['pay']:reasons.append('Tiền hàng + VAT khác thanh toán')
        if reasons:rec['action']='review';rec['reasons']+=list(dict.fromkeys(reasons));continue
        if rounding_policy and rec['doc'] in rounding_policy['newly_eligible_documents']:rec['reasons'].append('Chấp nhận chênh số lượng × đơn giá tối đa 2đ theo xác nhận người dùng; giữ nguyên SL/đơn giá/tiền hàng/VAT/thanh toán nguồn')
        n=len(new_audit)+1;rate=next(iter(rates));new_audit.append({'number':n,'doc':d['doc'],'first_row':len(newrows)+10,'rows':len(d['lines'])})
        for j,l in enumerate(d['lines']):
            r=l['v'];row=[n,None,d['buyer'],txt(r[13]) or None,txt(r[12]) or None,txt(r[14]) or None,None,None,None,'TM/CK',rate,d['vat'] if j==0 else None,txt(r[7] or r[16]),txt(r[19]),import_quantity(l),num(r[23]),num(r[24]),d['order'],d['doc'],txt(r[15]),d['date'],l['source']]
            if promotions:row.append(1 if confirmed else None)
            newrows.append(row)
    # A later detail gate also blocks its old chain from zeroing.
    for z in zero:
        if not z['reasons'] and any(bydoc[c]['action']!='new' for c in z['docs']):z['reasons'].append('Chứng từ xuất mới chưa đủ dữ liệu import')
    result={'manifest':manifest,'decisions':decisions,'new_headers':['Số thứ tự hóa đơn (*)','Ngày hóa đơn','Tên đơn vị mua hàng','Mã số thuế','Địa chỉ','Người mua hàng','Email','Số điện thoại','Căn cước công dân','Hình thức thanh toán (*)','Thuế suất GTGT (%)','Tiền thuế GTGT','Tên hàng hóa/dịch vụ (*)','ĐVT','Số lượng','Đơn giá','Thành tiền','Mã đơn hàng','Số chứng từ','Mã hàng','Ngày chứng từ AMIS','Dòng nguồn'],'new_rows':newrows,'new_audit':new_audit,'zero':zero,'chains':chains,'company_rows':company_rows,'sheet_only':sorted(set(live)-set(docs)),'returns':returns,'source_headers':headers,'docs':docs,'invoice_detail':dict(detail)}
    if promotions:result['new_headers'].append('Hàng khuyến mại')
    save(OUT/'prepared.json',native(result))
    summary={'docs':len(docs),'actions':dict(Counter(r['action'] for r in decisions)),'by_month':{str(m):dict(Counter(r['action'] for r in decisions if int(r['date'][5:7])==m)) for m in range(1,13)},'new_invoices':len(new_audit),'new_lines':len(newrows),'zero':dict(Counter(z['kind']+('_pending' if z['reasons'] else '_ready') for z in zero)),'review_reasons':dict(Counter(reason for r in decisions if r['action']=='review' for reason in set(r['reasons']))),'sheet_only':result['sheet_only'],'returns_lines':len(returns)}
    save(OUT/'summary.json',summary);print(json.dumps(summary,ensure_ascii=False),flush=True)

if __name__=='__main__':main()
