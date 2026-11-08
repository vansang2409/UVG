"""Reverse nonzero previous adjustments, then reverse F0 item details (2024 workflow)."""
from pathlib import Path
from decimal import Decimal,ROUND_HALF_UP
from datetime import date
import json,hashlib,sys
from prepare_invoice_reports import invoice_day
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def n(v):return Decimal(str(v)) if v is not None and str(v).strip() else None
def cell(v):return None if v is None or v=='' else v
def native(x):
    if isinstance(x,Decimal):return int(x) if x==int(x) else float(x)
    if isinstance(x,list):return list(map(native,x))
    if isinstance(x,dict):return {k:native(v) for k,v in x.items()}
    return x
def save(p,x):p.write_text(json.dumps(native(x),ensure_ascii=False,indent=2),encoding='utf-8')
def main():
    a=read(OUT/'authoring.json');d=read(OUT/'adjustment-rule-2024-audit.json');o=next(o for o in a['outputs'] if o['kind']=='adjustment')
    result=[];meta=[];chainmap=[];skipped=[];price_source_flags=[];seq=0;current=(date(2026,11,9)-date(1899,12,30)).days
    def money(k):
        h=d['headers'][k]['values'];m=[n(h.get(c)) for c in ['P','Q','R']]
        if m[2] is None:
            ls=d['detail'].get(k,[]);pay=[n(l['fields']['pay']) for l in ls]
            assert pay and all(v is not None for v in pay),(k,'missing total payment')
            assert sum(n(l['fields']['goods']) for l in ls)==m[0] and sum(n(l['fields']['tax']) for l in ls)==m[1]
            m[2]=sum(pay)
        assert all(v is not None for v in m) and m[0]+m[1]==m[2],k
        return m
    for c in sorted(d['chains'],key=lambda c:c['old_sequence']):
        root=c['root'];hd=d['headers'][root]['values'];s,no=root.split('/');day=(invoice_day(hd['D'])-date(1899,12,30)).days;ls=d['detail'][root]
        assert set(c['keys'])==set(c['active']),'Replacement/cancellation chain requires separate review'
        sums=[sum(money(k)[i] for k in c['active']) for i in range(3)];assert sums==list(map(n,c['before'])),root
        common=[None,current,cell(hd.get('H')),cell(hd.get('I')),cell(hd.get('J')),cell(hd.get('L')),None,cell(hd.get('M')) or 'TM/CK',None,s,no,day,cell(hd.get('E'))]
        docseq=[]
        for k in sorted((k for k in c['active'] if k!=root),key=lambda k:(d['ledger'][k]['date'],k)):
            m=money(k)
            if all(v==0 for v in m):skipped.append({'root':root,'invoice':k,'money':m});continue
            rates={str(l['fields']['rate']).strip().rstrip('%') for l in d['detail'].get(k,[]) if l['fields']['rate'] not in [None,'']}
            possible=[v for v in [0,5,8,10] if abs((m[0]*v/100).quantize(Decimal(1),rounding=ROUND_HALF_UP)-m[1])<=2]
            assert len(rates)==1 or len(possible)==1,(k,'rate')
            rate=Decimal(next(iter(rates))) if len(rates)==1 else Decimal(possible[0]);assert rate in [0,5,8,10]
            seq+=1;reason=f'Điều chỉnh {"tăng" if m[0]<=0 else "giảm"} để đảo khoản đã điều chỉnh trên HĐ {k}; tham chiếu HĐ gốc {root}'
            row=[*common,reason,reason,None,None,None,-m[0],-m[0],rate,-m[1],-m[2],None,'; '.join(c['docs']) or None,root,*m];row[0]=seq
            result.append(row);docseq.append(seq);meta.append({'seq':seq,'root':root,'source_invoice':k,'role':'Đảo khoản điều chỉnh cũ','first':True,'source':d['headers'][k]['source']})
        seq+=1;docseq.append(seq);m=money(root)
        assert [sum(n(l['fields'][f]) or 0 for l in ls) for f in ['goods','tax','pay']]==m,(root,'detail totals')
        assert all((n(l['fields'].get('discount')) or 0)==0 for l in ls),root
        rates={str(l['fields']['rate']).strip().rstrip('%') for l in ls if l['fields']['rate'] not in [None,'']};assert len(rates)==1,(root,'mixed rates');rate=Decimal(next(iter(rates)))
        reason=f'Điều chỉnh giảm toàn bộ hàng bán trên HĐ số {no} ngày {invoice_day(hd["D"]).strftime("%d/%m/%Y")} để đưa chuỗi về 0'
        for i,l in enumerate(ls):
            f=l['fields'];qty=n(f['qty']);price=n(f['price']);goods=n(f['goods']);assert qty is not None and qty>0 and price is not None and goods is not None and f['name'],(root,f)
            if abs(qty*price-goods)>2:price_source_flags.append({'root':root,'source':l['source'],'source_qty':qty,'source_price':price,'source_goods':goods,'preserved':True})
            first=i==0;row=[*common,reason,f['name'],cell(f['unit']),-qty,price,-goods,-m[0] if first else None,rate if first else None,-m[1] if first else None,-m[2] if first else None,1 if goods==0 else None,'; '.join(c['docs']) or None,root,*[v if first else None for v in m]];row[0]=seq
            result.append(row);meta.append({'seq':seq,'root':root,'source_invoice':root,'role':'Giảm toàn bộ chi tiết F0','first':first,'source':l['source']})
        chainmap.append({'root':root,'old_sequence':c['old_sequence'],'sequences':docseq,'original_reversal_sequence':seq,'before':c['before'],'origin':c['origin']})
    assert len(chainmap)==18 and seq==25 and len(result)==33 and len(skipped)==13
    o['rows']=result;o['invoices']=seq;o['headers'][26:29]=['Tiền hàng HĐ nguồn','VAT HĐ nguồn','Thanh toán HĐ nguồn']
    o['notes']=['Theo quy trình2024: đảo riêng từng khoản điều chỉnh cũ còn tiền, sau đó giảm toàn bộ chi tiết F0.','18 chuỗi /25 HĐ /33 dòng: 7 HĐ đảo khoản cũ và18 HĐ giảm chi tiết F0; 13 khoản cũ0đ không tạo thêm.','Tổng tiền hàng/VAT/thanh toán và thuế suất chỉ ghi dòng đầu mỗi HĐ. Dòng đảo tiền để trống SL/đơn giá.','Tên khách hàng và Người mua hàng lấy riêng nguyên trạng HĐ gốc; mọi HĐ mới tham chiếu F0.','Ngày dự thảo09/11/2026; dữ liệu kiểm đến01/11/2026. CộtY:AC chỉ tra cứu; không ghép vào trường import.']
    plan={'chains':chainmap,'rows':meta,'skipped_zero_adjustments':skipped,'source_price_flags':price_source_flags,'invoices':seq,'line_count':len(result),'identity_rule':'C=H và F=L của bảng tổng F0, giữ riêng hai trường, không fallback','workflow':'Đảo khoản điều chỉnh cũ khác0 → giảm chi tiết F0','audit_file':'adjustment-rule-2024-audit.json'}
    a['adjustment_rule_2024']=plan;a['summary']['adjustment']=seq;a['summary']['adjustment_chains']=18
    report=next(o for o in a['outputs'] if o['kind']=='report');report['sheets']=[s for s in report['sheets'] if s['name']!='Quy cach DC2024']
    byseq={r[0]:r for r in result};report['sheets'].append({'name':'Quy cach DC2024','headers':['F0','STT cũ','STT mới','Bước xử lý','HĐ nguồn đảo tiền','Tên khách hàng F0','Người mua hàng F0','Tiền hàng điều chỉnh','VAT điều chỉnh','Thanh toán điều chỉnh','Dòng nguồn','Ghi chú'],'rows':[[m['root'],next(c['old_sequence'] for c in chainmap if c['root']==m['root']),m['seq'],m['role'],m['source_invoice'],byseq[m['seq']][2],byseq[m['seq']][5],byseq[m['seq']][19],byseq[m['seq']][21],byseq[m['seq']][22],str(m['source']),'Tham chiếu F0; tên chép riêng từ F0'] for m in meta if m['first']]})
    for sh in report['sheets']:
        if sh['name']=='MI2025 chua map':
            for r in sh['rows']:
                if r[0] in {c['root'] for c in chainmap}:r[10]='; '.join(map(str,next(c['sequences'] for c in chainmap if c['root']==r[0])))
    for path in [OUT/'adjustment-rule-2024-audit.json',ROOT/'scripts/data/check-review/add-adjusted-import.mjs',ROOT/'scripts/data/check-review/verify-adjusted-import.py']:
        k=str(path.relative_to(ROOT));a['manifest']=[m for m in a['manifest'] if m['file']!=k];a['manifest'].append({'file':k,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
    for m in d['manifest']:
        a['manifest']=[x for x in a['manifest'] if x['file']!=m['file']];a['manifest'].append(m)
    save(OUT/'adjustment-rule-2024-plan.json',plan);save(OUT/'authoring.json',a);save(OUT/'final-summary.json',a['summary'])
    print(json.dumps({'chains':18,'invoices':seq,'rows':len(result),'zero_old_adjustments_skipped':len(skipped),'source_price_fields_preserved':len(price_source_flags)},ensure_ascii=False))
if __name__=='__main__':main()
