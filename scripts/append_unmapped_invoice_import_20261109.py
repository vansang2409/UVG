"""Append approved orphan-chain imports without changing any previous invoice sequence."""
from pathlib import Path
from collections import Counter
from datetime import date
import json,sys,hashlib
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def save(p,x):p.write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf-8')
def main():
    a=read(OUT/'authoring.json');audit=read(OUT/'unmapped-meinvoice-audit.json');ready=[r for r in audit['components'] if r['status']=='ready'];old=a.get('unmapped_zero_audit',[])
    for o in a['outputs']:
        if o['kind'] in ['adjustment','replacement']:
            keys={r['root'] for r in old if r['kind']==o['kind']};ki=25
            o['rows']=[r for r in o['rows'] if r[ki] not in keys];o['invoices']=len(o['rows'])
    assert not {r['root'] for r in ready}&{r['root'] for r in a['zero_audit']}
    seqpath=OUT/'unmapped-import-sequences.json';seqs=read(seqpath) if seqpath.exists() else {};added=[]
    current=(date(2026,11,9)-date(1899,12,30)).days
    for r in sorted(ready,key=lambda r:r['root']):
        root=r['root'];kind=r['kind'];o=next(o for o in a['outputs'] if o['kind']==kind);s,no=root.split('/');ident=r['identities'][root]
        day=(date.fromisoformat(audit['invoice_details'][root]['record']['date'])-date(1899,12,30)).days;b=r['balance'];seq=len(o['rows'])+1
        if root in seqs:assert seqs[root]=={'kind':kind,'sequence':seq}
        else:seqs[root]={'kind':kind,'sequence':seq}
        buyer=ident['buyer'];reason=f'Đưa số dư chuỗi hóa đơn {root} về 0 do chưa nối được chứng từ AMIS trong phạm vi rà soát năm 2025'
        if kind=='adjustment':row=[seq,current,ident['company'],ident['address'],ident['tax'],buyer,None,'TM/CK',None,s,no,day,ident['cqt'],reason,reason,None,None,None,-b[0],-b[0],r['rate'],-b[1],-b[2],None,None,root,*b]
        else:row=[seq,current,None,ident['tax'],ident['address'],buyer,None,None,None,'TM/CK',None,s,no,day,ident['cqt'],reason,None,0,0,0,r['rate'],0,reason,1,None,root]
        o['rows'].append(row);o['invoices']=len(o['rows']);added.append({'root':root,'kind':kind,'sequence':seq,'keys':r['keys'],'before':b,'change':[-v for v in b],'after':[0,0,0],'identity_source':ident['source']})
    report=next(o for o in a['outputs'] if o['kind']=='report');newnames={'MI2025 chua map','MI2025 chuoi','Bieu tang - KT check'}
    report['sheets']=[s for s in report['sheets'] if s['name'] not in newnames]
    labels={'ready':'Đã bổ sung import','already_zero':'Đã về0, không thêm','company':'Loại HĐ công ty','review':'Chờ xác minh','kt_review':'Chờ kế toán — biếu tặng'}
    scope=[r for r in audit['components'] if r['status'] in labels];summary=[];details=[];gifts=[]
    gift_headers=['HĐ gốc','Ngày HĐ','Tên đơn vị nhận','MST','Tiền hàng','VAT','Thanh toán','Toàn bộ HĐ liên quan','Nghiệp vụ nguồn','Xử lý hiện tại','Ý kiến kế toán','File nguồn','Dòng nguồn']
    for r in scope:
        b=r.get('balance',[None]*3);entry=audit['invoice_details'].get(r['root'],{});rec=entry.get('record') or {};ident=(r.get('identities') or {}).get(r['root']) or {}
        summary.append([r['root'],labels[r['status']],r.get('kind'),len(r['keys']),*b,'; '.join(r.get('active',[])),'; '.join(r.get('excluded',[])),'; '.join(r['issues']),seqs.get(r['root'],{}).get('sequence'),'; '.join(r['keys'])])
        if r['status']=='kt_review':
            gifts.append([r['root'],rec.get('date'),ident.get('company'),ident.get('tax'),*b,'; '.join(r['keys']),'Biếu tặng/không thu tiền','Chờ kế toán xác nhận; chưa đưa vào import',None,ident.get('source',{}).get('file'),ident.get('source',{}).get('row')])
        depths={k:0 for k in r['roots']};todo=list(depths)
        while todo:
            k=todo.pop(0)
            for child in audit['invoice_details'].get(k,{}).get('children',[]):
                if child not in depths:depths[child]=depths[k]+1;todo.append(child)
        for k in r['keys']:
            node=audit['invoice_details'][k];v=node.get('record') or {};m=v.get('money',[None]*3);resolved=node.get('detail_verified_money',v.get('resolved_money',m));proof=node.get('payment_evidence')
            details.append([r['root'],k,'F'+str(depths.get(k,'?')),v.get('date'),'; '.join(node['parents']),'; '.join(node['replacement_children']),v.get('status'),v.get('cqt_status'),*m,*resolved,'Tính số dư' if k in r.get('active',[]) else 'Loại khỏi số dư','Chi tiết '+str(proof) if proof else None,v.get('source',{}).get('file'),v.get('source',{}).get('row')])
    assert len(gifts)==16
    report['sheets'] += [{'name':'MI2025 chua map','headers':['F0','Kết quả','Loại import','Số HĐ trong chuỗi','Số dư tiền hàng','Số dư VAT','Số dư thanh toán','HĐ tính số dư','HĐ loại khỏi số dư','Chờ vì','STT import bổ sung','Toàn bộ chuỗi'],'rows':summary},{'name':'MI2025 chuoi','headers':['F0','HĐ','Cấp liên kết','Ngày HĐ','HĐ cha','HĐ thay thế con','Trạng thái HĐ','Trạng thái CQT','Tiền hàng nguồn','VAT nguồn','Thanh toán nguồn','Tiền hàng dùng tính','VAT dùng tính','Thanh toán dùng tính','Vai trò','Bằng chứng bổ sung','File nguồn','Dòng nguồn'],'rows':details},{'name':'Bieu tang - KT check','headers':gift_headers,'rows':gifts}]
    a['unmapped_zero_audit']=added;a['summary']['adjustment']=next(o['invoices'] for o in a['outputs'] if o['kind']=='adjustment');a['summary']['replacement']=next(o['invoices'] for o in a['outputs'] if o['kind']=='replacement')
    a['summary']['unmapped_meinvoice']={'reviewed_components':len(audit['components']),'classification':audit['summary'],'added':dict(Counter(r['kind'] for r in added)),'gift_kt_check':16,'remaining_source_exceptions':2,'zeroed_payment':sum(r['before'][2] for r in added)}
    manifest={m['file']:m for m in a['manifest']}
    for p in [OUT/'unmapped-meinvoice-audit.json',OUT/'unmapped-invoice-detail-review.json',OUT/'unmapped-gift-confirmation.json']:
        manifest[str(p.relative_to(ROOT))]={'file':str(p.relative_to(ROOT)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
    for m in audit['source_files']:manifest[m['file']]=m
    a['manifest']=list(manifest.values());save(OUT/'authoring.json',a);save(OUT/'final-summary.json',a['summary']);save(seqpath,seqs);save(OUT/'gift-kt-check.json',{'headers':gift_headers,'rows':gifts})
    print(json.dumps(a['summary']['unmapped_meinvoice'],ensure_ascii=False))
if __name__=='__main__':main()
