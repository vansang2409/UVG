"""Build template-shaped import payloads and a complete decision ledger."""
from pathlib import Path
from collections import defaultdict,Counter
from datetime import date
from decimal import Decimal,ROUND_HALF_UP
from io import BytesIO
import json,csv,sys,hashlib
from openpyxl import load_workbook
from prepare_invoice_reports import rows,invoice_day

ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def txt(x):return '' if x is None else str(x).strip()
def n(x):return Decimal(str(x or 0))
def save(name,v):(OUT/name).write_text(json.dumps(v,ensure_ascii=False,default=lambda x:float(x) if isinstance(x,Decimal) else str(x)),encoding='utf-8')
def headers(path,count):
    w=load_workbook(BytesIO(path.read_bytes()),read_only=True,data_only=True)
    h=list(next(w.active.iter_rows(min_row=9,max_row=9,max_col=count,values_only=True)));w.close();return h
def main():
    p=json.loads((OUT/'prepared.json').read_text(encoding='utf-8'));decisions={r['doc']:r for r in p['decisions']}
    cache=json.loads((OUT/'source-index.json').read_text(encoding='utf-8'));ledger=cache['ledger']
    # Read original identity fields from the exact invoice source row, never from another buyer.
    wanted=defaultdict(set)
    for z in p['zero']:
        if not z['reasons']:
            src=ledger[z['root']]['source'];wanted[src['file']].add(src['row'])
    identities={}
    for f,rownums in wanted.items():
        for rn,r in rows(ROOT/f):
            if rn in rownums:
                k=txt(r.get('B'))+'/'+txt(r.get('C')).zfill(8)
                identities[k]={'buyer':r.get('L'),'company':r.get('H'),'address':r.get('I'),'tax':r.get('J'),'date':str(invoice_day(r.get('D'))),'cqt':r.get('E'),'method':r.get('M')}
    template_new=ROOT/'data/Du lieu Meinvoice/Mau_hoa_don_GTGT_MTT_1_thue_suat.xls'
    template_adj=ROOT/'data/Du lieu Meinvoice/Mau HD Dieu chinh.xls'
    template_rep=ROOT/'data/Mau_hoa_don_GTGT_MTT_1_thue_suat_Thay_the.xls'
    ah=headers(template_adj,24)+['Chứng từ AMIS','Chuỗi HĐ','Số dư tiền hàng cũ','Số dư VAT cũ','Số dư thanh toán cũ']
    rh=headers(template_rep,22)+['Lý do thay thế','Là dòng ghi chú/diễn giải','Chứng từ AMIS','Chuỗi HĐ']
    ar=[];rr=[];zeroaudit=[]
    current=(date(2026,11,9)-date(1899,12,30)).days
    for z in p['zero']:
        if z['reasons']:continue
        root=z['root'];ident=identities.get(root);bal=z['balance'];rate_candidates=[r for r in [0,5,8,10] if abs((n(bal[0])*r/100).quantize(Decimal(1),rounding=ROUND_HALF_UP)-n(bal[1]))<=2]
        if not ident or not ident['date'] or not ident['buyer'] or len(rate_candidates)!=1 or n(bal[0])+n(bal[1])!=n(bal[2]):
            z['reasons'].append('Thiếu thông tin HĐ gốc hoặc chưa chốt thuế suất');continue
        s,no=root.split('/');day=(date.fromisoformat(ident['date'])-date(1899,12,30)).days;rate=rate_candidates[0]
        buyer=ident['buyer'];company=ident['company'] or buyer;docs='; '.join(z['docs'])
        if z['kind']=='adjustment':
            seq=len(ar)+1;reason=f'Điều chỉnh số dư hóa đơn {root} về 0 do nội dung chưa khớp AMIS; lập lại theo chứng từ {docs}'
            ar.append([seq,current,ident['company'],ident['address'],ident['tax'],buyer,None,'TM/CK',None,s,no,day,ident['cqt'],reason,reason,None,None,None,-n(bal[0]),-n(bal[0]),rate,-n(bal[1]),-n(bal[2]),None,docs,root,*bal])
        else:
            seq=len(rr)+1;reason=f'Thay thế hóa đơn số {no} ngày {date.fromisoformat(ident["date"]).strftime("%d/%m/%Y")} về 0 do chưa khớp AMIS; lập lại theo chứng từ {docs}'
            rr.append([seq,current,company,ident['tax'],ident['address'],buyer,None,None,None,'TM/CK',None,s,no,day,ident['cqt'],reason,None,0,0,0,rate,0,reason,1,docs,root])
        zeroaudit.append({'root':root,'kind':z['kind'],'sequence':seq,'docs':z['docs'],'before':bal,'change':[-n(b) for b in bal],'after':[0,0,0],'identity_source':ledger[root]['source']})
    # Propagate any zeroing dependency failure to new invoices, then rebuild all sequence numbers.
    blocked={c for z in p['zero'] if z['reasons'] for c in z['docs']}
    for c in blocked:
        if decisions[c]['action']=='new':decisions[c]['action']='review';decisions[c]['reasons'].append('Hóa đơn cũ chưa đủ điều kiện xử lý về 0')
    # If one invoice in a shared chain fails, the whole chain waits.
    blocked_roots={z['root'] for z in p['zero'] if z['reasons']}
    zeroaudit=[x for x in zeroaudit if x['root'] not in blocked_roots]
    outputs=[]
    new_by_month=defaultdict(list)
    for r in p['new_rows']:
        if decisions[r[18]]['action']!='new':continue
        month=int(r[20][5:7]);key=f'T{month:02d}' if month<=6 else 'T07_T12';new_by_month[key].append(r)
    for key,rs in sorted(new_by_month.items()):
        seqs={};nr=[]
        for r in rs:
            code=r[18]
            if code not in seqs:seqs[code]=len(seqs)+1
            nr.append([seqs[code],current,*r[2:]])
        outputs.append({'file':'IMPORT_HD_MOI_2025_'+key+'.xlsx','kind':'new','headers':p['new_headers'],'rows':nr,'invoices':len(seqs),'notes':['Hóa đơn mới theo AMIS xuất 08/11/2026.','Ngày hóa đơn dự thảo 09/11/2026; điều chỉnh đồng nhất theo ngày phát hành thực tế trước khi nhập.','Giữ nguyên dòng bán gốc; các phiếu trả lại nằm trong báo cáo chờ đối chiếu liên kết, chưa bù trừ.','Các cột bổ sung R:V dùng kiểm tra/ghép trường; không ghép nhầm vào trường tiền.']})
    outputs += [{'file':'IMPORT_DIEU_CHINH_VE_0_2025.xlsx','kind':'adjustment','headers':ah,'rows':ar,'invoices':len(ar),'notes':['Điều chỉnh tiền theo số dư còn lại của chuỗi đã có điều chỉnh.','Một dòng điều chỉnh số tiền chung, SL/đơn giá để trống theo mẫu hiện có; không bịa dòng hàng.','Số dư âm được điều chỉnh tăng để về 0.','Ngày dự thảo 09/11/2026. Cần kiểm lịch sử phát sinh sau bộ nguồn 01/11 trước khi nhập.']},{'file':'IMPORT_THAY_THE_VE_0_2025.xlsx','kind':'replacement','headers':rh,'rows':rr,'invoices':len(rr),'notes':['Chỉ gồm chuỗi có F0 và chưa có điều chỉnh/thay thế.','Một dòng ghi chú/diễn giải; ghép đúng cột Là dòng ghi chú/diễn giải = 1.','SL/đơn giá/tiền bằng 0 trong dự thảo; meInvoice có thể để trống tiền và thuế suất trên dòng ghi chú sau phát hành.','Ngày dự thảo 09/11/2026. Kiểm tra lịch sử mới trước khi nhập.']}]
    if 'Hàng khuyến mại' in p['new_headers']:
        for o in outputs:
            if o['kind']=='new':o['notes'][3]+=' Cột W - Hàng khuyến mại: giá trị 1 trên đúng dòng KM, cần ghép vào trường Hàng khuyến mại của meInvoice.'
    labels={'new':'Tạo import mới','keep':'Giữ nguyên, không import','company':'Loại theo HD CTY 2025','review':'Chờ xác minh'}
    summary=[]
    for m in range(1,13):
        group=[r for r in decisions.values() if int(r['date'][5:7])==m]
        summary.append([m,len(group),*[sum(r['action']==a for r in group) for a in ['new','keep','company','review']],sum(n(r['pay']) for r in group if r['action']=='new')])
    decision_headers=['Chứng từ AMIS','Ngày hạch toán','Phân loại','Tiền hàng AMIS','VAT AMIS','Thanh toán AMIS','Dòng AMIS bản sao','Lệch trên Sheet','Chuỗi liên quan','Kiểm chi tiết','Ghi chú']
    content_labels={'not_checked':'Chưa so chi tiết','matched':'Khớp nguyên văn','different':'Tên/ĐVT/SL khác nguyên văn','missing_detail':'Thiếu chi tiết meInvoice'}
    def drow(r):return [r['doc'],r['date'],labels[r['action']],r['goods'],r['vat'],r['pay'],r['live_row'],str(r['sheet_delta']) if isinstance(r['sheet_delta'],str) else r['sheet_delta'],'; '.join(r['roots']),content_labels[r['content']],'; '.join(r['reasons'])]
    report={'file':'BAO_CAO_IMPORT_2025.xlsx','kind':'report','sheets':[{'name':'Tong hop','headers':['Tháng 2025','Tổng chứng từ','Import mới','Giữ nguyên','Loại công ty','Chờ xác minh','Thanh toán import mới (VND)'],'rows':summary},{'name':'Cho xac minh','headers':decision_headers,'rows':[drow(r) for r in decisions.values() if r['action']=='review']},{'name':'Loai cong ty','headers':decision_headers,'rows':[drow(r) for r in decisions.values() if r['action']=='company']},{'name':'Xu ly chuoi cu','headers':['HĐ gốc','Loại xử lý','Chứng từ AMIS','Số dư tiền hàng','Số dư VAT','Số dư thanh toán','Các HĐ trong chuỗi','Chờ vì'],'rows':[[z['root'],z['kind'],'; '.join(z['docs']),*z['balance'],'; '.join(z['keys']),'; '.join(z['reasons'])] for z in p['zero']]},{'name':'BTL cho lien ket','headers':['Phiếu trả','Dòng nguồn','Ngày hạch toán','Khách hàng','Mã hàng','Tên hàng','ĐVT','SL trả','Tiền hàng trả','VAT','Thanh toán','Chứng từ bán gốc'],'rows':[[r['doc'],r['source'],r['v'][0],r['v'][9],r['v'][15],r['v'][7] or r['v'][16],r['v'][19],r['v'][32],r['v'][36],r['v'][41],r['v'][43],None] for r in p['returns']]},{'name':'Chi co tren Sheet','headers':['Chứng từ','Xử lý'],'rows':[[c,'Không đưa vào import vì không còn trong AMIS mới'] for c in p['sheet_only']]}]}
    review_path=OUT/'h1-12-invoice-link-review.json'
    if review_path.exists():
        review=json.loads(review_path.read_text(encoding='utf-8'))
        report['sheets'].append({'name':'Ra soat 12 don T6','headers':['Chứng từ AMIS','Mã đơn hàng','Thanh toán AMIS','HĐ trên Sheet chưa xác minh','Ngày HĐ trên Sheet','Trạng thái AMIS','Ký hiệu HĐ trên AMIS','Số hóa đơn trên AMIS','HĐ thực tế','Ghi chú kết quả','Ghi chú nguồn'],'rows':[[r['doc'],r['order'],r['pay'],r['snapshot_invoice'],r['snapshot_invoice_date'],r['amis_status'],r['amis_symbol'],r['amis_invoice_number'],'Chưa xác định',r['conclusion'],r['amis_url']] for r in review['records']]})
    outputs.append(report)
    with (OUT/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(decision_headers);w.writerows(drow(r) for r in decisions.values())
    with (OUT/'LIEN_KET_IMPORT_MOI.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['File','STT HĐ','Chứng từ AMIS','Dòng đầu Excel','Số dòng','HĐ cũ'])
        for o in outputs:
            if o['kind']!='new':continue
            grouped=defaultdict(list)
            for i,r in enumerate(o['rows'],10):grouped[r[18]].append((i,r))
            for code,ls in grouped.items():w.writerow([o['file'],ls[0][1][0],code,ls[0][0],len(ls),'; '.join(decisions[code]['roots'])])
    summary_final={'total':len(decisions),'actions':dict(Counter(r['action'] for r in decisions.values())),'new_invoices':sum(o['invoices'] for o in outputs if o['kind']=='new'),'new_lines':sum(len(o['rows']) for o in outputs if o['kind']=='new'),'adjustment':len(ar),'replacement':len(rr),'zero_pending':sum(bool(z['reasons']) for z in p['zero']),'new_goods':sum(n(r['goods']) for r in decisions.values() if r['action']=='new'),'new_vat':sum(n(r['vat']) for r in decisions.values() if r['action']=='new'),'new_pay':sum(n(r['pay']) for r in decisions.values() if r['action']=='new'),'source_returns_docs':len({r['doc'] for r in p['returns']}),'source_returns_lines':len(p['returns']),'source_returns_pay':sum(n(r['v'][43]) for r in p['returns']),'sheet_only':p['sheet_only'],'limitations':['412 dòng BTL và BGG trong nguồn mới chưa có trường số chứng từ bán gốc; không tự nối theo tên/tiền.','Khớp đủ tiền được giữ theo yêu cầu. Cột kiểm chi tiết chỉ ghi kết quả so tên/ĐVT/SL nguyên văn, chưa quy đổi combo/model.','Chỉ có nguồn meInvoice đến 29/10 và bổ sung 01/11/2026, chưa xác minh phát sinh mới hơn.','Chuỗi thay thế phức tạp và chuỗi dùng chung với phần giữ/loại tách chờ.']}
    save('authoring.json',{'outputs':outputs,'summary':summary_final,'zero_audit':zeroaudit,'manifest':p['manifest']+[{'file':str(t.relative_to(ROOT)),'sha256':hashlib.sha256(t.read_bytes()).hexdigest()} for t in [template_new,template_adj,template_rep]]})
    save('final-summary.json',summary_final);print(json.dumps(summary_final,ensure_ascii=False,default=str))
if __name__=='__main__':main()
