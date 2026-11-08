"""Check official modification links before preparing any Ktoan additions."""
import json,sys,csv
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal as D
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount
P=Path(__file__).resolve().parents[1];O=P/'outputs/doi-chieu-amis-j-20261031'
def main():
    x=json.loads((O/'year-signature-review.json').read_text(encoding='utf-8'))
    strong=[r for r in x['invoice_reviews'] if r['state'] in ('KHOP_DUY_NHAT_MA_DON_SO_HD_VAT_THANH_TOAN','KT_DA_XAC_NHAN_VA_KHOP_MA_TIEN')]
    edges=defaultdict(dict)
    for folder,kind in [('DIEU_CHINH','adjustment'),('THAY_THE','replacement')]:
        for p in sorted((P/'data/Meinv_SAU_PHAT_HANH_20261029'/folder).glob('*.xls')):
            for rn,r in rows(p):
                if rn<7 or amount(r.get('A')) is None:continue
                a=invoice_key(r.get('B'),r.get('D'));b=invoice_key(r.get('M'),r.get('O'))
                edges[a][b]={'kind':kind,'old':list(a),'new':list(b),'new_date':str(invoice_day(r.get('N'))),'new_net':str(amount(r.get('U'))),'new_vat':str(amount(r.get('V'))),'new_payment':str(amount(r.get('W'))),'file':p.relative_to(P).as_posix(),'row':rn}
    results=[]
    for r in strong:
        s=r['source'];key=tuple(s['key']);seen={key};todo=[key];links=[]
        while todo:
            k=todo.pop()
            for nk,e in edges.get(k,{}).items():
                if nk in seen:raise ValueError(('cycle/multi-parent',nk))
                seen.add(nk);todo.append(nk);links.append(e)
        if links:
            assert all(e['kind']=='adjustment' for e in links),'replacement requires distinct balance logic'
        money=[D(s[k]) for k in ('net','vat','payment')]
        for e in links:
            assert all(e[k]!='None' for k in ('new_net','new_vat','new_payment'))
            money=[m+D(e[k]) for m,k in zip(money,('new_net','new_vat','new_payment'))]
        results.append({'source':s,'amis_rows':r['full_match_rows'],'vouchers':r['full_match_vouchers'],'identity_state':r['state'],'links':links,'current_chain_money':[str(m) for m in money],'handling':'Cần đủ F0 + HĐ điều chỉnh, số dư chuỗi 0' if links and money==[D(0)]*3 else 'Cần giữ đủ lịch sử liên quan' if links else 'Chưa có liên kết sửa trong nguồn; HĐ mới khớp mã và tiền'})
    out={'summary':{'strong_identity_matches':len(results),'new_invoices_no_modification_links':sum(not r['links'] for r in results),'invoices_with_modification_links':sum(bool(r['links']) for r in results),'modified_chains_zero':sum(bool(r['links']) and r['current_chain_money']==['0','0','0'] for r in results),'ktoan_written':False},'invoices':results}
    (O/'year-review-chains.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
    with (O/'89_HD_KIEM_TRA_XONG_DANH_TINH_VA_CHUOI.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Dòng AMIS','Chứng từ','Ký hiệu','Số HĐ','Ngày','Mã đơn','Tiền hàng gốc','VAT gốc','Thanh toán gốc','Thanh toán chuỗi hiện tại','HĐ liên quan','Cách xử lý khi bổ sung','Căn cứ danh tính'])
        for r in results:
            s=r['source'];w.writerow([';'.join(map(str,r['amis_rows'])),';'.join(r['vouchers']),s['key'][0],s['key'][1],s['date'],s['order_code'],s['net'],s['vat'],s['payment'],r['current_chain_money'][2],json.dumps(r['links'],ensure_ascii=False),r['handling'],r['identity_state']])
    print(json.dumps(out['summary'],ensure_ascii=False,indent=2))
    for r in results:
        if r['links']:print(json.dumps(r,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
