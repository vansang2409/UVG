"""Exact-name/unit/quantity diagnostics using historical vouchers validated against live totals."""
import json,re,sys
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal as D
from prepare_invoice_reports import rows
P=Path(__file__).resolve().parents[1];O=P/'outputs/doi-chieu-amis-j-20261031'
def norm(s):return re.sub(r'\s+',' ',str(s or '')).strip().casefold()
def main():
    cause=json.loads((O/'cause-investigation.json').read_text(encoding='utf-8'))
    hist=json.loads((O/'historical-amis-detail-check.json').read_text(encoding='utf-8'))
    audit=json.loads((O/'multi-voucher-audit.json').read_text(encoding='utf-8'))
    ai={g['invoice']:g for g in audit['groups']};hi={g['invoice']:g for g in hist['groups']}
    needed=defaultdict(set)
    for g in cause['groups']:
        for l in g['source_lines']:needed[l['file']].add(l['row'])
    units={}
    for file,rns in needed.items():
        for rn,r in rows(P/file):
            if rn in rns:units[(file,rn)]={'unit':r.get('M') or '', 'unit_price':r.get('O')}
    allpairs=[]
    for g in cause['groups']:
        n=g['invoice'];h=hi[n]
        g['historical_matches_live']=h['historical_matches_live'];g['historical_vouchers']=h['vouchers']
        g['verified']+=f" Sổ AMIS lịch sử khớp VAT/thanh toán hiện tại {h['historical_matches_live']}/{h['vouchers']} chứng từ; không dùng chi tiết các chứng từ không khớp hoặc chưa có."
        g['exact_name_price_differences']=[]
        if h['historical_matches_live']!=h['vouchers']:continue
        am=defaultdict(list);mi=defaultdict(list)
        for v in ai[n]['vouchers']:
            for l in hist['voucher_checks'][v]['rows']:am[(norm(l['name']),norm(l['unit']))].append({**l,'voucher':v})
        for l in g['source_lines']:
            u=units[(l['file'],l['row'])];mi[(norm(l['name']),norm(u['unit']))].append({**l,**u})
        for key,als in am.items():
            mls=mi.get(key,[])
            if not mls:continue
            aq=sum(D(l['qty']) for l in als);mq=sum(D(l['qty'] or 0) for l in mls)
            if aq!=mq or aq==0:continue
            an=sum(D(l['net']) for l in als);mn=sum(D(l['net']) for l in mls)
            if an==mn:continue
            if any(l['payment']=='None' or l['vat']=='None' for l in mls):continue
            ap=sum(D(l['payment']) for l in als);mp=sum(D(l['payment']) for l in mls)
            pair={'invoice':n,'name':als[0]['name'],'unit':als[0]['unit'],'qty':str(aq),'amis_net':str(an),'meinvoice_net':str(mn),'net_delta':str(an-mn),'amis_payment':str(ap),'meinvoice_payment':str(mp),'payment_delta':str(ap-mp),'amis_vouchers':list(dict.fromkeys(l['voucher'] for l in als)),'amis_source_rows':[l['row'] for l in als],'meinvoice_source_rows':[l['row'] for l in mls],'meinvoice_file':mls[0]['file'],'scope':'Same literal normalized name + unit + summed quantity; comparison invoice signature still unconfirmed; not proof of order identity'}
            g['exact_name_price_differences'].append(pair);allpairs.append(pair)
        if g['exact_name_price_differences']:
            ps=g['exact_name_price_differences'];contrib=sum(D(p['payment_delta']) for p in ps)
            g['verified']+=f" Có {len(ps)} mục cùng tên chuẩn hóa/ĐVT/tổng SL nhưng tiền trước thuế khác; cộng phần chênh thanh toán của các mục này {contrib}đ. Đây là đối chiếu mục hàng với HĐ ứng viên, chưa xác nhận danh sách đơn gộp."
            if g['category']=='Chưa đủ căn cứ giải thích chênh thanh toán':g['category']='Có chênh tiền hàng trên mục cùng tên/ĐVT/SL; chưa chốt toàn bộ'
        # Correct earlier missing-source wording after locating validated historical evidence.
        g['unconfirmed']=g['unconfirmed'].replace('Chưa có chi tiết bán hàng AMIS 2025 theo mặt hàng/SL hoặc bảng kê gộp có mã chứng từ để xác định khoản thiếu/thừa.','Đã có chi tiết AMIS T1–6/2025 lịch sử cho chứng từ khớp tổng hiện tại, nhưng chưa có bảng kê gộp nối trực tiếp từng chứng từ với HĐ để chốt khoản thiếu/thừa.')
    cause['summary']['categories']=dict(Counter(g['category'] for g in cause['groups']))
    cause['summary']['historical_detail_groups_fully_validated']=hist['summary']['groups_all_vouchers_match_live']
    cause['summary']['historical_detail_vouchers_validated']=hist['summary']['matched_live_vouchers']
    cause['summary']['exact_name_unit_qty_price_difference_groups']=sum(bool(g['exact_name_price_differences']) for g in cause['groups'])
    cause['summary']['exact_name_unit_qty_price_difference_items']=len(allpairs)
    cause['summary']['historical_source']=hist['summary']
    cause['summary']['source_only_has_amis_summary_not_2025_goods_details']=False
    (O/'cause-investigation-final.json').write_text(json.dumps(cause,ensure_ascii=False,indent=2),encoding='utf-8')
    import csv
    with (O/'CHENH_TIEN_HANG_CUNG_TEN_DVT_SL.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Số HĐ J','Tên hàng','ĐVT','SL hai bên','Tiền hàng AMIS','Tiền hàng HĐ so sánh','Chênh tiền hàng','Thanh toán AMIS','Thanh toán HĐ','Chênh thanh toán','Chứng từ AMIS','Dòng nguồn AMIS lịch sử','Dòng nguồn meInvoice','Giới hạn'])
        for p in allpairs:w.writerow([p['invoice'],p['name'],p['unit'],p['qty'],p['amis_net'],p['meinvoice_net'],p['net_delta'],p['amis_payment'],p['meinvoice_payment'],p['payment_delta'],';'.join(p['amis_vouchers']),';'.join(map(str,p['amis_source_rows'])),';'.join(map(str,p['meinvoice_source_rows'])),'HĐ ứng viên chưa chốt ký hiệu; chưa chứng minh cùng tập đơn hàng'])
    print(json.dumps(cause['summary'],ensure_ascii=False,indent=2))

if __name__=='__main__':main()
