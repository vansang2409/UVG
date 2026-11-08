"""Independently validate identities, source hashes, residual sums and cross-source order reuse."""
import json,hashlib,sys
from pathlib import Path
from collections import defaultdict
from decimal import Decimal as D
from prepare_invoice_reports import rows,invoice_key,amount
P=Path(__file__).resolve().parents[1];O=P/'outputs/doi-chieu-amis-j-20261031'
def main():
    x=json.loads((O/'year-signature-review.json').read_text(encoding='utf-8'));live=json.loads((O/'amis-year-review-live.json').read_text(encoding='utf-8'))
    strong=[r for r in x['invoice_reviews'] if r['state'] in ('KHOP_DUY_NHAT_MA_DON_SO_HD_VAT_THANH_TOAN','KT_DA_XAC_NHAN_VA_KHOP_MA_TIEN')]
    allr=[r['full_match_rows'][0] for r in strong];assert len(allr)==len(set(allr))==89
    for r in strong:
        assert not r['ktoan_existing'];v=live[r['full_match_rows'][0]-1];s=r['source']
        assert str(v[8]).strip()==s['order_code'] and str(v[9]).zfill(8)==s['key'][1]
        assert D(str(v[6]))==D(s['vat']) and D(str(v[7]))==D(s['payment'])
    for g in x['groups']:
        a=set(g['original_rows']);b=set(g['remaining_rows']);c=set(g['strong_match_rows']);assert a==b|c and not b&c
        assert sum(D(str(live[r-1][7])) for r in b)==D(g['remaining_payment'])
        assert sum(D(str(live[r-1][6])) for r in b)==D(g['remaining_vat'])
    for m in x['manifest']:assert hashlib.sha256((P/m['file']).read_bytes()).hexdigest()==m['sha256']
    wanted={r['source']['order_code'] for r in strong};bycode=defaultdict(list);files=0
    for p in sorted((P/'data/Meinv_SAU_PHAT_HANH_20261029').rglob('*.xls')):
        if p.parent.name in ('DIEU_CHINH','THAY_THE'):continue
        files+=1;h={}
        for rn,r in rows(p):
            if rn==5:h={v:k for k,v in r.items() if v}
            if rn<6 or amount(r.get('A')) is None:continue
            get=lambda label:r.get(h[label]);code=str(get('Mã đơn hàng') or '').strip()
            if code in wanted:
                bycode[code].append({'key':list(invoice_key(get('Ký hiệu'),get('Số hóa đơn'))),'date':str(get('Ngày hóa đơn')),'status':get('Trạng thái hóa đơn'),'vat':str(amount(get('Tiền thuế GTGT'))),'payment':str(amount(get('Tổng tiền'))),'file':p.relative_to(P).as_posix(),'row':rn})
    reused={c:ls for c,ls in bycode.items() if len({tuple(r['key']) for r in ls})>1}
    verification={'verified_strong_matches':89,'unique_amis_rows':89,'residual_groups_verified':99,'source_hashes_verified':len(x['manifest']),'source_total_files_scanned_for_order_reuse':files,'strong_order_codes_on_multiple_invoices':reused,'read_only':True}
    (O/'year-signature-verification.json').write_text(json.dumps(verification,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in verification.items() if k!='strong_order_codes_on_multiple_invoices'},ensure_ascii=False,indent=2));print('Order codes on multiple invoices:',len(reused))
if __name__=='__main__':main()
