"""Prepare exactly 89 verified F0 rows and the official adjustment of invoice 31."""
import json,hashlib,sys
from pathlib import Path
from datetime import date
from prepare_invoice_reports import rows,invoice_key
P=Path(__file__).resolve().parents[1];O=P/'outputs/doi-chieu-amis-j-20261031'
def main():
    x=json.loads((O/'year-review-chains.json').read_text(encoding='utf-8'))
    assert len(x['invoices'])==89
    byfile={}
    for item in x['invoices']:byfile.setdefault(item['source']['file'],{})[item['source']['row']]=item
    for file,wanted in byfile.items():
        for rn,r in rows(P/file):
            if rn==5:h={v:k for k,v in r.items() if v}
            if rn in wanted:
                item=wanted[rn];s=item['source']
                assert list(invoice_key(r[h['Ký hiệu']],r[h['Số hóa đơn']]))==s['key']
                item['tax_id']=str(r.get(h['MST/CCCD chủ hộ']) or '')
                assert str(r[h['Mã đơn hàng']]).strip()==s['order_code']
    data=[]
    for i,item in enumerate(sorted(x['invoices'],key=lambda a:int(a['source']['key'][1]))):
        rn=37088+i;s=item['source'];r=[None]*34
        r[0]=37088+i;r[1]=s['buyer'];r[2]=item['tax_id'] or None;r[3]=s['order_code'];r[7]=len(item['links'])
        r[8]={'formulaValue':f'=IF(O{rn}="";"F0; ";"")&IF(AND(Q{rn}<>"";V{rn}="");"F1; ";"")&IF(AND(X{rn}<>"";AC{rn}="");"F2; ";"")'}
        r[9:15]=[s['key'][0],s['key'][1],(date.fromisoformat(s['date'])-date(1899,12,30)).days,int(s['net']),int(s['vat']),int(s['payment'])]
        if item['links']:
            assert len(item['links'])==1 and s['key']==['1C26MUV','00000031']
            e=item['links'][0];r[15:22]=['Điều chỉnh',e['new'][0],e['new'][1],(date.fromisoformat(e['new_date'])-date(1899,12,30)).days,int(e['new_net']),int(e['new_vat']),int(e['new_payment'])]
        for col,parts in [(29,['M','T'] if item['links'] else ['M']),(30,['N','U'] if item['links'] else ['N']),(31,['O','V'] if item['links'] else ['O'])]:
            args=';'.join(f'{c}{rn}' for c in parts);r[col]={'formulaValue':f'=IF(COUNT({args})={len(parts)};SUM({args});"")'}
        r[32]='Có điều chỉnh' if item['links'] else 'Chưa có HĐ liên quan'
        r[33]=f'Bổ sung HĐ phát hành 2026 cho chứng từ AMIS 2025: {item["vouchers"][0]} (AMIS dòng {item["amis_rows"][0]}). Khớp mã đơn, số HĐ, VAT và thanh toán với nguồn đến 29/10/2026.'
        if item['links']:r[33]+=' Có HĐ điều chỉnh 1C26MUV/00058521 ngày 03/07/2026, số dư chuỗi 0 đủ 3 khoản.'
        data.append({'row':rn,'amis_row':item['amis_rows'][0],'voucher':item['vouchers'][0],'source':s,'links':item['links'],'values':r,'expected_current_money':[int(v) for v in item['current_chain_money']]})
    assert len({tuple(r['source']['key']) for r in data})==89
    (O/'ktoan-89-prepared.json').write_text(json.dumps({'insert_before_row':37088,'new_rows':[37088,37176],'old_last_row_after':37177,'rows':data},ensure_ascii=False,indent=2),encoding='utf-8')
    print('Prepared 89 rows / 90 invoice keys; source tax IDs preserved')
if __name__=='__main__':main()
