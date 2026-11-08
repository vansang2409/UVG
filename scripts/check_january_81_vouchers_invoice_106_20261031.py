"""Check the specified 81 vouchers without turning a number-only match into identity proof."""
import json,csv,sys,re,hashlib
from pathlib import Path
from collections import defaultdict
from decimal import Decimal as D
from prepare_invoice_reports import rows
P=Path(__file__).resolve().parents[1];O=P/'outputs/doi-chieu-amis-j-20261031'
def family(t):
    t=t.upper();ms=sorted(set(re.sub(r'[\s_-]','',m) for m in re.findall(r'\b(?:KA|KH|KM|KC)\s*[-_]?\s*\d{2,4}',t)))
    machine=bool(re.match(r'^\s*(MÁY|ĐÈN)\b',t));flt=not machine and bool(re.search(r'FILTER|MÀNG LỌC|LÕI LỌC|BỘ LỌC|FT\s*[-:]|FT-',t))
    return ('FILTER:' if flt else 'MACHINE:')+','.join(ms) if ms else 'UNKNOWN'
def main():
    yr=json.loads((O/'year-signature-review.json').read_text(encoding='utf-8'));g=next(g for g in yr['groups'] if g['invoice_number']=='00000106')
    hist=json.loads((O/'historical-amis-detail-check.json').read_text(encoding='utf-8'));checks=hist['voucher_checks']
    live=json.loads((O/'live-jan-106-block.json').read_text(encoding='utf-8'));am=[];vouchers=[]
    for rn in g['remaining_rows']:
        v=live[rn-62393];check=checks[v[2]]
        assert check['matches_live'] and D(str(v[6]))==D(check['sums']['vat']) and D(str(v[7]))==D(check['sums']['payment'])
        assert str(v[9])=='106' and not v[10]
        am.extend({**l,'voucher':v[2],'amis_row':rn} for l in check['rows']);vouchers.append({'row':rn,'voucher':v[2],'date_serial':v[0],'order_code':v[8],'vat':v[6],'payment':v[7]})
    assert len(vouchers)==81 and len({v['voucher'] for v in vouchers})==81
    src=P/'data/Du lieu Meinvoice/Bang_ke_chi_tiet_HD_da_su_dung_2025.xls';mi=[]
    for rn,r in rows(src):
        if r.get('B')=='00000106' and r.get('D')=='1C25TUV' and r.get('L'):
            mi.append({'row':rn,'date':r['C'],'name':r['L'],'qty':r['N'],'unit':r['M'],'net':str(D(r['P'] or '0')-D(r['R'] or '0')),'vat':r['T'],'payment':r['U']})
    sums=lambda ls:{k:str(sum(D(str(l[k])) for l in ls)) for k in ('net','vat','payment')}
    a=sums(am);b=sums(mi)
    assert b=={'net':'110371942','vat':'8829755','payment':'119201697'}
    assert a['payment']=='127097054' and a['vat']=='9414590'
    comparisons=[]
    def agg(ls):
        out=defaultdict(lambda:{'qty':D(0),'vat':D(0),'payment':D(0)})
        for l in ls:
            z=out[(family(l['name']),str(l['unit']).casefold())]
            for k in z:z[k]+=D(str(l[k]))
        return out
    aa,bb=agg(am),agg(mi)
    for key in sorted(set(aa)|set(bb)):
        comparisons.append({'family':key[0],'unit':key[1],'amis':{k:str(v) for k,v in aa[key].items()},'invoice':{k:str(v) for k,v in bb[key].items()}})
    result={'conclusion':'CHUA_DU_CAN_CU_XAC_NHAN_CA_81_CHUNG_TU_THUOC_HD','invoice':['1C25TUV','00000106'],'live_amis_vouchers_verified':81,'amis_detail_lines':len(am),'invoice_detail_lines':len(mi),'amis_totals':a,'invoice_original_totals':b,'payment_delta':str(D(a['payment'])-D(b['payment'])),'old_2025_adjustment_payment':'5664046','payment_delta_after_old_2025_adjustment':str(D(a['payment'])-D(b['payment'])-D('5664046')),'explicit_historical_voucher_invoice_refs':sum(bool(l['signature'] or l['invoice']) for l in am),'diagnostic_only_family_unit_comparisons':comparisons,'vouchers':vouchers,'amis_lines':am,'invoice_lines':mi,'source_file':str(src.relative_to(P)),'source_sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'limit':'Family comparison does not establish SKU equivalence or order identity. Need a bill attachment or approved voucher-to-invoice list. No Google Sheet changes.'}
    (O/'january-81-vouchers-invoice-106.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    with (O/'81_CHUNG_TU_THANG_1_HD_106.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Dòng AMIS','Số chứng từ','Mã đơn','VAT','Thanh toán','Trạng thái đối chiếu'])
        for v in vouchers:w.writerow([v['row'],v['voucher'],v['order_code'],v['vat'],v['payment'],'Chưa có căn cứ nối trực tiếp với 1C25TUV/00000106'])
    print(json.dumps({k:v for k,v in result.items() if k not in ('vouchers','amis_lines','invoice_lines','diagnostic_only_family_unit_comparisons')},ensure_ascii=False,indent=2))
if __name__=='__main__':main()
