"""Read-only review of the 99 ambiguous groups using source order IDs, not amounts alone."""
import json,csv,sys,hashlib
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal as D
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount
P=Path(__file__).resolve().parents[1];O=P/'outputs/doi-chieu-amis-j-20261031'
def code(v):return str(v or '').strip()
def num(v):
    s=code(v)
    return s.zfill(8) if s.isdigit() else s
def main():
    cause=json.loads((O/'cause-investigation-final.json').read_text(encoding='utf-8'))
    old=json.loads((O/'multi-voucher-audit.json').read_text(encoding='utf-8'))
    oi={g['invoice']:g for g in old['groups']}
    gs=[g for g in cause['groups'] if g['category']=='Cần phân biệt số HĐ trùng năm/ký hiệu']
    assert len(gs)==99
    live=json.loads((O/'amis-year-review-live.json').read_text(encoding='utf-8'))
    ix={i+1:r for i,r in enumerate(live)}
    bycode=defaultdict(list)
    for rn,v in enumerate(live[1:],2):
        if code(v[8]) and not code(v[8]).startswith('#'):bycode[code(v[8])].append(rn)
    targets=defaultdict(set)
    old_candidates=[]
    for g in gs:
        for c in g['other_year_candidates']:targets[c['invoice']['file']].add(c['invoice']['row'])
        for c in oi[g['invoice']]['candidates']:targets[c['file']].add(c['row'])
    source={};manifest=[]
    for file,rns in targets.items():
        p=P/file;manifest.append({'file':file,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
        for rn,r in rows(p):
            if rn==5:h={v:k for k,v in r.items() if v}
            if rn not in rns:continue
            get=lambda label:r.get(h[label]);k=invoice_key(get('Ký hiệu'),get('Số hóa đơn'))
            source[k]={'key':list(k),'date':str(invoice_day(get('Ngày hóa đơn'))),'order_code':code(get('Mã đơn hàng')),'buyer':get('Người mua hàng') or '', 'net':str(amount(get('Doanh số bán chưa thuế'))),'vat':str(amount(get('Tiền thuế GTGT'))),'payment':str(amount(get('Tổng tiền'))),'status':get('Trạng thái hóa đơn'),'cqt_status':get('Trạng thái gửi CQT'),'file':file,'row':rn}
    kt=json.loads((O/'ktoan-year-review-live.json').read_text(encoding='utf-8'))
    ktindex=defaultdict(list)
    for block in kt:
        for rn,vals in enumerate(block['values'][1:],2):
            if len(vals)>=2 and vals[0] and vals[1]:ktindex[invoice_key(*vals[:2])].append({'row':rn,'range':block['range']})
    reviewed=[];groups=[];arithmetic_false=[]
    for g in gs:
        n=g['invoice'];og=oi[n];members=set(og['rows']);mapped=set();confirmed=set();group_records=[]
        keys=sorted({tuple(c['invoice']['key']) for c in g['other_year_candidates']})
        for key in keys:
            s=source[key];candidates=sorted({c['row'] for c in g['other_year_candidates'] if tuple(c['invoice']['key'])==key})
            exact_code_rows=bycode.get(s['order_code'],[]) if s['order_code'] else []
            full=[rn for rn in exact_code_rows if num(ix[rn][9])==n and amount(ix[rn][6])==amount(s['vat']) and amount(ix[rn][7])==amount(s['payment'])]
            same_group=[rn for rn in full if rn in members]
            if len(exact_code_rows)==1 and len(same_group)==1:
                rn=same_group[0];v=ix[rn];known=rn==10 and v[2]=='BHHT2025040149' and key==('1C26MUV','00000106')
                state='KT_DA_XAC_NHAN_VA_KHOP_MA_TIEN' if known else 'KHOP_DUY_NHAT_MA_DON_SO_HD_VAT_THANH_TOAN'
                mapped.add(rn)
                if known:confirmed.add(rn)
            elif len(same_group)>0:state='TRUNG_MA_DON_NHIEU_CHUNG_TU_CAN_KIEM'
            elif s['order_code'] and not exact_code_rows:state='MA_DON_NGUON_KHONG_CO_TRONG_AMIS'
            elif exact_code_rows:state='MA_DON_CO_NHUNG_KHONG_KHOP_DU_SO_HD_VAT_THANH_TOAN'
            else:state='NGUON_KHONG_CO_MA_DON_CHI_KHOP_TIEN'
            rejected=[rn for rn in candidates if rn not in same_group]
            for rn in rejected:arithmetic_false.append({'group':n,'invoice':s['key'],'row':rn,'voucher':ix[rn][2],'amis_code':ix[rn][8],'source_code':s['order_code'],'reason':'Cùng số/VAT/thanh toán nhưng mã đơn không khớp hoặc nguồn thiếu mã'})
            rec={'group':n,'state':state,'source':s,'amount_only_rows':candidates,'source_order_amis_rows':exact_code_rows,'full_match_rows':same_group,'ktoan_existing':ktindex.get(key,[]),'rejected_amount_only_rows':rejected,'full_match_vouchers':[ix[rn][2] for rn in same_group]}
            reviewed.append(rec);group_records.append(rec)
        remain=members-mapped
        original_vat=sum(amount(ix[rn][6]) for rn in members);original_pay=sum(amount(ix[rn][7]) for rn in members)
        rv=sum(amount(ix[rn][6]) for rn in remain);rp=sum(amount(ix[rn][7]) for rn in remain)
        year2025=og['candidates']
        exact=[c for c in year2025 if amount(c['vat'])==rv and amount(c['payment'])==rp]
        paymatch=[c for c in year2025 if amount(c['payment'])==rp]
        residual='CON_LAI_KHOP_DU_VAT_THANH_TOAN' if len(exact)==1 else 'CON_LAI_KHOP_THANH_TOAN_VAT_CON_LECH' if paymatch else 'CON_LAI_VAN_CHENH_HOAC_CHUA_XAC_DINH_KY_HIEU'
        groups.append({'invoice_number':n,'original_rows':sorted(members),'original_vat':str(original_vat),'original_payment':str(original_pay),'strong_match_rows':sorted(mapped),'accountant_confirmed_rows':sorted(confirmed),'unconfirmed_strong_match_rows':sorted(mapped-confirmed),'remaining_rows':sorted(remain),'remaining_vat':str(rv),'remaining_payment':str(rp),'remaining_state':residual,'remaining_exact_invoice_candidates':exact,'remaining_payment_candidates':paymatch,'original_year_candidates':year2025,'reviewed':group_records})
    states=dict(Counter(r['state'] for r in reviewed));strong=[r for r in reviewed if r['state'] in ('KT_DA_XAC_NHAN_VA_KHOP_MA_TIEN','KHOP_DUY_NHAT_MA_DON_SO_HD_VAT_THANH_TOAN')]
    summary={'groups':len(groups),'different_year_invoices_reviewed':len(reviewed),'states':states,'amount_only_pairs_rejected':len(arithmetic_false),'strong_matches':len(strong),'accountant_confirmed':sum(r['state']=='KT_DA_XAC_NHAN_VA_KHOP_MA_TIEN' for r in reviewed),'strong_matches_not_yet_accountant_confirmed':sum(r['state']=='KHOP_DUY_NHAT_MA_DON_SO_HD_VAT_THANH_TOAN' for r in reviewed),'strong_match_invoices_missing_in_ktoan':sum(not r['ktoan_existing'] for r in strong),'remaining_states':dict(Counter(g['remaining_state'] for g in groups)),'read_only_review':True,'source_export_cutoff':'2026-10-29'}
    out={'summary':summary,'manifest':manifest,'groups':groups,'invoice_reviews':reviewed,'rejected_amount_only_pairs':arithmetic_false}
    (O/'year-signature-review.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
    with (O/'RA_SOAT_99_NHOM_NAM_KY_HIEU.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Số HĐ J','Kết quả','Ký hiệu nguồn','Ngày HĐ','Mã đơn nguồn','VAT HĐ','Thanh toán HĐ','Dòng AMIS khớp đủ mã và tiền','Chứng từ khớp','Đã có ở Ktoan','Dòng chỉ khớp tiền bị loại','File nguồn','Dòng nguồn'])
        for r in reviewed:
            s=r['source'];w.writerow([r['group'],r['state'],s['key'][0],s['date'],s['order_code'],s['vat'],s['payment'],';'.join(map(str,r['full_match_rows'])),';'.join(r['full_match_vouchers']),json.dumps(r['ktoan_existing']),';'.join(map(str,r['rejected_amount_only_rows'])),s['file'],s['row']])
    with (O/'SAU_KHI_TACH_UNG_VIEN_KHAC_NAM_99_NHOM.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Số HĐ J','Số chứng từ ban đầu','Dòng khớp mạnh khác năm','Dòng đã được KT xác nhận','Số chứng từ còn lại (mô phỏng)','VAT còn lại','Thanh toán còn lại','Kết quả còn lại','HĐ 2025 khớp đủ','HĐ 2025 chỉ khớp thanh toán'])
        for g in groups:w.writerow([g['invoice_number'],len(g['original_rows']),';'.join(map(str,g['strong_match_rows'])),';'.join(map(str,g['accountant_confirmed_rows'])),len(g['remaining_rows']),g['remaining_vat'],g['remaining_payment'],g['remaining_state'],json.dumps(g['remaining_exact_invoice_candidates'],ensure_ascii=False),json.dumps(g['remaining_payment_candidates'],ensure_ascii=False)])
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    for n in ['00000106','00000478','00007322']:
        g=next((g for g in groups if g['invoice_number']==n),None)
        if g:print(n,json.dumps({k:g[k] for k in ['strong_match_rows','remaining_payment','remaining_vat','remaining_state']},ensure_ascii=False))

if __name__=='__main__':main()
