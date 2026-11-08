"""Bounded evidence audit; arithmetic candidates never become confirmed mappings."""
import sys,json,csv,re,hashlib
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal as D, ROUND_HALF_UP
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount

ROOT=Path(__file__).resolve().parents[1]
O=ROOT/'outputs/doi-chieu-amis-j-20261031'
def rounded(x):return x.quantize(D('1'),rounding=ROUND_HALF_UP)
def main():
    audit=json.loads((O/'multi-voucher-audit.json').read_text(encoding='utf-8'))
    allrows=json.loads((O/'amis-all-snapshot.json').read_text(encoding='utf-8'))
    sn=json.loads((O/'amis-j-snapshot.json').read_text(encoding='utf-8'))
    ix={r['row']:r['values'] for r in sn['rows']}
    numbers={g['invoice'] for g in audit['groups']}
    totals={};edges=defaultdict(list);detail=defaultdict(list);manifest=[]
    base=ROOT/'data/Meinv_SAU_PHAT_HANH_20261029'
    for p in sorted(base.rglob('*.xls')):
        rel=p.relative_to(ROOT).as_posix();manifest.append({'file':rel,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
        kind='adjustment' if p.parent.name=='DIEU_CHINH' else 'replacement' if p.parent.name=='THAY_THE' else 'total'
        h={}
        for rn,r in rows(p):
            if rn==5 and kind=='total':h={v:k for k,v in r.items() if v}
            if rn<(6 if kind=='total' else 7) or amount(r.get('A')) is None:continue
            if kind=='total':
                get=lambda label:r.get(h[label]);k=invoice_key(get('Ký hiệu'),get('Số hóa đơn'))
                if k[1] not in numbers:continue
                totals[k]={'key':list(k),'date':str(invoice_day(get('Ngày hóa đơn'))),'net':str(amount(get('Doanh số bán chưa thuế'))),'vat':str(amount(get('Tiền thuế GTGT'))),'payment':str(amount(get('Tổng tiền'))),'status':get('Trạng thái hóa đơn'),'file':rel,'row':rn}
            else:
                old=invoice_key(r.get('B'),r.get('D'));new=invoice_key(r.get('M'),r.get('O'))
                edges[old].append({'kind':kind,'key':list(new),'date':str(invoice_day(r.get('N'))),'net':str(amount(r.get('U'))),'vat':str(amount(r.get('V'))),'payment':str(amount(r.get('W'))),'file':rel,'row':rn})
    print('Loaded total and linkage sources',flush=True)
    for fn in ['Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
        p=ROOT/'data/Du lieu Meinvoice'/fn
        manifest.append({'file':p.relative_to(ROOT).as_posix(),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
        key=None
        for rn,r in rows(p):
            if rn<7:continue
            if r.get('B'):key=invoice_key(r.get('D'),r.get('B'))
            if not key or key[1] not in numbers or not r.get('L'):continue
            detail[key].append({'row':rn,'name':r.get('L'),'qty':r.get('N'),'net':str((amount(r.get('P')) or 0)-(amount(r.get('R')) or 0)),'vat':str(amount(r.get('T'))),'payment':str(amount(r.get('U'))),'rate':r.get('S'),'file':p.relative_to(ROOT).as_posix()})
    findings=[];evidence=[]
    for g in sorted(audit['groups'],key=lambda g:int(g['invoice'])):
        n=g['invoice'];vs=[(rn,ix[rn]) for rn in g['rows']]
        # Reproduce the comparison column, not an authoritative invoice assignment.
        cs=sorted(g['candidates'],key=lambda c:(abs(D(g['amis_payment'])-D(c['payment'])),abs(D(g['amis_vat'])-D(c['vat']))))
        c=cs[0];key=tuple(c['key']);dp=D(g['amis_payment'])-D(c['payment']);dv=D(g['amis_vat'])-D(c['vat'])
        months=Counter(str(invoice_day(v[0]))[:7] for _,v in vs)
        periods=defaultdict(lambda:{'rows':[],'payment':D(0),'vat':D(0)})
        for rn,v in vs:
            z=periods[str(invoice_day(v[0]))[:7]];z['rows'].append(rn);z['payment']+=D(str(v[7]));z['vat']+=D(str(v[6]))
        direct=edges.get(key,[])
        ds=detail.get(key,[])
        detail_complete=bool(ds) and all(d['vat']!='None' and d['payment']!='None' for d in ds)
        sums={k:sum(D(d[k]) for d in ds) for k in ('net','vat','payment')} if detail_complete else None
        source_consistent=bool(sums and sums['vat']==D(c['vat']) and sums['payment']==D(c['payment']))
        amis8=sum(D(str(v[6]))==rounded(D(str(v[7]))*D(8)/D(108)) for _,v in vs)
        source8=sum(d['vat']!='None' and D(d['vat'])==rounded(D(d['net'])*D('0.08')) and str(d['rate']).strip() in ('8%','8') for d in ds)
        # Identify exact row explanations; these remain candidates without order evidence.
        single=[{'row':rn,'voucher':v[2],'vat':v[6],'payment':v[7]} for rn,v in vs if D(str(v[6]))==dv and D(str(v[7]))==dp]
        other_year=[]
        for rn,v in vs:
            for k,t in totals.items():
                if k[1]==n and not t['date'].startswith('2025') and t['vat']!='None' and t['payment']!='None' and D(t['vat'])==D(str(v[6])) and D(t['payment'])==D(str(v[7])):
                    other_year.append({'row':rn,'voucher':v[2],'amis_date':str(invoice_day(v[0])),'invoice':t})
        omitted=[]
        members=set(g['rows'])
        for rn,v in enumerate(allrows[1:],2):
            j=str(v[9]).strip()
            if rn not in members and j.isdigit() and j.zfill(8)==n:
                omitted.append({'row':rn,'voucher':v[2],'date':str(invoice_day(v[0])),'signature':v[10],'vat':v[6],'payment':v[7]})
        basecat='Khớp đủ VAT/thanh toán' if dp==dv==0 else 'Khớp thanh toán; VAT khác cách cộng' if dp==0 else 'Chưa đủ căn cứ giải thích chênh thanh toán'
        verified=[];hyp=[]
        if dp==0 and dv!=0:
            verified.append(f'Thanh toán khớp; VAT AMIS - HĐ = {dv}đ. AMIS có {amis8}/{len(vs)} chứng từ đúng ROUND(thanh toán*8/108).')
            if source_consistent and amis8==len(vs) and source8==len(ds):
                basecat='Tái hiện được chênh VAT do làm tròn khác cấp'
                verified.append(f'{len(ds)} dòng meInvoice đều đúng ROUND(tiền hàng*8%); cộng chi tiết khớp bảng kê. Hai cách tính tái hiện đúng chênh VAT.')
            else:hyp.append('Chưa đủ dữ liệu chi tiết AMIS để kết luận VAT lệch chỉ do làm tròn.')
        if len(months)>1:
            verified.append('Các chứng từ chung số J thuộc nhiều tháng: '+', '.join(f'{m} ({ct} CT)' for m,ct in sorted(months.items())))
            hyp.append('Có nguy cơ gom các HĐ khác ký hiệu/năm; chưa tự tách chỉ theo ngày hoặc tiền.')
        if other_year:
            verified.append(f'{len(other_year)} lượt chứng từ khớp riêng VAT/thanh toán với HĐ cùng số năm khác; là ứng viên, chưa phải liên kết đã xác nhận.')
            if dp!=0:basecat='Cần phân biệt số HĐ trùng năm/ký hiệu'
        if single:hyp.append(f'{len(single)} chứng từ có VAT/thanh toán bằng đúng chênh; chưa có căn cứ bỏ khỏi nhóm.')
        if omitted:verified.append(f'{len(omitted)} dòng AMIS khác cũng có số J này ngoài nhóm đã cộng; giữ riêng vì khác ký hiệu hoặc thuộc luồng công thức.')
        if direct:verified.append(f'HĐ so sánh có {len(direct)} liên kết sửa trực tiếp trong nguồn mới; số đang so là HĐ gốc, không phải số dư chuỗi.')
        if ds:verified.append(f'Chi tiết meInvoice: {len(ds)} dòng, '+('tổng VAT/thanh toán khớp bảng kê.' if source_consistent else 'chưa đủ/không khớp bảng kê hiện hành, không dùng để ép khớp.'))
        if dp!=0:hyp.append('Chưa có chi tiết bán hàng AMIS 2025 theo mặt hàng/SL hoặc bảng kê gộp có mã chứng từ để xác định khoản thiếu/thừa. Tab AMIS không có chứng từ trả hàng âm; chưa kiểm được tác động hoàn/trả.')
        findings.append({'invoice':n,'compared_key':list(key),'category':basecat,'payment_delta':str(dp),'vat_delta':str(dv),'verified':' '.join(verified),'unconfirmed':' '.join(hyp),'periods':{m:{'rows':z['rows'],'vat':str(z['vat']),'payment':str(z['payment'])} for m,z in periods.items()},'detail_rows':len(ds),'detail_sums':{k:str(v) for k,v in sums.items()} if sums else None,'detail_matches_current_total':source_consistent,'amis_round8_matches':amis8,'source_round8_matches':source8,'direct_links':direct,'other_year_candidates':other_year,'single_delta_candidates':single,'same_j_outside_group':omitted,'source_lines':ds})
    summary={'groups':len(findings),'categories':dict(Counter(g['category'] for g in findings)),'multi_month_groups':sum(len(g['periods'])>1 for g in findings),'groups_with_other_year_candidates':sum(bool(g['other_year_candidates']) for g in findings),'detail_matches_current_total':sum(g['detail_matches_current_total'] for g in findings),'source_files':len(manifest),'source_snapshot_cutoff':'2026-10-29','source_only_has_amis_summary_not_2025_goods_details':True}
    (O/'cause-investigation.json').write_text(json.dumps({'summary':summary,'manifest':manifest,'groups':findings},ensure_ascii=False,indent=2),encoding='utf-8')
    with (O/'NGUYEN_NHAN_VA_CAN_CU_417_NHOM.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Số HĐ J','HĐ so sánh chưa chốt','Phân loại','Chênh thanh toán','Chênh VAT','Đã xác minh','Chưa đủ căn cứ','Chi tiết nguồn khớp bảng tổng'])
        for g in findings:w.writerow([g['invoice'],'/'.join(g['compared_key']),g['category'],g['payment_delta'],g['vat_delta'],g['verified'],g['unconfirmed'],g['detail_matches_current_total']])
    print(json.dumps(summary,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
