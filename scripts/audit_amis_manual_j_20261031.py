"""Read-only audit of manually supplied AMIS invoice numbers; no fuzzy assignment."""
import json, csv, hashlib, re
from pathlib import Path
from collections import defaultdict, Counter
from datetime import date
from prepare_invoice_reports import rows, invoice_key, invoice_day, amount

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/doi-chieu-amis-j-20261031'

def main():
    snapshot = json.loads((OUT/'amis-j-snapshot.json').read_text(encoding='utf-8'))
    records = {}
    manifest = []
    for p in sorted((ROOT/'data/Meinv_SAU_PHAT_HANH_20261029').rglob('*.xls')):
        if p.parent.name in ('DIEU_CHINH','THAY_THE'): continue
        manifest.append({'file':p.relative_to(ROOT).as_posix(),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
        headers = {}
        for rn,r in rows(p):
            if rn == 5: headers = {v:k for k,v in r.items() if v}
            if rn < 6 or amount(r.get('A')) is None: continue
            get=lambda s:r.get(headers[s])
            key=invoice_key(get('Ký hiệu'),get('Số hóa đơn'))
            rec={'key':list(key),'date':str(invoice_day(get('Ngày hóa đơn'))),'vat':str(amount(get('Tiền thuế GTGT'))),'payment':str(amount(get('Tổng tiền'))),'status':get('Trạng thái hóa đơn'),'file':p.relative_to(ROOT).as_posix(),'row':rn}
            if key in records: assert all(records[key][c]==rec[c] for c in ('date','vat','payment','status')), key
            records[key]=rec
    bynum=defaultdict(list)
    for rec in records.values(): bynum[rec['key'][1]].append(rec)
    groups=defaultdict(list)
    for item in snapshot['rows']:
        v=item['values']; n=invoice_key('?',v[9])[1]
        groups[(invoice_day(v[0]).year, str(v[10] or '').strip(), n)].append(item)
    result=[]
    for (year,sig,n),items in groups.items():
        candidates=[r for r in bynum[n] if r['date'].startswith(str(year)) and (not sig or r['key'][0]==sig)]
        other_year_candidates=[r for r in bynum[n] if not r['date'].startswith(str(year)) and (not sig or r['key'][0]==sig)]
        vat=sum(amount(x['values'][6]) for x in items)
        pay=sum(amount(x['values'][7]) for x in items)
        exact=[r for r in candidates if amount(r['vat'])==vat and amount(r['payment'])==pay]
        near=[r for r in candidates if amount(r['vat']) is not None and amount(r['payment']) is not None and abs(amount(r['vat'])-vat)<=2 and abs(amount(r['payment'])-pay)<=2]
        state='GHI_CHU_KHONG_PHAI_SO_HD' if not n.isdigit() else 'KHOP_TONG_DUY_NHAT' if len(exact)==1 else 'NHIEU_UNG_VIEN_KHOP' if len(exact)>1 else 'LECH_TOI_DA_2D' if len(near)==1 else 'LECH_TIEN' if candidates else 'KHONG_TIM_THAY_CUNG_NAM'
        chosen=exact[0] if len(exact)==1 else near[0] if len(near)==1 else None
        dates=sorted({str(invoice_day(x['values'][0])) for x in items})
        result.append({'year':year,'signature_entered':sig,'invoice':n,'state':state,'rows':[x['row'] for x in items],'vouchers':[x['values'][2] for x in items],'amis_dates':dates,'amis_vat':str(vat),'amis_payment':str(pay),'matched':chosen,'same_date':bool(chosen and dates==[chosen['date']]),'candidates':candidates,'other_year_candidates':other_year_candidates})
    summary={'manual_rows':len(snapshot['rows']),'groups':len(result),'source_invoices':len(records),'states_groups':dict(Counter(r['state'] for r in result)),'states_rows':dict(Counter({s:sum(len(r['rows']) for r in result if r['state']==s) for s in {r['state'] for r in result}})),'exact_same_date_groups':sum(r['state']=='KHOP_TONG_DUY_NHAT' and r['same_date'] for r in result)}
    special=[]
    for g in result:
        if g['state']!='GHI_CHU_KHONG_PHAI_SO_HD':continue
        dated=re.fullmatch(r'(\d+)\s*-\s*(\d{2}/\d{2}/\d{4})',g['invoice'])
        zero=re.fullmatch(r'(\d+)\s*-\s*0',g['invoice'])
        multi=re.fullmatch(r'(\d+)\+(\d+)',g['invoice'])
        refs=[dated[1]] if dated else [zero[1]] if zero else [multi[1],multi[2]] if multi else []
        found=[r for n in refs for r in bynum[n.zfill(8)] if (r['date']==str(invoice_day(dated[2])) if dated else r['date'].startswith(str(g['year'])))]
        exact=[r for r in found if amount(r['vat'])==amount(g['amis_vat']) and amount(r['payment'])==amount(g['amis_payment'])]
        special.append({'entry':g['invoice'],'rows':g['rows'],'kind':'dated_reference' if dated else 'zero_annotation' if zero else 'multiple_invoices' if multi else 'text_note','amis_vat':g['amis_vat'],'amis_payment':g['amis_payment'],'references':found,'exact_single_candidates':exact})
    (OUT/'audit.json').write_text(json.dumps({'summary':summary,'manifest':manifest,'groups':result,'special_entries':special},ensure_ascii=False,indent=2),encoding='utf-8')
    with (OUT/'DOI_CHIEU_THEO_NHOM.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f); w.writerow(['Ket qua','Nam AMIS','Ky hieu da dien','So HD J','Dong AMIS','Chung tu','Ngay AMIS','VAT AMIS','Thanh toan AMIS','Ung vien meInvoice - ky hieu/ngay/VAT/thanh toan/trang thai'])
        for r in result:w.writerow([r['state'],r['year'],r['signature_entered'],r['invoice'],';'.join(map(str,r['rows'])),';'.join(r['vouchers']),';'.join(r['amis_dates']),r['amis_vat'],r['amis_payment'],json.dumps(r['candidates'],ensure_ascii=False)])
    print(json.dumps(summary,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
