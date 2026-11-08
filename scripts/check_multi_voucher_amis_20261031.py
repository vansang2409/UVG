"""Sum the 417 numeric-J groups with multiple distinct AMIS vouchers."""
import csv,json,hashlib
from collections import Counter
from decimal import Decimal as D
from pathlib import Path
P=Path(__file__).resolve().parents[1]
O=P/'outputs/doi-chieu-amis-j-20261031'

def main():
    audit=json.loads((O/'audit.json').read_text(encoding='utf-8'))
    snap=json.loads((O/'amis-j-snapshot.json').read_text(encoding='utf-8'))
    live=json.loads((O/'live-c-l-recheck.json').read_text(encoding='utf-8'))
    ix={r['row']:r['values'] for r in snap['rows']}
    groups=[g for g in audit['groups'] if g['invoice'].isdigit() and len(set(g['vouchers']))>1]
    assert len(groups)==417
    for m in audit['manifest']:assert hashlib.sha256((P/m['file']).read_bytes()).hexdigest()==m['sha256']
    findings=[]
    for g in groups:
        values=[ix[rn] for rn in g['rows']]
        for rn in g['rows']:
            for col in (2,6,7,8,9,10):assert live[rn-1][col-2]==ix[rn][col],(rn,col,'live drift')
        assert len({v[2] for v in values})==len(values),('duplicate voucher',g['invoice'])
        vat=sum(D(str(v[6])) for v in values); pay=sum(D(str(v[7])) for v in values)
        assert str(vat)==g['amis_vat'] and str(pay)==g['amis_payment']
        exact_pay=[c for c in g['candidates'] if c['payment']!='None' and D(c['payment'])==pay]
        record={**g,'payment_exact_candidates':exact_pay}
        findings.append(record)
    summary={'groups':len(groups),'vouchers':sum(len(g['rows']) for g in groups),'both_vat_payment':dict(Counter(g['state'] for g in groups)),'payment_only_exact_groups':sum(bool(g['payment_exact_candidates']) for g in findings),'payment_only_not_exact_groups':sum(not g['payment_exact_candidates'] for g in findings),'live_recheck_unchanged_rows':sum(len(g['rows']) for g in groups),'source_hashes_verified':len(audit['manifest'])}
    with (O/'DOI_CHIEU_417_HD_NHIEU_CHUNG_TU.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Số HĐ J','Số chứng từ','Dòng AMIS','Chứng từ AMIS','VAT cộng','Thanh toán cộng','Khớp thanh toán ứng viên nào','Kết quả cả VAT/thanh toán','Các HĐ nguồn: ký hiệu, ngày, VAT, thanh toán, trạng thái, file, dòng'])
        for g in findings:w.writerow([g['invoice'],len(g['rows']),';'.join(map(str,g['rows'])),';'.join(g['vouchers']),g['amis_vat'],g['amis_payment'],';'.join('/'.join(c['key']) for c in g['payment_exact_candidates']),g['state'],json.dumps(g['candidates'],ensure_ascii=False)])
    with (O/'CHI_TIET_CHENH_417_HD.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Số HĐ J','Số chứng từ','Ký hiệu meInvoice','Ngày HĐ','VAT AMIS cộng','VAT HĐ','Chênh VAT AMIS - HĐ','Thanh toán AMIS cộng','Thanh toán HĐ','Chênh thanh toán AMIS - HĐ','Trạng thái nguồn','File nguồn','Dòng nguồn'])
        for g in findings:
            for c in g['candidates']:
                w.writerow([g['invoice'],len(g['rows']),c['key'][0],c['date'],g['amis_vat'],c['vat'],str(D(g['amis_vat'])-D(c['vat'])) if c['vat']!='None' else '',g['amis_payment'],c['payment'],str(D(g['amis_payment'])-D(c['payment'])) if c['payment']!='None' else '',c['status'],c['file'],c['row']])
    (O/'multi-voucher-verification.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    (O/'multi-voucher-audit.json').write_text(json.dumps({'summary':summary,'groups':findings},ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    for n in ['00000427','00000431','00000465','00000428','00000489']:
        g=next(g for g in findings if g['invoice']==n)
        c=next(c for c in g['candidates'] if c['key'][0]=='1C25TUV')
        print(n,len(g['rows']),g['amis_payment'],c['payment'],str(D(g['amis_payment'])-D(c['payment'])),'VAT',str(D(g['amis_vat'])-D(c['vat'])))

if __name__=='__main__':main()
