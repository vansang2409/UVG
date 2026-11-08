"""Recover missing header payments only from explicit, matching detail-report cells."""
from pathlib import Path
from collections import Counter
from decimal import Decimal,ROUND_HALF_UP
import json,sys,hashlib
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def n(v):return Decimal(str(v)) if v is not None and str(v).strip() else None
def refine(a):
    d=json.loads((OUT/'unmapped-invoice-detail-review.json').read_text(encoding='utf-8'));records=a['invoice_details'];recovered=[]
    for m in d['manifest']:assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256']
    for k,entry in records.items():
        r=entry.get('record') or {};money=r.get('resolved_money',r.get('money',[None]*3))
        if money[2] is not None:continue
        lines=d['found'].get(k,[])
        if len(lines)!=1 or 'MTT' not in lines[0]['file']:continue
        line=lines[0];v=line['values']
        if not all(n(v.get(c)) is not None for c in ['W','Y','AA','AB']):continue
        if n(v['W'])-n(v['Y'])!=n(money[0]) or n(v['AA'])!=n(money[1]):continue
        if n(v['W'])-n(v['Y'])+n(v['AA'])!=n(v['AB']):continue
        entry['detail_verified_money']=[float(n(money[0])),float(n(money[1])),float(n(v['AB']))]
        entry['payment_evidence']={'file':line['file'],'row':line['row'],'column':'AB','value':v['AB'],'header_payment_retained':None};recovered.append(k)
    for r in a['components']:
        if r['status']!='review' or r['issues']!=['Số dư còn ô trống, không mặc định0']:continue
        money=[records[k].get('detail_verified_money',records[k]['record'].get('resolved_money',records[k]['record']['money'])) for k in r['active']]
        if any(n(v) is None for row in money for v in row):continue
        r['raw_balance']=r['balance'];b=[sum(n(row[i]) for row in money) for i in range(3)];r['balance']=[float(v) for v in b];r['issues']=[]
        r['balance_evidence']='Thanh toán còn trống bảng tổng được lấy từ ô AB bảng chi tiết khớp tiền hàng/VAT; không đổi nguồn'
        if all(v==0 for v in b):r['status']='already_zero';continue
        adjusted=any(records[k]['parents'] and any(k not in records[q]['replacement_children'] for q in records[k]['parents']) for k in r['keys'])
        r['kind']='adjustment' if adjusted else 'replacement' if len(r['keys'])==1 else None
        rates=[x for x in [0,5,8,10] if abs((b[0]*x/100).quantize(Decimal(1),rounding=ROUND_HALF_UP)-b[1])<=2]
        if b[0]+b[1]!=b[2] or not r['kind'] or len(rates)!=1:r['status']='review';r['issues']=['Cần xác minh số dư/thuế suất sau bổ sung chi tiết']
        else:r['status']='ready';r['rate']=rates[0]
    a['detail_payment_recoveries']=recovered;a['summary']=dict(Counter(r['status'] for r in a['components']));a['ready']=dict(Counter(r['kind'] for r in a['components'] if r['status']=='ready'))
    manifest={m['file']:m for m in a['source_files']};manifest.update({m['file']:m for m in d['manifest']});a['source_files']=list(manifest.values());return a
def main():
    path=OUT/'unmapped-meinvoice-audit.json';a=refine(json.loads(path.read_text(encoding='utf-8')))
    decision=OUT/'unmapped-gift-confirmation.json'
    if decision.exists():
        keep=set(json.loads(decision.read_text(encoding='utf-8'))['invoice_roots'])
        assert len(keep)==16
        for r in a['components']:
            if r['root'] in keep:
                assert r['status'] in ['review','kt_review']
                r['status']='kt_review';r['issues']=['Biếu tặng/không thu tiền; người dùng yêu cầu tách sheet hỏi lại kế toán']
        a['summary']=dict(Counter(r['status'] for r in a['components']))
    controls={m['file']:m for m in a['source_files']}
    inputs={OUT/'prepared.json',OUT/'source-index.json',ROOT/'data/0811/Ban_hang.xlsx',ROOT/'data/0811/So_chi_tiet_ban_hang.xlsx',OUT/'source/company.json'}
    inputs.update(ROOT/v['record']['source']['file'] for v in a['invoice_details'].values() if v.get('record'))
    inputs.update((ROOT/'data/Du lieu Meinvoice/Xuat_lai_2026-10-22').glob('*.xls'))
    inputs.update([ROOT/'data/Du lieu Meinvoice/Bang_ke_hoa_don_dieu_chinh_xuat_theo_HD_dieu_chinh_2025.xls',ROOT/'data/Du lieu Meinvoice/Bang_ke_hoa_don_thay_the_xuat_theo_HD_thay_the_2024_2025.xls'])
    for f in sorted(inputs):controls[str(f.relative_to(ROOT))]={'file':str(f.relative_to(ROOT)),'sha256':hashlib.sha256(f.read_bytes()).hexdigest()}
    a['source_files']=list(controls.values())
    path.write_text(json.dumps(a,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:a[k] for k in ['summary','ready','detail_payment_recoveries']},ensure_ascii=False))
if __name__=='__main__':main()
