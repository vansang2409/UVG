"""Attach AMIS UI sale-document links and draft return adjustments without guessing invoice numbers."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal,ROUND_HALF_UP
from datetime import date
import csv,json,hashlib,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def n(x):return Decimal(str(x or 0))
def save(name,x):(OUT/name).write_text(json.dumps(x,ensure_ascii=False,default=lambda x:float(x) if isinstance(x,Decimal) else str(x)),encoding='utf-8')
def main():
    p=json.loads((OUT/'prepared.json').read_text(encoding='utf-8'));a=json.loads((OUT/'authoring.json').read_text(encoding='utf-8'))
    with (OUT/'source/btl-links-ui.csv').open(encoding='utf-8',newline='') as f:raw=list(csv.DictReader(f))
    links={'BTL'+str(int(r['BTL'])).zfill(5):r['BHH'] for r in raw};assert len(links)==len(raw)==322
    assert set(links)=={r['doc'] for r in p['returns']}
    with (OUT/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig',newline='') as f:decisions={r['Chứng từ AMIS']:r for r in csv.DictReader(f)}
    newmap={}
    for o in a['outputs']:
        if o['kind']=='new':
            for r in o['rows']:newmap[r[18]]=(o['file'],r[0])
    by_btl=defaultdict(list);by_sale=defaultdict(list)
    for r in p['returns']:by_btl[r['doc']].append(r);by_sale[links[r['doc']]].append(r)
    assert all(sale in p['docs'] for sale in by_sale)
    checks=[];full=[];issues={};csvrows=[]
    for sale,rs in by_sale.items():
        d=p['docs'][sale];sales=defaultdict(lambda:[Decimal(0)]*3);rets=defaultdict(lambda:[Decimal(0)]*3)
        for l in d['lines']:
            r=l['v'];k=(r[15],r[19]);sales[k]=[x+y for x,y in zip(sales[k],[n(r[22]),n(r[24])-n(r[29]),n(r[41])])]
        for l in rs:
            r=l['v'];k=(r[15],r[19]);rets[k]=[x+y for x,y in zip(rets[k],[n(r[32]),n(r[36]),-n(r[41])])]
        reasons=[]
        for k,values in rets.items():
            if k not in sales:reasons.append('Mã hàng/ĐVT trả không có ở chứng từ bán')
            elif any(rv>sv for rv,sv in zip(values,sales[k])):reasons.append('Số lượng hoặc tiền trả vượt số bán cùng mã/ĐVT')
        is_full=rets==sales
        if is_full:full.append(sale)
        issues[sale]=reasons
        checks.append({'sale':sale,'return_docs':sorted({r['doc'] for r in rs}),'full_return':is_full,'reasons':reasons,'sales_action':decisions[sale]['Phân loại'],'goods':sum(n(r['v'][36]) for r in rs),'vat':-sum(n(r['v'][41]) for r in rs),'pay':-sum(n(r['v'][43]) for r in rs)})
    draft=[];audit=[]
    current=(date(2026,11,9)-date(1899,12,30)).days
    for btl,rs in sorted(by_btl.items()):
        sale=links[btl];d=p['docs'][sale];decision=decisions[sale];reasons=list(issues[sale]);new=newmap.get(sale)
        if decision['Phân loại']=='Loại theo HD CTY 2025':status='Loại công ty, không tạo import trả lại'
        elif reasons:status='Chờ kiểm tra số lượng/tiền trả'
        elif new:status='Chờ số HĐ bán mới sau phát hành'
        else:status='Chờ đối chiếu trạng thái trả hàng trên HĐ đang giữ'
        for l in rs:
            r=l['v'];candidates=[rate for rate in [0,5,8,10] if abs((n(r[36])*rate/100).quantize(Decimal(1),rounding=ROUND_HALF_UP)+n(r[41]))<=1]
            rate=candidates[0] if len(candidates)==1 else None
            csvrows.append([btl,sale,l['source'],r[15],r[7] or r[16],r[19],r[32],r[23],r[36],r[41],r[43],new[0] if new else None,new[1] if new else None,status,'; '.join(reasons)])
        if decision['Phân loại']=='Loại theo HD CTY 2025':continue
        seq=len(audit)+1;goods=sum(n(l['v'][36]) for l in rs);vat=sum(n(l['v'][41]) for l in rs);pay=sum(n(l['v'][43]) for l in rs)
        assert -goods+vat==pay,btl
        for i,l in enumerate(rs):
            r=l['v'];candidates=[rate for rate in [0,5,8,10] if abs((n(r[36])*rate/100).quantize(Decimal(1),rounding=ROUND_HALF_UP)+n(r[41]))<=1];rate=candidates[0] if len(candidates)==1 else None
            reason=f'Trả hàng theo {btl}; chứng từ bán gốc {sale}'
            draft.append([seq,current,d['buyer'],r[12] or None,r[13] or None,r[14] or None,None,'TM/CK',None,None,None,None,None,reason,r[7] or r[16],r[19],-n(r[32]),n(r[23]),-n(r[36]),-goods if i==0 else None,rate,vat if i==0 else None,pay if i==0 else None,None,btl,sale,r[15],l['source'],new[0] if new else None,new[1] if new else None,status])
        audit.append({'btl':btl,'sale':sale,'sequence':seq,'rows':len(rs),'goods':-goods,'vat':vat,'pay':pay,'status':status})
    report=next(o for o in a['outputs'] if o['kind']=='report');report['sheets']=[s for s in report['sheets'] if s['name'] not in ['BTL cho lien ket','Lien ket BTL','Tra lai cho so HD']]
    lh=['BTL','Chứng từ bán gốc','Dòng nguồn','Mã hàng','Tên hàng','ĐVT','SL trả','Đơn giá trả','Tiền hàng trả','VAT trả','Thanh toán trả','File import bán mới','STT HĐ bán mới','Trạng thái','Ghi chú']
    rh=['Số thứ tự hóa đơn (*)','Ngày hóa đơn','Tên khách hàng','Địa chỉ','Mã số thuế','Người mua hàng','Email','Hình thức thanh toán (*)','HĐ bị điều chỉnh thuộc hệ thống khác','Ký hiệu HĐ bị điều chỉnh (*)','Số hóa đơn bị điều chỉnh (*)','Ngày hóa đơn bị điều chỉnh','Mã của CQT trên HĐ bị điều chỉnh','Lý do điều chỉnh','Tên hàng hóa/dịch vụ (*)','ĐVT','Số lượng','Đơn giá','Thành tiền','Tổng tiền hàng','Thuế suất GTGT (%)','Tổng tiền thuế GTGT','Tổng tiền thanh toán','Hàng khuyến mãi','BTL','Chứng từ bán gốc','Mã hàng','Dòng nguồn','File import bán mới','STT HĐ bán mới','Trạng thái']
    report['sheets'] += [{'name':'Lien ket BTL','headers':lh,'rows':csvrows},{'name':'Tra lai cho so HD','headers':rh,'rows':draft,'start':9,'kind':'adjustment','notes':['DỰ THẢO TRẢ LẠI THEO PHIẾU AMIS 2025','Liên kết số chứng từ bán gốc đã đọc trực tiếp trên AMIS ngày 09/11/2026.','Chưa nhập: cột J:M trống vì phải xác minh/điền HĐ thực tế cần điều chỉnh.','Phần bán gốc giữ nguyên. Không nhập phiếu trả hai lần nếu HĐ hiện hành đã xử lý.']}]
    a['outputs']=[o for o in a['outputs'] if o['kind']!='return_adjustment']
    a['outputs'].append({'file':'IMPORT_DIEU_CHINH_BTL_2025_CHO_SO_HD.xlsx','kind':'return_adjustment','headers':rh,'rows':draft,'invoices':len(audit),'notes':[
        f'{len(audit)} phiếu/{len(draft)} dòng trả lại; đã loại phiếu thuộc danh sách công ty không xuất meInvoice.',
        'CHƯA NHẬP: cột J:M trống, phải xác minh và điền ký hiệu/số/ngày/mã CQT của hóa đơn thực tế cần điều chỉnh.',
        '; '.join(f'{count} phiếu: {status}' for status,count in Counter(x['status'] for x in audit).items())+'.',
        'BTL00449 và BTL00488 cần kiểm tra riêng; không tự sửa số tiền hoặc coi phần vượt là giảm giá được duyệt.',
        'Cột AC:AD là file/STT import bán mới để tra cứu, không phải số hóa đơn đã phát hành. Ngày hóa đơn dự thảo 09/11/2026.'
    ]})
    for o in a['outputs']:
        if o['kind']=='new':o['notes'][2]='Giữ nguyên bán gốc. Phiếu trả lại đã nối số CT bán gốc trong báo cáo; phần trả lập điều chỉnh riêng sau khi xác định số HĐ thực tế.'
    a['summary']['limitations'][0]='Đã nối đủ 322 phiếu/412 dòng trả hàng trực tiếp trên AMIS; dự thảo trả lại còn chờ số HĐ thực tế, không tự điền số tương lai.'
    a['summary']['returns']={'docs':322,'lines':412,'sales':len(by_sale),'full_return_sales':len(full),'draft_docs':len(audit),'draft_lines':len(draft),'draft_statuses':dict(Counter(x['status'] for x in audit)),'issues':{k:v for k,v in issues.items() if v},'source_goods':sum(n(r['v'][36]) for r in p['returns']),'source_vat':sum(n(r['v'][41]) for r in p['returns']),'source_pay':sum(n(r['v'][43]) for r in p['returns'])}
    a['summary']['returns']['import_file']='IMPORT_DIEU_CHINH_BTL_2025_CHO_SO_HD.xlsx'
    a['summary']['delivery_files']=len(a['outputs'])
    files={m['file'] for m in a['manifest']}
    cache=json.loads((OUT/'source-index.json').read_text(encoding='utf-8'))
    candidates=[ROOT/r['file'] for r in cache['files']]
    candidates += [ROOT/'scripts/data/post-issue-20261029/audit.json',ROOT/'data/Du lieu Meinvoice/Bang_ke_chi_tiet_HD_da_su_dung_2025.xls',ROOT/'data/Du lieu Meinvoice/Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']
    candidates += list((OUT/'source').glob('*.json'))+[OUT/'source/btl-links-ui.csv',OUT/'source/Tra_lai_hang_ban_20261109.xlsx']
    for path in candidates:
        key=str(path.relative_to(ROOT))
        if key not in files:
            a['manifest'].append({'file':key,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()});files.add(key)
    with (OUT/'LIEN_KET_BTL_BHH_2025.csv').open('w',encoding='utf-8-sig',newline='') as f:w=csv.writer(f);w.writerow(lh);w.writerows(csvrows)
    save('return-verification.json',{'linkage_passed':True,'amount_checks_passed':not any(issues.values()),'amount_exception_sales':{k:v for k,v in issues.items() if v},'checks':checks,'draft_audit':audit,'full_return_sales':full,'source_ui_docs':322,'source_ui_lines':412})
    save('authoring.json',a);save('final-summary.json',a['summary']);print(json.dumps(a['summary']['returns'],ensure_ascii=False,default=str))
if __name__=='__main__':main()
