from pathlib import Path
from collections import defaultdict,Counter
from zipfile import ZipFile
from openpyxl import load_workbook
import json,io,sys
sys.stdout.reconfigure(encoding='utf-8')
O=Path(__file__).parent;R=Path('D:/UVG')
def read(n):return json.loads((O/n).read_text(encoding='utf-8'))
a=read('reverse-chain-unmapped.json');records=read('mapping.json');byid={r['id']:r for r in records};gifts={r[0]:r for r in read('reverse-live-gifts.json')['values']}
assert len(gifts)==24
main=[];special=[];review=[]
headers=['HĐ chưa nối AMIS','Ngày HĐ gốc','Tiền hàng còn (đ)','VAT còn (đ)','Thanh toán còn (đ)','Nhóm / kết luận','BHH liên quan','HĐ hiện đang map AMIS','Mã đơn AMIS','Trạng thái HĐ','Trạng thái CQT','Toàn bộ chuỗi','Nguồn bảng kê / ghi chú']
for r in a['unmapped_nonzero']:
    assert not r['mapped_docs'] and not r['issues'] and len(r['keys'])==1 and r['active']==r['keys']
    k=r['root'];invoice=r['invoices'][0];hints=sorted(set(r['order_candidates']+r['historical_candidates']))
    docs=[d for d in hints if d in byid];mapped_keys=sorted({byid[d]['symbol']+'/'+byid[d]['invoice'] for d in docs if byid[d]['invoice']})
    orders=sorted({x for d in docs for x in byid[d]['order']})
    if k in gifts:
        g=gifts[k];group=g[8]+' — '+g[9];related='';selected='';note='Tab 24 HĐ biếu tặng; đã theo dõi riêng, không mặc định là sai/thiếu HĐ.'
        special.append([k,r['root_day'],*r['balance'],group,related,selected,'; '.join(r['orders']),invoice['status'],invoice['cqt_status'],' → '.join(r['keys']),invoice['source']+' | '+note])
    elif k=='1C24TUV/00000111':
        group='Khác kỳ — AMIS2023';note='BHH2023005090; HĐ lập năm 2024 nhưng đơn ghi nhận 2023. Bảng Doi chieu lech 2024 dòng 39.'
        special.append([k,r['root_day'],*r['balance'],group,'BHH2023005090','', '',invoice['status'],invoice['cqt_status'],' → '.join(r['keys']),invoice['source']+' | '+note])
    elif k=='1C25TUV/00001943':
        group='Ngoại lệ KT đã chốt';note='Tab Rà soát toàn bộ 2025 dòng 1551: ngoại lệ KT đã chốt không xử lý; giữ theo bằng chứng trước.'
        special.append([k,r['root_day'],*r['balance'],group,'','','',invoice['status'],invoice['cqt_status'],' → '.join(r['keys']),invoice['source']+' | '+note])
    else:
        assert len(docs)==1 and len(mapped_keys)==1 and k not in mapped_keys
        assert r['balance']==byid[docs[0]]['amis_money']
        main.append([k,r['root_day'],*r['balance'],'Cần rà soát HĐ còn tiền',byid[docs[0]]['doc'],'; '.join(mapped_keys),'; '.join(orders),invoice['status'],invoice['cqt_status'],' → '.join(r['keys']),invoice['source']+' | BHH hiện đã gắn HĐ khác; chưa kết luận HĐ trùng.'])
for r in a['review']:
    invoice=r['invoices'][0] if r['invoices'] else {};review.append([r['root'],r['root_day'],*r['balance'],'CQT không hợp lệ / cần kiểm','; '.join(r['order_candidates']+r['historical_candidates']),'','; '.join(r['orders']),invoice.get('status'),invoice.get('cqt_status'),' → '.join(r['keys']),'; '.join(r['issues'])])
assert len(main)==5 and len(special)==26 and len(review)==3
tot=lambda rows:[sum(r[i] for r in rows if r[i] is not None) for i in [2,3,4]]
assert tot(main)==[9001557,720124,9721681]
assert [x+y for x,y in zip(tot(main),tot(special))]==a['summary']['unmapped_nonzero_balances']
# Kiểm chứng độc lập toàn bộ 31 HĐ trực tiếp trong ZIP bảng kê mới bằng openpyxl.
archive=R/'data/1311/meinvoice/Bang_ke_hoa_don_da_su_dung_1794536037209.zip';wanted=defaultdict(dict)
for r in a['unmapped_nonzero']:
    rec=r['invoices'][0];member,row=rec['source'].rsplit('!',1);wanted[member][int(row)]=rec
checked=0
with ZipFile(archive) as z:
    for member,rs in wanted.items():
        w=load_workbook(io.BytesIO(z.read(member)),read_only=True,data_only=True)
        for rn,raw in enumerate(w.active.iter_rows(min_row=min(rs),max_row=max(rs),values_only=True),min(rs)):
            if rn not in rs:continue
            e=rs[rn];assert str(raw[1]).upper()+'/'+str(raw[2]).removesuffix('.0').zfill(8)==e['key']
            assert list(raw[15:18])==e['money'] and raw[20]==e['status'];checked+=1
        w.close()
assert checked==31
result={'summary':a['summary'],'headers':headers,'main':main,'special':special,'review':review,'groups':[
 ['Bán hàng cần rà soát',len(main),*tot(main)],
 ['24 HĐ theo dõi riêng',24,*tot([r for r in special if r[0] in gifts])],
 ['Khác kỳ AMIS2023',1,*tot([r for r in special if r[0]=='1C24TUV/00000111'])],
 ['Ngoại lệ kế toán',1,*tot([r for r in special if r[0]=='1C25TUV/00001943'])],
 ['Tổng chưa nối sổ AMIS2024/25',31,*a['summary']['unmapped_nonzero_balances']],
 ['CQT không hợp lệ — tách ngoài',3,*tot(review)]
 ],'independent_source_headers_verified':checked,'source':['data/1311/amis/So_chi_tiet_ban_hang (2024).xlsx','data/1311/amis/So_chi_tiet_ban_hang (2025).xlsx','data/1311/meinvoice — bảng tổng, chi tiết, điều chỉnh, thay thế ngày 13/11/2026','Xử lý HĐ 2025: tab Import HD moi; HĐ meInvoice giữ lại 2025; 24 HĐ biếu tặng','IMPORT MEINV 2024: Doi chieu lech 2024 dòng 39; Ktoan: I30482:AH30482']}
(O/'reverse-report-data.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'main':len(main),'main_totals':tot(main),'special':len(special),'review':len(review),'headers_verified':checked,'groups':result['groups']},ensure_ascii=False),flush=True)
