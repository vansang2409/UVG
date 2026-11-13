"""Phân loại bằng chứng 5 cặp HĐ; không tự kết luận từ số tiền đơn thuần."""
from pathlib import Path
from collections import Counter,defaultdict
from decimal import Decimal
from datetime import datetime
import json,re,unicodedata,sys
sys.stdout.reconfigure(encoding='utf-8')
O=Path(__file__).parent
def read(n):return json.loads((O/n).read_text(encoding='utf-8'))
def norm(s):return ''.join(c for c in unicodedata.normalize('NFD',str(s or '').casefold()) if c.isalnum()).replace('đ','d')
def num(s):return None if s in (None,'') else Decimal(str(s))
def val(c):
    d=c.get('effectiveValue') or {};v=next(iter(d.values()),None)
    if v is not None:return v
    d=c.get('userEnteredValue') or {}
    return d.get('stringValue',d.get('numberValue',c.get('formattedValue')))
def grid(n):
    out={}
    for s in read(n)['sheets']:
        for g in s.get('data',[]):
            for i,r in enumerate(g.get('rowData',[])):out[g.get('startRow',0)+i+1]=[val(c) for c in r.get('values',[])]
    return out
def sig(lines):
    return Counter((norm(r['name']),norm(r['unit']),num(r['qty']),num(r['price']),num(r['goods']),norm(r['rate']),num(r['vat']),num(r['payment'])) for r in lines)
raw=read('five-invoice-deep-evidence.json');kt=grid('deep-live-kt-cells.json');nb=grid('deep-live-nb-cells.json');misa=[]
for p in O.glob('deep-misa-*.json'):misa+=json.loads(p.read_text(encoding='utf-8'))['hits']
cases=[];summary_rows=[];item_rows=[];amis_rows=[]
for c in raw['cases']:
    a,b=c['orphan'],c['kept'];ar,br=c['invoice_lines'][a],c['invoice_lines'][b];ha,hb=c['raw_headers'][a]['values'],c['raw_headers'][b]['values']
    assert sig(ar)==sig(br),(a,'Khác nội dung chi tiết hai HĐ')
    assert [ha[x] for x in ['P','Q','R']]==[hb[x] for x in ['P','Q','R']]
    assert ha['U']==hb['U']=='Hoá đơn mới'
    assert ha.get('E') and hb.get('E')
    source_link=[x for x in misa if x['v'][0]==c['doc'] and str(x['v'][1]) in c['order']];assert source_link
    current_kt=[]
    for rn,r in kt.items():
        if len(r)>10 and str(r[9])+'/'+str(r[10]).zfill(8) in [a,b]:current_kt.append({'row':rn,'values':r})
    assert len(current_kt)==2
    current_nb=[{'row':rn,'values':r} for rn,r in nb.items() if len(r)>8 and str(r[8]) in c['order']]
    if a.endswith(('00000900','00006823','00007204')):
        assert len(current_nb)==1 and str(int(current_nb[0]['values'][2])).zfill(8)==a.split('/')[1]
        conclusion='Khả năng cao xuất lại cho công ty';group='3 cặp nghi trùng cao'
        limitation='HĐ công ty không ghi mã đơn. Chưa thấy KT chốt xử lý HĐ khách lẻ trong các dòng đã đọc.'
    elif a.endswith('00032352'):
        oldkt=next(r for r in current_kt if r['values'][10]=='00032352');newkt=next(r for r in current_kt if r['values'][10]=='00032610')
        assert oldkt['values'][5]=='Điều chỉnh =0, do xuất trùng' and newkt['values'][5]=='xuất đúng'
        assert ha['F']==hb['F']==c['order'][0]
        conclusion='KT đã chốt HĐ khách lẻ xuất trùng';group='KT đã chốt trùng'
        limitation='KT đã ghi điều chỉnh 32352 về 0; trong nguồn 13/11 chưa thấy HĐ xử lý tương ứng.'
    else:
        assert a=='1C25MUV/00020896' and ha['F']==c['order'][0]
        conclusion='Chưa đủ căn cứ kết luận trùng';group='Cần xác minh thêm'
        limitation='BHH 08/07, HĐ CBC 09/07, HĐ khách lẻ 10/11. HĐ CBC không ghi mã đơn; KT cảnh báo NB trùng số HĐ khác ngày.'
    gap=(datetime.strptime(hb['D'],'%d/%m/%Y')-datetime.strptime(ha['D'],'%d/%m/%Y')).days
    notes=[]
    if a.endswith('00000900'):notes.append('AMIS phân bổ KA600 cao hơn HĐ 57.771đ, KA100 thấp hơn 57.771đ; tổng tiền/VAT khớp. 3 dòng phụ kiện/màng lọc trên HĐ để trống tiền, AMIS chỉ có 2 SKU máy.')
    if a.endswith('00006823'):notes.append('AMIS chia 1.497.128đ máy + 683.472đ lõi; HĐ ghi 2.180.600đ ở máy, dòng lõi trống tiền. HĐ ghi Filter KA350S, AMIS SKU KA350FILTERFULL.')
    sums=[sum(num(x['values'][k]) for x in c['amis_lines']) for k in ['Y','AP','AR']]
    assert sums==[num(ha[k]) for k in ['P','Q','R']]
    evidence='MISA A'+str(source_link[0]['row'])+':B'+str(source_link[0]['row'])+'; Ktoan '+', '.join(str(x['row']) for x in current_kt)
    if current_nb:evidence+='; NB '+str(current_nb[0]['row'])
    kt_notes='; '.join(str(r['values'][10])+': '+str(r['values'][5] or r['values'][6] or 'Chưa có ý kiến xử lý') for r in current_kt)
    summary_rows.append([a,b,int(ha['R']),conclusion,c['doc'],c['header_orphan']['day'],c['header_kept']['day'],str(len(ar))+'/'+str(len(br))+' dòng khớp sau chuẩn hóa khoảng trắng/dấu câu',hb['H'],c['order'][0],kt_notes,limitation+' '+' '.join(notes),evidence])
    for role,k in [('Khách lẻ',a),('Công ty',b)]:
        for x in c['invoice_lines'][k]:
            item_rows.append([c['doc'],role,k,x['date'],x['name'],float(num(x['qty'])) if num(x['qty']) is not None else None,x['unit'],float(num(x['price'])) if num(x['price']) is not None else None,float(num(x['goods'])) if num(x['goods']) is not None else None,float(num(x['vat'])) if num(x['vat']) is not None else None,float(num(x['payment'])) if num(x['payment']) is not None else None,x['rate'],x['file']+'!'+str(x['row'])])
    for x in c['amis_lines']:
        v=x['values'];amis_rows.append([c['doc'],v['P'],v['Q'],float(num(v['W'])),v['T'],float(num(v['X'])),float(num(v['Y'])),float(num(v['AP'])),float(num(v['AR'])),v['D']+'/'+v['F'],x['sheet']+'!'+str(x['row']),' '.join(notes)])
    cases.append({'orphan':a,'kept':b,'doc':c['doc'],'group':group,'conclusion':conclusion,'payment':int(ha['R']),'line_counts':[len(ar),len(br)],'normalized_line_signature_equal':True,'company_date_minus_consumer_days':gap,'current_kt':current_kt,'current_nb':current_nb,'misa_link':source_link,'limitation':limitation,'amis_notes':notes,'source_header_references':[c['raw_headers'][k]['file']+'!'+str(c['raw_headers'][k]['row']) for k in [a,b]]})
counts=Counter(c['group'] for c in cases);totals={g:sum(c['payment'] for c in cases if c['group']==g) for g in counts}
assert totals=={'3 cặp nghi trùng cao':7876581,'Cần xác minh thêm':606100,'KT đã chốt trùng':1239000}
result={'as_of':'2026-11-13','cases':cases,'group_counts':dict(counts),'group_payments':totals,'sum_payment':sum(c['payment'] for c in cases),'summary_rows':summary_rows,'item_rows':item_rows,'amis_rows':amis_rows,'sources':raw['sources'],'no_source_mutation':True}
def ser(v):return int(v) if isinstance(v,Decimal) and v==int(v) else float(v) if isinstance(v,Decimal) else str(v)
(O/'five-deep-findings.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=ser),encoding='utf-8')
print(json.dumps({'counts':dict(counts),'payments':totals,'sum':result['sum_payment'],'comparisons':[(c['orphan'],c['line_counts'],c['company_date_minus_consumer_days'],c['conclusion']) for c in cases]},ensure_ascii=False),flush=True)
