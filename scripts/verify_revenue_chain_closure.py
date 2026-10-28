"""Đối chiếu độc lập bảng kê vừa nhận, tập hóa đơn và tổng dự tính."""
from pathlib import Path
from collections import defaultdict
import json,hashlib
from prepare_invoice_reports import rows,invoice_key

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/revenue-chain-audit-20261028'
filename='Bang_ke_hoa_don_thay_the_xuat_theo_HD_bi_thay_the_1793158686512.xls'
src=ROOT/'data/Du lieu Meinvoice'/filename
original=Path('C:/Users/tranq/Downloads')/filename
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
assert sha(src)==sha(original),'Bản nguồn lưu trong data khác file người dùng gửi'
raw=list(rows(src))
raw_edges={(invoice_key(r.get('B'),r.get('D')),invoice_key(r.get('M'),r.get('O'))) for rn,r in raw if 7<=rn<=20}
audit=json.loads((WORK/'prepared.json').read_text(encoding='utf8'))
summary=json.loads((WORK/'summary.json').read_text(encoding='utf8'))
forecast=json.loads((ROOT/'scripts/data/revenue-forecast-20261028/prepared.json').read_text(encoding='utf8'))
assert not audit['issues'] and not summary['provisional']
chain={tuple(c['root']):c for c in audit['chains']}
keys=set();grouped=defaultdict(lambda:[0,0,0])
for r in forecast['old_invoices']:
    k=tuple(r['key'])
    assert k not in keys
    keys.add(k)
    for i in range(3):grouped[tuple(r['root'])][i]+=r['money'][i]
assert keys=={tuple(k) for c in audit['chains'] for k in c['included']}
for rt,money in grouped.items():assert chain[rt]['before']==money,(rt,money,chain[rt]['before'])
matrix=[]
for a,b,c in [('80','188','2601'),('87','189','2610'),('116','190','2602'),('117','191','2603'),('118','192','2604'),('184','216','2617')]:
    first=('1C24TUV',a.zfill(8));mid=('1C24TUV',b.zfill(8));last=('1C25TUV',c.zfill(8))
    assert (first,mid) in raw_edges and (mid,last) in raw_edges
    current=chain[first]
    assert current['included']==[list(last)]
    assert {tuple(r['key']) for r in current['excluded']}=={first,mid}
    manual=[r for r in forecast['manual_records'] if tuple(r['root'])==first]
    assert len(manual)==1 and manual[0]['after']==current['before']
    matrix.append({'chain':[first,mid,last],'money':current['before'],'manual_reduction':[-v for v in current['before']]})
t=forecast['totals']
assert t['before']==summary['before']==[47402157888,3792174626,51194328614]
after=[t['before'][i]+t['adjustments'][i]+t['manual_replacement_reduction'][i]+t['new_gross'][i]+t['new_returns'][i] for i in range(3)]
assert after==t['forecast_after_manual']==[49930540760,3994443568,53924984328]
processed_four=[]
for n in ['00000003','00000478','00000498','00000574']:
    rt=('1C24TUV',n);original=chain[rt]['before']
    rs=[r for r in forecast['adjustment_records'] if tuple(r['root'])==rt]
    reduction=[sum(r['money'][i] for r in rs) for i in range(3)]
    assert [original[i]+reduction[i] for i in range(3)]==[0,0,0]
    processed_four.append({'root':rt,'before':original,'reduction':reduction,'after':[0,0,0]})
assert [sum(r['money'][i] for r in matrix) for i in range(3)]==[371344277,29707543,401051820]
result={'source':str(src.relative_to(ROOT)),'sha256':sha(src),'source_copy_identical':True,'chain_count':len(chain),'contributing_invoices':len(keys),'unresolved_chains':0,'six_chains':matrix,'four_chains_zeroed':processed_four,'before':t['before'],'after_plan':after,'amis_gap':t['after_all_minus_amis'],'scope':'Gốc năm 2024 và các điều chỉnh/thay thế liên quan năm sau; số sau là dự tính','status':'passed'}
(WORK/'closure-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(result,ensure_ascii=False))
