"""Lập bảng rà chuỗi để người dùng xem từng hóa đơn được cộng/bị bỏ."""
from pathlib import Path
from collections import defaultdict
from datetime import datetime
import json
ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/revenue-chain-audit-20261028'
p=json.loads((WORK/'prepared.json').read_text(encoding='utf8'))
old=json.loads((ROOT/'scripts/data/revenue-forecast-20261028/prepared.json').read_text(encoding='utf8'))
records={'/'.join(r['key']):r for r in p['records']}
aliases={'/'.join(r['key']):'/'.join(r['root']) for r in p['root_aliases']}
def rootof(k):
    text='/'.join(k)
    return aliases.get(text,text)
plan=defaultdict(int)
for r in old['adjustment_records']:plan[rootof(r['root'])]+=r['money'][0]
manual={rootof(r['root']):-r['after'][0] for r in old['manual_records']}
edges=defaultdict(list)
link_sources=defaultdict(list)
for e in p['edges']:
    edges['/'.join(e['parent'])].append('/'.join(e['parent'])+(' — thay thế → ' if e['kind']=='replacement' else ' — điều chỉnh → ')+'/'.join(e['child']))
    for ev in e['evidence']:
        if '1793158686512' in ev['file']:
            link_sources['/'.join(e['parent'])].append('Liên kết: '+ev['file']+' | dòng '+str(ev['row']))
headers=['HĐ gốc','Ngày gốc','Loại chuỗi','Số HĐ trong chuỗi','HĐ được cộng','HĐ bị loại','DT trước xử lý','VAT trước','Thanh toán trước','ĐC cũ theo import','KT dự kiến đưa về 0','DT cũ giữ lại','Chênh so báo cáo trước','Kết quả rà','Giá trị gốc đang chờ (không cộng DT)','Liên kết trong chuỗi','Nguồn chứng cứ','Ghi chú']
rows=[]
for c in p['chains']:
    rt='/'.join(c['root']);pending=bool(c['issues']);included=['/'.join(k) for k in c['included']]
    excluded=['/'.join(r['key'])+': '+r['reason'] for r in c['excluded']]
    sourcekeys=included or [rt]
    source='\n'.join(k+' | '+records[k]['source']+' | dòng '+str(records[k]['row']) for k in sourcekeys)
    fresh_links=[s for k in c['nodes'] for s in link_sources['/'.join(k)]]
    if fresh_links:source+='\n'+'\n'.join(fresh_links)
    day=(datetime.fromisoformat(c['date'])-datetime(1899,12,30)).days
    baseline=[None]*3 if pending else c['before']
    rows.append([rt,day,'Thiếu HĐ thay thế' if pending else c['kind'],c['node_count'],'\n'.join(included),'\n'.join(excluded),*baseline,plan[rt],manual.get(rt,0),None,None if pending else c['difference'][0],'Cần kiểm tra' if pending else 'Đã đối chiếu',records[rt]['money'][0] if pending else None,'\n'.join(e for k in c['nodes'] for e in edges['/'.join(k)]),source,'; '.join(c['issues']) if pending else ''])
assert len(rows)==778 and not any(r[13]=='Cần kiểm tra' for r in rows)
assert sum(r[6] or 0 for r in rows)==47402157888
assert sum(r[9] or 0 for r in rows)==old['totals']['adjustments'][0]==-46407850480
assert sum(r[10] or 0 for r in rows)==-371344277
assert sum(r[14] or 0 for r in rows)==0
t=old['totals']
data={'headers':headers,'rows':rows,'source_count':len(p['manifest']),'unknown_roots':[r[0] for r in rows if r[13]=='Cần kiểm tra'],'known_before':t['before'][0],'known_after':t['forecast_after_manual'][0],'new_sales':t['new_gross'][0],'new_returns':t['new_returns'][0],'amis_net':t['amis_net'][0],'adjustment_zero':t['adjustments'][0],'replacement_zero':t['manual_replacement_reduction'][0],'retained':t['retained'][0]+t['manual_replacement_reduction'][0],'gap':t['after_all_minus_amis'][0],'warning':'Đã tính lại theo bộ import từ commit 9cb2cca: cả 4 cặp 3/508, 498/2140, 574/2189, 478/2578 đã được đưa về 0 trong file chuẩn bị. Chưa xác minh đã nhập/ký/phát hành thực tế.'}
(WORK/'review-sheet.json').write_text(json.dumps(data,ensure_ascii=False),encoding='utf8')
print(json.dumps({'chains':len(rows),'pending':data['unknown_roots'],'known_before':data['known_before'],'known_after':data['known_after']},ensure_ascii=False))
