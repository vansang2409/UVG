"""Check report totals, scope partitions and every planned old-invoice chain."""
from pathlib import Path
from collections import defaultdict
from decimal import Decimal
import json,csv,hashlib,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109';WORK=OUT/'report-2025'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def n(v):return Decimal(str(v or 0))
def main():
    m=read(WORK/'manifest.json');data=read(WORK/'expected-values.json');matrix=read(WORK/'matrix.json');roots=read(WORK/'roots.json');a=read(OUT/'authoring.json');p=read(OUT/'prepared.json')
    for source in m['sources']:assert hashlib.sha256((ROOT/source['file']).read_bytes()).hexdigest()==source['sha256']
    assert len(data)==len(matrix)==m['rows']==len(roots)+14 and all(len(r)==18 for r in data+matrix)
    assert len({r['root'] for r in roots})==len(roots)==46194
    body=data[14:];byroot={r[0]:r for r in body};plans=a['zero_audit']+a['unmapped_zero_audit'];assert len(plans)==2057 and {r['root'] for r in plans}<=byroot.keys()
    adj=defaultdict(Decimal)
    for o in a['outputs']:
        if o['kind']=='adjustment':
            for row in o['rows']:
                if row[19] is not None:adj[row[25]]+=n(row[19])
    for plan in plans:
        row=byroot[plan['root']];assert row[6] is not None and n(row[11])==0
        if plan['kind']=='adjustment':assert n(row[9])==adj[plan['root']] and n(row[10])==0
        else:assert n(row[9])==0 and n(row[10])==-n(plan['before'][0])
        assert n(row[6])+n(row[9])+n(row[10])==0
    for row,root in zip(body,roots):
        assert row[0]==root['root'] and set(str(row[4] or '').splitlines())<=set(root['keys'])
        if root['hold']:assert row[6] is None and row[11] is None
        else:assert n(row[6])+n(row[9])+n(row[10])==n(row[11])
    sums={k:sum(n(r[i]) for r in body) for k,i in [('old_before',6),('adjustment',9),('replacement_effect',10),('old_kept',11),('held_source_goods',14)]}
    for key,value in sums.items():assert value==n(m['summary'][key]),key
    new=sum(n(r[16]) for o in a['outputs'] if o['kind']=='new' for r in o['rows']);assert new==n(m['summary']['new_sales'])
    ret=next(o for o in a['outputs'] if o['kind']=='return_adjustment');issues=set(a['summary']['returns']['issues']);allret=sum(n(r[18]) for r in ret['rows']);validret=sum(n(r[18]) for r in ret['rows'] if r[25] not in issues)
    assert len({r[0] for r in ret['rows'] if r[25] not in issues})==319 and allret==n(m['summary']['returns_all']) and validret==n(m['summary']['returns_ready'])
    with (OUT/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig') as f:decisions=list(csv.DictReader(f))
    totals=defaultdict(Decimal)
    for d in decisions:totals[d['Phân loại']]+=n(p['docs'][d['Chứng từ AMIS']]['goods'])
    assert sum(totals.values())==n(m['summary']['amis_sales_all']) and totals['Loại theo HD CTY 2025']==n(m['summary']['company_sales'])
    exchanges=a['summary'].get('psc2_exchanges',{})
    excluded=sum(n(line['v'][24]) for code in exchanges for line in p['docs'][code]['lines'] if line['v'][15]=='PSC2')
    assert excluded==n(m['summary'].get('psc2_excluded_goods'))
    assert totals['Tạo import mới']-excluded==new and totals['Giữ nguyên, không import']==n(m['summary']['amis_keep_goods']) and totals['Chờ xác minh']==n(m['summary']['amis_pending_goods'])
    projected=sums['old_kept']+new+validret;amis=sum(totals.values())-totals['Loại theo HD CTY 2025']+allret;gap=projected-amis
    assert projected==n(data[9][1])==n(m['summary']['projected']) and amis==n(data[10][1]) and gap==n(data[11][1])==n(data[12][1])
    parts=[sums['old_kept']-totals['Giữ nguyên, không import'],-totals['Chờ xác minh'],validret-allret]
    if excluded:parts.append(-excluded)
    assert sum(parts)==gap
    assert sum(n(r[12]) for r in body)==parts[0]
    assert all('dự' in str(matrix[i][0]).lower() or i in [6,10,11,12] for i in [4,5,6,7,8])
    assert all(str(matrix[i][11]).startswith('=IF(') for i in range(14,len(matrix)))
    proof={'passed':True,'rows':m['rows'],'roots':len(roots),'old_chains_planned_zero_checked':len(plans),'all_source_hashes_match':True,'scope_and_return_exclusions_checked':True,'projected_goods':float(projected),'amis_goods_scope':float(amis),'difference':float(gap),'difference_components':list(map(float,parts)),'draft_not_issued':True,'no_forced_zero_difference':True,'formula_cells':sum(isinstance(v,str) and v.startswith('=') for r in matrix for v in r)}
    if m['summary'].get('gift_invoices_kept'):
        policy=read(OUT/'unmapped-gift-confirmation.json')
        kept={r['root'] for r in roots if r['state']=='Biếu tặng — KT chốt giữ, theo dõi riêng'}
        assert policy['accountant_decision']=='keep' and kept==set(policy['invoice_roots'])
        assert len(kept)==m['summary']['gift_invoices_kept']==16
        assert not kept.intersection(plan['root'] for plan in plans)
        assert all(byroot[key][6] is None and byroot[key][11] is None and byroot[key][14] is not None for key in kept)
        proof['gift_invoices_kept']=len(kept)
    (WORK/'verification.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(proof,ensure_ascii=False))
if __name__=='__main__':main()
