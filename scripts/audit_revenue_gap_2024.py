"""Truy nguyên chênh doanh thu giữa hóa đơn cũ giữ và AMIS loại, không sửa import."""
from pathlib import Path
from collections import defaultdict
import json,gzip

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/revenue-gap-20261028'
def load(p):return json.loads((ROOT/p).read_text(encoding='utf8'))
def key(k):return '/'.join(k)
def total(rs,field):return [sum(r[field][i] for r in rs) for i in range(3)]
def main():
    WORK.mkdir(parents=True,exist_ok=True)
    f=load('scripts/data/revenue-forecast-20261028/prepared.json')
    p=load('scripts/data/amis-2024-linked-20261026/prepared.json')
    old=json.loads(gzip.decompress((ROOT/'scripts/data/new-amis-2024-20261026/prepared.json.gz').read_bytes()))
    audit=load('scripts/data/revenue-chain-audit-20261028/prepared.json')
    aliases={key(r['key']):key(r['root']) for r in audit['root_aliases']}
    graph=defaultdict(set);entries=defaultdict(list)
    for e in old['exclusions']:
        canonical=aliases.get(e['key'],e['key']);entries[canonical].append(e)
        root='INV:'+canonical;graph[root]
        for code in e['codes']:
            graph[root].add('AMIS:'+code);graph['AMIS:'+code].add(root)
    manual={key(r['root']) for r in f['manual_records']}
    kept={key(r['root']):r for r in f['retained_roots'] if key(r['root']) not in manual}
    excluded={d['doc']:d for d in p['docs'] if d['category']=='Đã loại theo danh sách'}
    groups=[];seen=set();covered_amis=set();covered_inv=set()
    for start in sorted(graph):
        if start in seen:continue
        component=set();stack=[start]
        while stack:
            n=stack.pop()
            if n in component:continue
            component.add(n);stack.extend(graph[n]-component)
        seen|=component
        roots=sorted(n[4:] for n in component if n.startswith('INV:'))
        codes=sorted(n[5:] for n in component if n.startswith('AMIS:'))
        kr=[kept[k] for k in roots if k in kept]
        docs=[excluded[c] for c in codes if c in excluded]
        im=total(kr,'after');am=[sum(d['source_sums'][col] for d in docs) for col in ['goods','vat','payment']]
        covered_amis|={d['doc'] for d in docs};covered_inv|={key(r['root']) for r in kr}
        source_entries=[e for k in roots for e in entries[k]]
        groups.append({'roots':roots,'codes':codes,'missing_codes':[c for c in codes if c not in excluded],'invoice_money':im,'amis_money':am,'difference':[a-b for a,b in zip(im,am)],'buyers':list(dict.fromkeys(e['values'][4] for e in source_entries)),'exclusion_rows':[e['row'] for e in source_entries],'exclusion_values':[e['values'] for e in source_entries],'amis_docs':[{'doc':d['doc'],'date':d['date'],'rows':d['rows'],'money':[d['source_sums'][c] for c in ['goods','vat','payment']]} for d in docs]})
    assert covered_amis==set(excluded),set(excluded)-covered_amis
    outside=[r for k,r in kept.items() if k not in covered_inv]
    company_diff=[sum(g['difference'][i] for g in groups) for i in range(3)]
    outside_sum=total(outside,'after')
    gap=[a+b for a,b in zip(company_diff,outside_sum)]
    assert gap==f['totals']['after_all_minus_amis']
    summary={'kept_old':total(list(kept.values()),'after'),'excluded_amis':[sum(d['source_sums'][c] for d in excluded.values()) for c in ['goods','vat','payment']],'company_invoice_money':[sum(g['invoice_money'][i] for g in groups) for i in range(3)],'company_minus_amis':company_diff,'outside_list':outside_sum,'gap':gap,'components':len(groups),'nonzero_components':sum(any(g['difference']) for g in groups),'outside_chains':outside}
    (WORK/'prepared.json').write_text(json.dumps({'summary':summary,'groups':groups,'outside_chains':outside,'unresolved_exclusions':p['unresolved_exclusions']},ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps({'summary':summary,'nonzero_groups':[{k:g[k] for k in ['roots','codes','missing_codes','invoice_money','amis_money','difference','buyers','exclusion_rows']} for g in groups if any(g['difference'])]},ensure_ascii=False))
if __name__=='__main__':main()
