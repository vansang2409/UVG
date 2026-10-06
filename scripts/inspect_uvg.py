from pathlib import Path
import os
from openpyxl import load_workbook
from collections import Counter, defaultdict
import json

base=Path(os.environ.get('UVG_DATA_DIR', str(Path(__file__).resolve().parents[1] / 'data'))) / 'Du lieu Meinvoice'
data={}
for p in base.glob('*.xls'):
    w=load_workbook(p.open('rb'),read_only=True,data_only=True)
    data[p.name]=list(w.worksheets[0].values)
def valid(rows,start):
    return [(i,r) for i,r in enumerate(rows,1) if i>=start and isinstance(r[0],(int,float))]
def num(v): return v if isinstance(v,(int,float)) else 0
used={}; details={}; edges=[]
for name,rows in data.items():
    if name.startswith('Bang_ke'):
        rs=valid(rows,6); keys=[(r[1],r[2]) for _,r in rs]
        print('USED',name,'rows',len(rs),'unique',len(set(keys)),'firstlast',keys[:1],keys[-1:],'statuses',dict(Counter(r[20] for _,r in rs)),'order_nonempty',sum(bool(r[5]) for _,r in rs))
        for i,r in rs:
            k=(r[1],r[2])
            if k in used and used[k][2]!=r[3:]: print('CONFLICT',k)
            used[k]=(name,i,r[3:])
    if name.startswith('Bảng kê chi tiết'):
        rs=valid(rows,6); groups=defaultdict(list)
        for i,r in rs: groups[(r[1],r[2])].append((i,r)); details[(r[1],r[2])]=groups[(r[1],r[2])]
        print('DETAIL',name,'lines',len(rs),'invoices',len(groups),'statuses',dict(Counter(g[0][1][23] for g in groups.values())),'dates',min(r[3] for _,r in rs),max(r[3] for _,r in rs),'max_lines',max(map(len,groups.values())))
        big=sorted(groups.items(),key=lambda kv:len(kv[1]),reverse=True)[:3]
        for k,g in big: print('BIG',k,'lines',len(g),'rows',g[0][0],g[-1][0],'total_lines',sum(num(r[18]) for _,r in g),'buyer',g[0][1][4:8])
    if name.startswith('DS'):
        rs=valid(rows,7); typ='adjust' if 'điều chỉnh' in name else 'replace'
        print('LINKS',name,len(rs),'old',len(set((r[1],r[3]) for _,r in rs)),'new',len(set((r[12],r[14]) for _,r in rs)))
        for i,r in rs: edges.append((typ,(r[1],r[3]),(r[12],r[14]),num(r[11]),num(r[22]),name,i))
print('UNION_USED',len(used),'UNION_DETAIL',len(details),'detail_missing_used',len(set(details)-set(used)),'used_missing_detail',len(set(used)-set(details)))
print('UNION_STATUS',dict(Counter(v[2][17] for v in used.values())))
for name,rows in data.items():
 if name.startswith('Bang_ke'):
  ks=set((r[1],r[2]) for _,r in valid(rows,6))
  for other,rr in data.items():
   if other.startswith('Bang_ke') and name<other: print('OVERLAP',name,other,len(ks & set((r[1],r[2]) for _,r in valid(rr,6))))
byold=defaultdict(list)
for e in edges: byold[e[1]].append(e)
print('REPEATED_ADJUST',[(k,len(es)) for k,es in byold.items() if sum(e[0]=='adjust' for e in es)>1][:10])
overs=[]
for k,es in byold.items():
 aa=[e for e in es if e[0]=='adjust']
 if aa and aa[0][3]+sum(e[4] for e in aa)<0: overs.append((k,aa[0][3],sum(e[4] for e in aa),aa[0][3]+sum(e[4] for e in aa),[e[6] for e in aa]))
print('NEGATIVE_AFTER_ADJUST',len(overs),overs[:8])
print('REPLACEMENT_CHAINS',[(e[1],e[2],[(x[0],x[2]) for x in byold[e[2]]]) for e in edges if e[0]=='replace' and e[2] in byold][:10])
draft=data['Mau HD Dieu chinh.xls']; dg=defaultdict(list)
for i,r in enumerate(draft,1):
 if i>=10 and r[0] is not None and r[9] and r[10]: dg[str(r[0])].append((i,r))
print('DRAFT_COUNT',len(dg))
for st,g in dg.items():
 r=g[0][1]; k=(r[9],r[10]); print('DRAFT',st,'rows',g[0][0],g[-1][0],'old',k,'amount',r[19:23],'sum_lines',sum(num(x[18]) for _,x in g),'history',[(e[2],e[3],e[4]) for e in byold[k]])
