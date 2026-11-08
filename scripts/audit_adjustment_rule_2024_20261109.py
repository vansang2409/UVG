"""Read all originals and their details for the 18 chains; compare the 2024 reversal rule."""
from pathlib import Path
from collections import defaultdict
import json,sys,hashlib
from prepare_invoice_reports import rows,invoice_key
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def key(s,n):
    k=invoice_key(s,n);return '/'.join(k) if k else None
def main():
    a=read(OUT/'authoring.json');p=read(OUT/'prepared.json');u=read(OUT/'unmapped-meinvoice-audit.json');cache=read(OUT/'source-index.json')['ledger'];chains=[];wanted={};files=defaultdict(set)
    for old in a['zero_audit']+a['unmapped_zero_audit']:
        if old['kind']!='adjustment':continue
        root=old['root'];ch=p['chains'][root] if root in p['chains'] else next(r for r in u['components'] if r['root']==root)
        chains.append({'root':root,'old_sequence':old['sequence'],'keys':ch['keys'],'active':ch['active'],'before':old['before'],'docs':old.get('docs',[]),'origin':'AMIS' if old in a['zero_audit'] else 'Chưa nối AMIS'})
        for k in ch['keys']:
            r=(u['invoice_details'].get(k) or {}).get('record') or cache[k];wanted[k]=r;files[r['source']['file']].add(r['source']['row'])
    headers={};manifest=[]
    for f,ns in files.items():
        path=ROOT/f;manifest.append({'file':f,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
        for rn,r in rows(path):
            if rn not in ns:continue
            k=key(r.get('B'),r.get('C'))
            if k in wanted:headers[k]={'source':{'file':f,'row':rn},'values':r}
    detail=defaultdict(list)
    paths=[ROOT/'data/Du lieu Meinvoice'/name for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']]
    paths+=list((ROOT/'data/Meinv_SAU_DIEU_CHINH_THAY_THE_20261101/CHI_TIET').glob('*.xls'))
    for path in paths:
        mt='MTT' in path.name;latest='CHI_TIET' in str(path);selected=defaultdict(list)
        manifest.append({'file':str(path.relative_to(ROOT)),'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
        labels={}
        for rn,r in rows(path):
            if latest and rn==5:labels={v:k for k,v in r.items() if v}
            if rn<(6 if latest else 7):continue
            get=lambda name:r.get(labels.get(name))
            k=key(get('Ký hiệu'),get('Số hóa đơn')) if latest else key(r.get('F' if mt else 'D'),r.get('B'))
            if k not in wanted:continue
            fields={'name':get('Tên hàng'),'unit':get('ĐVT'),'qty':get('Số lượng'),'price':get('Đơn giá'),'goods':get('Thành tiền'),'rate':get('Thuế suất (%)'),'tax':get('Tiền thuế GTGT'),'pay':get('Tổng tiền TT'),'discount':get('Tiền CK'),'kind':get('Tính chất HHDV'),'promotion':get('Hàng KM')} if latest else {'name':r.get('S' if mt else 'L'),'unit':r.get('T' if mt else 'M'),'qty':r.get('U' if mt else 'N'),'price':r.get('V' if mt else 'O'),'goods':r.get('W' if mt else 'P'),'rate':r.get('Z' if mt else 'S'),'tax':r.get('AA' if mt else 'T'),'pay':r.get('AB' if mt else 'U'),'discount':r.get('Y' if mt else 'R'),'kind':None,'promotion':r.get('AD') if mt else None}
            selected[k].append({'source':{'file':str(path.relative_to(ROOT)),'row':rn},'values':r,'mtt':mt,'latest':latest,'fields':fields})
        for k,ls in selected.items():
            if k not in detail or latest:detail[k]=ls
    result={'chains':chains,'headers':headers,'detail':dict(detail),'ledger':wanted,'manifest':manifest,'rules_source':['scripts/data/check-review/add-adjusted-import.mjs','scripts/data/check-review/verify-adjusted-import.py']}
    (OUT/'adjustment-rule-2024-audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    brief=[]
    for c in chains:
        root=c['root'];ls=detail.get(root,[]);brief.append({'root':root,'nodes':len(c['keys']),'lines':len(ls),'missing_nodes':[k for k in c['keys'] if k not in headers or k not in detail],'source_names':[headers[root]['values'].get(k) for k in ['H','L']],'items':[{'name':l['values'].get('S' if l['mtt'] else 'L'),'qty':l['values'].get('U' if l['mtt'] else 'N'),'price':l['values'].get('V' if l['mtt'] else 'O'),'goods':l['values'].get('W' if l['mtt'] else 'P'),'discount':l['values'].get('Y' if l['mtt'] else 'R'),'rate':l['values'].get('Z' if l['mtt'] else 'S')} for l in ls]})
    print(json.dumps({'chains':len(chains),'old_nodes':len(wanted),'review':brief},ensure_ascii=False))
if __name__=='__main__':main()
