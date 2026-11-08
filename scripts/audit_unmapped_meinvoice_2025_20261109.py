"""Audit whole invoice components, including every ancestor/descendant, before import."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal,ROUND_HALF_UP
import json,sys,re,hashlib,gzip
from prepare_invoice_reports import rows
from import_2025_buyer_fields import is_organization
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def n(v):return Decimal(str(v)) if v is not None and str(v).strip() else None
def main():
    src=ROOT/'scripts/classify_2025_20261105.py'
    prefix=src.read_text(encoding='utf-8').split("amis=read(SRC/'plan25AMIS.json')",1)[0]
    prefix=prefix.replace("OUT=ROOT/'outputs/phan-loai-2025-20261105'",'OUT=Path('+repr(str(OUT))+')')
    ns={'__file__':str(src),'__name__':'chain_reader'};exec(compile(prefix,str(src),'exec'),ns)
    ledger,children,parents,replacement,roots=map(ns.get,['ledger','children','parents','replacement','roots'])
    p=read(OUT/'prepared.json');a=read(OUT/'authoring.json');known=defaultdict(set)
    for d in p['decisions']:
        for k in d['roots']:known[k].add('AMIS2025:'+d['doc'])
    # Unverified snapshot links are protected, not grounds to zero an alleged orphan.
    for c,d in p['docs'].items():
        for k in d.get('unverified_snapshot_roots',[]):known[k].add('Liên kết AMIS chưa xác minh:'+c)
    orders={str(d['order']) for d in p['docs'].values() if d.get('order')}
    codes=set(p['docs']);company={r['key'] for r in p['company_rows'] if r['key']}
    f24=ROOT/'data/Du lieu Amis/So_chi_tiet_ban_hang_2024.xlsx'
    for rn,r in rows(f24):
        if rn<=4:continue
        c=str(r.get('C') or '').strip()
        if not re.match(r'(BHH|BHDV)',c):continue
        codes.add(c);k=ns['kval'](r.get('D'),r.get('F'))
        if k:known[k].add('AMIS2024:'+c)
        orders.update(re.findall(r'(?<!\d)\d{12,22}(?!\d)',str(r.get('G') or '')))
    for order in orders|codes:
        for k in ns['raw_exact'].get(order,set()):known[k].add('Mã nguồn lịch sử:'+order)
    for order,keys in ns['raw_exact'].items():
        for c in re.findall(r'(?:BHH[A-Z]*|BHDV)\d+',order):
            if c in codes:
                for k in keys:known[k].add('Số CT nguồn lịch sử:'+c)
    old=read(ROOT/'scripts/data/post-issue-20261029/audit.json')
    excluded=old.get('accountant_exclusions',[])
    graph=defaultdict(set)
    for x,ys in children.items():
        for y in ys:graph[x].add(y);graph[y].add(x)
    allkeys=set(ledger)|set(graph);seen=set();components=[]
    order_prefix={o[:15] for o in orders if re.fullmatch(r'\d{18}',o)}
    for seed in sorted(allkeys):
        if seed in seen:continue
        todo=[seed];component=set()
        while todo:
            k=todo.pop()
            if k in component:continue
            component.add(k);seen.add(k);todo.extend(graph[k]-component)
        targets=[k for k in component if str(ledger.get(k,{}).get('date','')).startswith('2025-')]
        if not targets:continue
        rs=sorted(k for k in component if not parents[k]);proof=set()
        for k in component:
            proof.update(known.get(k,set()))
            order=str(ledger.get(k,{}).get('order') or '').strip()
            if order and (order in orders or order in codes):proof.add('Mã đơn/chứng từ nguồn:'+order)
            for c in re.findall(r'(?:BHH[A-Z]*|BHDV)\d+',order):
                if c in codes:proof.add('Số CT nguồn:'+c)
        # Same first 15 digits with suffix000: protect possible matches, never zero on uncertain mapping.
        for k in component:
            order=str(ledger.get(k,{}).get('order') or '')
            if re.fullmatch(r'\d{18}',order) and order[:15] in order_prefix:proof.add('Ứng viên mã đơn cùng15số:'+order)
        rec={'root':rs[0] if len(rs)==1 else '; '.join(rs),'roots':rs,'keys':sorted(component),'invoices_2025':targets,'mapping_evidence':sorted(proof),'issues':[],'status':None,'kind':None}
        if component&company:rec['status']='company'
        elif proof:rec['status']='mapped_or_candidate'
        elif len(rs)!=1:rec['status']='review';rec['issues'].append('Không có duy nhất F0 / có chu trình hoặc nhiều gốc')
        elif not str(ledger.get(rs[0],{}).get('date','')).startswith('2025-'):rec['status']='protected_other_year'
        else:rec['status']='candidate'
        components.append(rec)
    wanted=set(k for r in components if r['status']=='candidate' for k in r['keys']);ident={};files=defaultdict(set)
    for k in wanted:
        if k in ledger:files[ledger[k]['source']['file']].add(ledger[k]['source']['row'])
    for f,rownums in files.items():
        for rn,r in rows(ROOT/f):
            if rn not in rownums:continue
            k=ns['kval'](r.get('B'),r.get('C'))
            if k not in wanted:continue
            ident[k]={'company':r.get('H'),'buyer':r.get('L'),'tax':r.get('J'),'address':r.get('I'),'cqt':r.get('E'),'source':{'file':f,'row':rn}}
    for rec in components:
        if rec['status']!='candidate':continue
        k=rec['root'];ch=ns['chain'](k);rec.update(balance=ch['balance'],active=ch['active'],excluded=ch['excluded'])
        assert set(ch['keys'])==set(rec['keys'])
        rec['identities']={x:ident.get(x) for x in rec['keys']}
        if any(is_organization((ident.get(x) or {}).get('company')) or re.search(r'\bCO\.?[,]?\s*(LTD|LIMITED)\b|\bCORPORATION\b',str((ident.get(x) or {}).get('company') or ''),re.I) for x in rec['keys']):rec['status']='company';continue
        if any((ident.get(x) or {}).get('tax') for x in rec['keys']):rec['status']='review';rec['issues'].append('Có MST nhưng chưa xác định chắc khách lẻ/công ty');continue
        rec['issues']+=ch['issues']
        if any(v is None for v in ch['balance']):rec['issues'].append('Số dư còn ô trống, không mặc định0')
        if rec['issues']:rec['status']='review';continue
        b=[n(v) for v in ch['balance']]
        if all(v==0 for v in b):rec['status']='already_zero';continue
        if b[0]+b[1]!=b[2]:rec['issues'].append('Tiền hàng+VAT khác thanh toán')
        adjusted=any(parents[x] and all(x not in replacement[q] for q in parents[x]) for x in rec['keys'])
        rec['kind']='adjustment' if adjusted else 'replacement' if len(rec['keys'])==1 else None
        if not rec['kind']:rec['issues'].append('Đã có thay thế nhưng chưa có điều chỉnh; cần xác minh đích xử lý')
        rates=[rate for rate in [0,5,8,10] if abs((b[0]*rate/100).quantize(Decimal(1),rounding=ROUND_HALF_UP)-b[1])<=2]
        if len(rates)!=1:rec['issues'].append('Chưa xác định duy nhất thuế suất')
        else:rec['rate']=rates[0]
        if ident.get(k) and not ident[k].get('buyer') and ns['norm'](ident[k].get('company')).startswith(('khach le','khach hang le')):
            ident[k]['buyer']=ident[k]['company'];ident[k]['buyer_evidence']='Dùng nhãn khách lẻ nguồn vì Người mua hàng nguồn trống'
        if not ident.get(k) or not ident[k].get('buyer'):rec['issues'].append('Thiếu thông tin người mua nguồn')
        if '1C25TUV/00001943' in rec['keys']:rec['issues'].append('Ngoại lệ kế toán đã chốt không xử lý; giữ theo bằng chứng trước')
        rec['status']='review' if rec['issues'] else 'ready'
    result={'source_cutoff':'2026-11-01','components':components,'summary':dict(Counter(r['status'] for r in components)),'ready':dict(Counter(r['kind'] for r in components if r['status']=='ready')),'source_files':[{'file':f,'sha256':hashlib.sha256((ROOT/f).read_bytes()).hexdigest()} for f in sorted(set(files)|{str(f24.relative_to(ROOT))})]}
    scope={k for r in components if r['status'] not in ['mapped_or_candidate','protected_other_year'] for k in r['keys']}
    result['invoice_details']={k:{'record':ledger.get(k),'parents':sorted(parents[k]),'children':sorted(children[k]),'replacement_children':sorted(replacement[k])} for k in sorted(scope)}
    (OUT/'unmapped-meinvoice-audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:result[k] for k in ['source_cutoff','summary','ready']},ensure_ascii=False))
if __name__=='__main__':main()
