"""Lập bảng rà soát hóa đơn gốc 2024 đã điều chỉnh, ngoài nhóm ngoại lệ."""
from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime
from decimal import Decimal
import json, hashlib
from openpyxl import load_workbook
from filter_invoice_exceptions import txt, is_company, tax_shaped, normal

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/hoa-don-dieu-chinh-2024'

def main():
    for e in json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf-8')):
        assert hashlib.sha256((ROOT/e['path']).read_bytes()).hexdigest()==e['sha256']
    headers={}; observations=defaultdict(list); dates={}; children=defaultdict(set); parents=defaultdict(set); edges=[]
    def observe(k,name,tax,date):
        observations[k].append((txt(name),txt(tax)))
        dates.setdefault(k,date.strftime('%d/%m/%Y') if isinstance(date,datetime) else txt(date))
    for p in sorted((ROOT/'data/Du lieu Meinvoice').glob('*.xls')):
        if p.name.startswith('Mau'):continue
        with p.open('rb') as stream:
            wb=load_workbook(stream,read_only=True,data_only=True)
            for i,r in enumerate(wb.worksheets[0].values,1):
                if not isinstance(r[0],(int,float)):continue
                ref=f'{p.relative_to(ROOT).as_posix()} | dòng {i}'
                if p.name.startswith('Bang_ke'):
                    k=(txt(r[1]),txt(r[2]));observe(k,r[7] or r[11],r[9],r[3])
                    headers[k]={'date':txt(r[3]),'name':txt(r[7] or r[11]),'tax':txt(r[9]),'status':txt(r[20]),
                                'money':[v if isinstance(v,(int,float)) else None for v in r[15:18]],'source':ref}
                elif p.name.startswith('Bảng kê chi tiết'):
                    observe((txt(r[1]),txt(r[2])),r[4],r[6],r[3])
                elif p.name.startswith('DS'):
                    old,new=(txt(r[1]),txt(r[3])),(txt(r[12]),txt(r[14]))
                    observe(old,r[5],r[6],r[2]);observe(new,r[16],r[17],r[13])
                    kind='Điều chỉnh' if 'điều chỉnh' in p.name else 'Thay thế'
                    children[old].add(new);parents[new].add(old)
                    edges.append((old,new,kind,ref))
            wb.close()
    exceptional={k for k,v in observations.items() if any(is_company(name) or tax_shaped(tax) for name,tax in v)}
    seeds={k for k in observations if not parents[k] and datetime.strptime(dates[k],'%d/%m/%Y').year==2024
           and (any(e[0]==k and e[2]=='Điều chỉnh' for e in edges) or headers.get(k,{}).get('status')=='Hóa đơn đã bị điều chỉnh')}
    groups=[]; exclusions=[]; issues=Counter(); detail=[]; used_edges=[]
    for root in sorted(seeds):
        nodes=set();queue=[root]
        while queue:
            k=queue.pop()
            if k in nodes:continue
            nodes.add(k);queue.extend(children[k])
        linked=[e for e in edges if e[0] in nodes and e[1] in nodes]
        reason=''
        if nodes & exceptional:reason='Ngoại lệ: chuỗi có công ty hoặc định danh dạng MST'
        elif any(e[2]=='Thay thế' for e in linked) or any('thay the' in normal(headers.get(k,{}).get('status','')) for k in nodes):reason='Chuỗi có hóa đơn thay thế – để ngoài phạm vi lần này'
        if reason:
            exclusions.append([*root,reason]);continue
        pending=set(nodes);order=[];flags=[]
        while pending:
            available=[k for k in pending if not(parents[k]&pending)]
            if not available:
                flags.append('Vòng lặp liên kết');available=list(pending)
            chosen=min(available,key=lambda k:(datetime.strptime(dates[k],'%d/%m/%Y'),k))
            order.append(chosen);pending.remove(chosen)
        if not linked:flags.append('Trạng thái bị điều chỉnh nhưng thiếu hóa đơn điều chỉnh liên kết')
        if any(len(parents[k])>1 for k in nodes):flags.append('Một hóa đơn sửa tham chiếu nhiều nguồn; chưa phân bổ được')
        rows=[];active=[]
        for k in order:
            h=headers.get(k);name,tax=observations[k][0]
            money=h['money'] if h else [None,None,None]
            if not h:flags.append(f'{k[0]}/{k[1]} thiếu bảng tổng')
            if any(v is None for v in money):flags.append(f'{k[0]}/{k[1]} thiếu tiền')
            if all(v is not None for v in money) and sum(Decimal(str(v)) for v in money[:2])!=Decimal(str(money[2])):
                flags.append(f'{k[0]}/{k[1]} tiền hàng + thuế lệch thanh toán')
            cancelled=h and 'huy' in normal(h['status'])
            if cancelled:flags.append(f'{k[0]}/{k[1]} đã hủy; loại tiền khỏi cộng dồn')
            else:active.append(money)
            refs=sorted(set(f'{e[2]} cho {e[0][0]}/{e[0][1]}' for e in linked if e[1]==k))
            kind='Hóa đơn gốc' if k==root else '; '.join(refs)
            rows.append([*k,dates[k],kind,*money])
            detail.append([*root,*k,dates[k],kind,h['status'] if h else 'Không có trong bảng tổng',*money,
                           h['source'] if h else '; '.join(e[3] for e in linked if e[1]==k)])
        flags=list(dict.fromkeys(flags))
        sums=[float(sum(Decimal(str(m[j])) for m in active)) if active and all(m[j] is not None for m in active) else None for j in range(3)]
        # Quy ước riêng cho hóa đơn 574 đã được người dùng xác nhận; không sửa ô nguồn.
        if root==('1C24TUV','00000574') and headers.get(root,{}).get('money',[None,None,None])[2] is None:
            remaining=[headers[k]['money'][2] for k in order if k!=root and k in headers]
            if len(remaining)==len(order)-1 and all(v is not None for v in remaining):
                sums[2]=float(sum(Decimal(str(v)) for v in remaining))
                flags=[f for f in flags if f!='1C24TUV/00000574 thiếu tiền']
                flags.append('Tạm tính thanh toán gốc = 0 theo xác nhận người dùng; ô nguồn giữ trống.')
        for flag in flags:issues['Thiếu nguồn/tiền' if 'thiếu' in flag else 'Lệch số học' if 'lệch' in flag else 'Khác']+=1
        groups.append({'root':list(root),'name':observations[root][0][0],'tax':observations[root][0][1],
                       'invoices':rows,'sums':sums,'notes':'; '.join(flags)})
        used_edges.extend([[*e[0],*e[1],e[2],e[3]] for e in linked])
    summary={'roots':len(groups),'invoices':len(detail),'max_blocks':max(len(g['invoices']) for g in groups),
             'excluded':dict(Counter(r[2] for r in exclusions)), 'roots_with_notes':sum(bool(g['notes']) for g in groups),
             'max_date':max(datetime.strptime(r[4],'%d/%m/%Y') for r in detail).strftime('%d/%m/%Y')}
    OUT.mkdir(parents=True,exist_ok=True)
    helper_dir=ROOT/'scripts/data'
    helper_dir.mkdir(parents=True,exist_ok=True)
    (helper_dir/'adjusted_2024_review.json').write_text(json.dumps({'groups':groups,'detail':detail,'edges':used_edges,'excluded':exclusions,'summary':summary},ensure_ascii=False),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=True))

if __name__=='__main__':main()
