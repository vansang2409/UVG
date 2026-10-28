"""Prepare the 13 authorized supplements. Read Excel only; preserve input snapshots."""
from pathlib import Path
from collections import defaultdict
from datetime import date,datetime
from decimal import Decimal
from zipfile import ZipFile
from lxml import etree as E
import hashlib,io,json,shutil,sys
import openpyxl

HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[2]
sys.path.insert(0,str(ROOT/'scripts'))
from prepare_invoice_reports import rows,amount,invoice_key,invoice_day
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
def encode(v):
    if isinstance(v,(datetime,date)):return v.isoformat()
    if isinstance(v,Decimal):return int(v) if v==v.to_integral_value() else float(v)
    raise TypeError(type(v))
def save(name,obj):
    (HERE/name).write_text(json.dumps(obj,ensure_ascii=False,indent=2,default=encode),encoding='utf8')
prior=json.loads((ROOT/'scripts/data/decision-85-20261019/data.json').read_text(encoding='utf8'))
for p,h in prior['source_hashes'].items():assert sha(ROOT/p)==h,('changed source',p)
targets={tuple(r['key']) for r in prior['additional_records']}
roots={tuple(r['root'].split('/')) for r in prior['additional_records']}
wanted=targets|{('1C25TUV',n) for n in ('00006108','00004203','00004204')}
assert len(targets)==13 and len(roots)==8 and len(wanted)==16

if (HERE/'sources.json').exists():
    src=json.loads((HERE/'sources.json').read_text(encoding='utf8'))
    for p,h in src['hashes'].items():assert sha(ROOT/p)==h,('source changed',p)
else:
    heads={};details=defaultdict(list);hashes={}
    for p in sorted((ROOT/'data/Du lieu Meinvoice/Xuat_lai_2026-10-09').iterdir()):
        if p.suffix not in ('.xls','.xlsx','.zip') or '1791532943641' in p.name or any(t in p.name for t in ('dieu_chinh_xuat','thay_the_xuat')):continue
        hashes[p.relative_to(ROOT).as_posix()]=sha(p)
        if p.suffix=='.zip':
            with ZipFile(p) as z:books=[(n,z.read(n)) for n in z.namelist() if n.endswith(('.xlsx','.xls'))]
        else:books=[(p.name,p.read_bytes())]
        for name,blob in books:
            is_detail='chi_tiet' in name
            for rn,v in rows(io.BytesIO(blob)):
                if amount(v.get('A')) is None:continue
                k=invoice_key(v.get('B'),v.get('C'))
                if k not in wanted:continue
                key='/'.join(k)
                evidence=p.relative_to(ROOT).as_posix()+(' / '+name if p.suffix=='.zip' else '')+f' | dòng {rn}'
                if is_detail:
                    details[key].append({'name':v.get('L'),'unit':v.get('M'),'qty':amount(v.get('N')),'price':amount(v.get('O')),'goods':amount(v.get('P')),'rate':v.get('Q'),'tax':amount(v.get('R')),'payment':amount(v.get('S')),'buyer':v.get('H'),'customer':v.get('E'),'address':v.get('F'),'taxid':v.get('G'),'method':v.get('J'),'nature':v.get('K'),'source':evidence})
                else:
                    assert key not in heads,('duplicate header',key)
                    heads[key]={'date':invoice_day(v.get('D')),'money':[amount(v.get(c)) for c in ('P','Q','R')],'buyer':v.get('L'),'customer':v.get('H'),'address':v.get('I'),'taxid':v.get('J'),'method':v.get('M'),'status':v.get('U'),'source':evidence}
    assert {'/'.join(k) for k in wanted}<=set(heads)
    xmls={}
    for p in (ROOT/'data/Bo sung').rglob('*.xml'):
        key=p.stem.replace('_','/')
        if key not in {'/'.join(k) for k in wanted}:continue
        hashes[p.relative_to(ROOT).as_posix()]=sha(p)
        tree=E.parse(str(p))
        def text(el,name):return el.findtext('.//'+name)
        def num(v):return None if v is None else Decimal(v)
        lines=[]
        for l in tree.findall('.//HHDVu'):
            lines.append({'name':text(l,'THHDVu'),'unit':text(l,'DVTinh'),'qty':num(text(l,'SLuong')),'price':num(text(l,'DGia')),'goods':num(text(l,'ThTien')),'rate':text(l,'TSuat'),'nature':text(l,'TChat')})
        xmls[key]={'money':[num(text(tree,n)) for n in ('TgTCThue','TgTThue','TgTTTBSo')],'lines':lines,'rates':[x.text for x in tree.findall('.//THTTLTSuat/LTSuat/TSuat')],'file':p.relative_to(ROOT).as_posix()}
    src={'headers':heads,'details':dict(details),'xmls':xmls,'hashes':hashes}
    save('sources.json',src)

if '--inspect' in sys.argv:
    for key,h in src['headers'].items():
        ls=src['details'].get(key,[])
        print(json.dumps({'key':key,'date':h['date'],'money':h['money'],'lines':len(ls),'sum_goods':sum(Decimal(str(l['goods'])) for l in ls if l['goods'] is not None),'odd_lines':[l for l in ls if l['price'] in (None,0) or l['qty'] in (None,0) or l['price']<0][:4],'first':ls[:2],'xml':src['xmls'].get(key)},ensure_ascii=False,default=encode))
    sys.exit()

heads,details,xmls=src['headers'],src['details'],src['xmls']
adj_name='HOA_DON_DA_DIEU_CHINH_F0_2024_T6_2025.xlsx'
compact_name='HOA_DON_DA_DIEU_CHINH_IMPORT_2_SHEET.xlsx'
exc_name='NGOAI_LE_F0_2024_T6_2025.xlsx'
list_name='UVG_72_HOA_DON_KHONG_CAN_DIEU_CHINH_20261019.xlsx'
def book_spec(kind,name,dest):
    before=HERE/'before'/name
    assert before.exists() and sha(before)==sha(ROOT/dest),('changed workbook',name)
    return {'kind':kind,'name':name,'destination':dest,'before_sha256':sha(before),'sheets':{}}
adj=book_spec('adjusted',adj_name,'outputs/'+adj_name)
exc=book_spec('exceptions',exc_name,'outputs/'+exc_name)
lst=book_spec('list72',list_name,'outputs/doi-chieu-85-20261019/'+list_name)
compact=book_spec('compact',compact_name,'outputs/'+compact_name)
wa=openpyxl.load_workbook(HERE/'before'/adj_name,data_only=False)
we=openpyxl.load_workbook(HERE/'before'/exc_name,data_only=False)
def sheet_spec(book,s):
    out={'changes':[],'blocks':[],'formulas':[],'header':list(next(s.iter_rows(min_row=2 if s.title in ('Dieu chinh','Ngoai le') else 9 if s.title=='Import dieu chinh' else 1,max_row=2 if s.title in ('Dieu chinh','Ngoai le') else 9 if s.title=='Import dieu chinh' else 1,values_only=True))),
         'header_row':2 if s.title in ('Dieu chinh','Ngoai le') else 9 if s.title=='Import dieu chinh' else 1,
         'widths':{k:v.width for k,v in s.column_dimensions.items()},'tables':{},'row_heights':{}}
    book['sheets'][s.title]=out
    return out
def change(s,cell,value):s['changes'].append({'cell':cell,'value':value})
def block(s,cell_range,values,dates=()):s['blocks'].append({'range':cell_range,'values':values,'date_columns':list(dates)})
def formula(s,cell,text,cache=0):s['formulas'].append({'cell':cell,'formula':text,'cache':cache})

imp=sheet_spec(adj,wa['Import dieu chinh']);main=sheet_spec(adj,wa['Dieu chinh']);source=sheet_spec(adj,wa['Chi tiet nguon']);guide=sheet_spec(adj,wa['Huong dan']);pending=sheet_spec(adj,wa['Cho KT xac nhan'])
old=[r for r in wa['Import dieu chinh'].iter_rows(min_row=10,max_col=24,values_only=True) if r[0] is not None]
assert max(r[0] for r in old)==2032 and len(old)==41790
assert not {tuple(r[9:11]) for r in old}&roots
newrows=[];actions=[]
for record in prior['additional_records']:
    target='/'.join(record['key']);root=record['root'];h=heads[target];rh=heads[root]
    reverse=target!=root
    original_money=xmls[target]['money'] if target in xmls else h['money']
    original_money=list(original_money)
    if target=='1C24TUV/00000574':
        assert original_money==[0,0,None]
        original_money[2]=0 # Earlier user-confirmed convention; leave the source cell blank.
    assert all(isinstance(x,(int,float)) for x in original_money),(target,original_money)
    money=[-x for x in original_money]
    assert money[0]+money[1]==money[2],(target,money)
    seq=2033+len(actions);start=41800+len(newrows)
    original_lines=details.get(target,[])
    if reverse:
        assert not original_lines or all(l['qty'] in (None,0) and l['price'] in (None,0) and l['goods'] in (None,0) for l in original_lines)
        lines=[{'name':f'Đảo khoản điều chỉnh HĐ {target} ngày {datetime.fromisoformat(str(h["date"])[:10]).strftime("%d/%m/%Y")}',
                'unit':None,'qty':None,'price':None,'goods':money[0],'rate':'8%'}]
        mode='Đảo khoản điều chỉnh cũ; chỉ đảo thành tiền, nguồn không có số lượng/đơn giá'
    else:
        assert original_lines,target
        lines=[]
        for l in original_lines:
            q=l['qty'];p=l['price'];g=l['goods']
            assert q is not None and q>0 and (p is None or p>0),(target,l)
            assert p is None if g is None else p>0,(target,l)
            lines.append({**l,'qty':-q,'price':p,'goods':None if g is None else -g})
        mode='Giảm số lượng gốc, giữ đơn giá dương' if money[0]!=0 else 'Chỉ giảm số lượng gốc; nguồn không có đơn giá/thành tiền'
    assert sum(Decimal(str(l['goods'])) for l in lines if l['goods'] is not None)==Decimal(str(money[0])),target
    rl=details[root][0]
    buyer=rl.get('buyer') or rh['buyer'] or rh['customer']
    address=rl.get('address') or rh['address'];taxid=rl.get('taxid') or rh['taxid'];method=rl.get('method') or rh['method']
    rootdate=datetime.fromisoformat(str(rh['date'])[:10]).strftime('%d/%m/%Y')
    direction='tăng' if money[2]>0 else 'giảm'
    reason=f'Điều chỉnh {direction} cho hóa đơn ký hiệu {root.split("/")[0]}, số {root.split("/")[1]}, ngày {rootdate} do xuất sai nội dung'
    for i,l in enumerate(lines):
        assert l['rate']=='8%'
        row=[seq,'2026-10-19',buyer,address,taxid,buyer,None,method,None,*root.split('/'),rootdate,None,reason,l['name'],l['unit'],l['qty'],l['price'],l['goods'],money[0] if i==0 else None,8 if i==0 else None,money[1] if i==0 else None,money[2] if i==0 else None,None]
        assert len(row)==24 and row[7]=='TM/CK' and row[6] is None
        newrows.append(row)
    actions.append({'target':target,'root':root,'date':str(h['date'])[:10],'seq':seq,'start':start,'end':start+len(lines)-1,'money':money,'source_money':original_money,'mode':mode,'lines':lines})
assert len(actions)==13 and len(newrows)==206
adj.update(invoices=2045,roots=1015,rows=41996,new_actions=actions,new_rows=newrows)
block(imp,'A41800:X42005',newrows,[1]);imp['append_template']=41799;imp['append_start']=41800
imp['row_heights']={str(a['start']+i):max(62,14*((len(l['name'])+65)//66)) for a in actions for i,l in enumerate(a['lines'])}
for cell,value in {'A3':'- Phần bổ sung mới dùng ngày 19/10/2026; phần trước giữ nguyên ngày 08/10, 09/10 và 12/10/2026.',
                   'A4':'- 1.015 chuỗi / 2.045 hóa đơn / 41.996 dòng; bổ sung 13 hóa đơn / 206 dòng theo chốt ngày 19/10/2026.',
                   'A7':'- Giữ 3 cặp cùng sản phẩm và 4 hóa đơn chỉ đổi thông tin. 13 hóa đơn mới không có đơn giá âm; ô số lượng/giá vốn trống giữ trống. 8 khoản đảo cũ tô vàng vẫn giữ số đã phát hành.'}.items():change(imp,cell,value)

exrows=[(rn,list(r)) for rn,r in enumerate(we['Ngoai le'].iter_rows(min_row=3,values_only=True),3) if r[2]]
ex_by_root={'/'.join(r[2:4]):(rn,r) for rn,r in exrows}
mr={'/'.join(r[:2]):rn for rn,r in enumerate(wa['Dieu chinh'].iter_rows(min_row=3,values_only=True),3) if r[0]}
new_main=[];main['display_rows']=[];root_checks=[]
for root in dict.fromkeys(a['root'] for a in actions):
    acts=[a for a in actions if a['root']==root]
    if root in mr:
        rn=mr[root];values=[wa['Dieu chinh'].cell(rn,c).value for c in range(1,56)]
        main['display_rows'].append({'row':rn,'values':values[:27]})
    else:
        rn=1050+len(new_main);r=ex_by_root[root][1]
        values=r[2:20]+[None]*6+r[20:23]+[None]*28
        assert len(values)==55
        new_main.append(values);mr[root]=rn
    values[27]='Đã bổ sung theo chốt 19/10/2026; '+('; '.join(f'{a["mode"]}: {a["target"]}, STT {a["seq"]}' for a in acts))+'.'
    values[28:52]=[None]*24
    for i,a in enumerate(acts):values[28+i*6:34+i*6]=[f'{a["mode"]}: {a["target"]} — STT import {a["seq"]}',None,datetime(2026,10,19),*a['money']]
    if root.endswith(('/00001348','/00002377','/00002378')):values[27]+=' Giữ điều chỉnh thông tin tiền 0 theo XML; ô thanh toán bảng nguồn trống vẫn giữ trống.'
    if root=='1C24TUV/00000574':values[27]+=' Thanh toán gốc dùng quy ước 0 đã chốt; nguồn vẫn để trống.'
    if rn<1050:block(main,f'AB{rn}:BC{rn}',[values[27:55]],[3,9,15,21])
    else:new_main[rn-1050]=values
    for col,base,terms in [('BA','Y',['AF','AL','AR','AX']),('BB','Z',['AG','AM','AS','AY'])]:
        formula(main,f'{col}{rn}',f'=SUM({base}{rn},'+','.join(f'{c}{rn}' for c in terms)+')')
    base=f'AA{rn}' if isinstance(values[26],(int,float)) else f'SUM(F{rn},L{rn},R{rn},X{rn})'
    formula(main,f'BC{rn}',f'=SUM({base},AH{rn},AN{rn},AT{rn},AZ{rn})')
    main['row_heights'][str(rn)]=130
    source_money=[]
    # The information-only adjustment is retained in each of these three chains.
    info_map={'1C25TUV/00001348':'1C25TUV/00006108','1C25TUV/00002377':'1C25TUV/00004203','1C25TUV/00002378':'1C25TUV/00004204'}
    relkeys=list(dict.fromkeys([root]+[a['target'] for a in acts if a['target']!=root]+([info_map[root]] if root in info_map else [])))
    for k in relkeys:
        m=list(xmls[k]['money'] if k in xmls else heads[k]['money'])
        if k=='1C24TUV/00000574':m[2]=0
        assert all(x is not None for x in m),(k,m)
        source_money.append(m)
    totals=[sum(m[i] for m in source_money)+sum(a['money'][i] for a in acts) for i in range(3)]
    assert totals==[0,0,0],(root,totals)
    root_checks.append({'root':root,'row':rn,'source_invoices':relkeys,'source_money':source_money,'sequences':[a['seq'] for a in acts],'after_goods_tax_payment':totals})
assert len(new_main)==2
block(main,'A1050:BC1051',new_main,[2,8,14,20,30,36,42,48]);main['append_start']=1050;main['append_template']=1049
main['tables']={'DieuChinhChuoi':'A2:BC1051'}

eroots={'/'.join(k) for k in roots}
excsource=[(rn,list(r)) for rn,r in enumerate(we['Chi tiet nguon'].iter_rows(min_row=2,values_only=True),2) if r[0]]
new_source=[r for _,r in excsource if '/'.join(r[:2]) in {'1C24TUV/00000478','1C25TUV/00000040'}]
assert len(new_source)==4
block(source,'A2115:M2118',new_source,[4]);source['append_start']=2115;source['append_template']=2114;source['tables']={'NguonDieuChinh':'A1:M2118'}
for rn,r in enumerate(wa['Cho KT xac nhan'].iter_rows(min_row=3,values_only=True),3):
    root='/'.join(str(v) for v in r[:2])
    acts=[a for a in actions if a['root']==root]
    if acts:
        change(pending,f'G{rn}','Đã bổ sung import')
        change(pending,f'H{rn}','Bổ sung 19/10/2026, STT '+', '.join(str(a['seq']) for a in acts)+'. '+('Giữ hóa đơn chỉ đổi thông tin, tiền 0; chỉ giảm gốc.' if root in info_map else 'Đảo khoản điều chỉnh cũ và giảm gốc; tổng chuỗi về 0.'))
change(pending,'A1','19 chuỗi đã chốt — đều đã có trong import; 6 chuỗi bổ sung theo quyết định 19/10/2026')
guide_updates={
 'A3':'1.049 F0 lịch sử, 2.117 hóa đơn nguồn riêng biệt. Đã bổ sung 13 hóa đơn trên 8 chuỗi theo chốt 19/10/2026.',
 'A4':'F0 năm 2024: 592; F0 tháng 1–6/2025: 457.',
 'A5':'Giữ dữ liệu nguồn; 3 cặp cùng sản phẩm và 4 hóa đơn chỉ đổi thông tin không lập đảo tiền. 8 chuỗi vừa chuyển từ Ngoai le sang xử lý theo xác nhận mới.',
 'A13':'Import dieu chinh: 1.015 chuỗi / 2.045 hóa đơn / 41.996 dòng. Giữ nguyên 2.032 hóa đơn trước; bổ sung 13 hóa đơn mới.',
 'A14':'Cho KT xac nhan: 19 hồ sơ đã có import. 42 - Can check 11: 8 chuỗi làm lại và 3 chuỗi cùng sản phẩm giữ nguyên.',
 'A15':'Phần mới ngày 19/10/2026: giảm số lượng gốc, giá có sẵn giữ dương; 5 khoản điều chỉnh chỉ có tiền giữ trống số lượng/giá. Gốc574 chỉ giảm số lượng. Giữ 6108/4203/4204 và9186, không đảo tiền các hóa đơn thông tin.'}
for cell,v in guide_updates.items():change(guide,cell,v)

for sn,rows0,start,maxcol,dates,table,ref in [
 ('Ngoai le',exrows,3,24,[4,10,16],'NgoaiLeChuoi','A2:X238'),
 ('Chi tiet nguon',excsource,2,13,[4],'NguonNgoaiLe','A1:M323')]:
    kept=[(rn,r) for rn,r in rows0 if '/'.join(r[2:4] if sn=='Ngoai le' else r[:2]) not in eroots]
    assert len(kept)==(236 if sn=='Ngoai le' else 322)
    ss=sheet_spec(exc,we[sn]);block(ss,f'A{start}:{openpyxl.utils.get_column_letter(maxcol)}{start+len(kept)-1}',[r for _,r in kept],dates)
    ss['replace_data']={'start':start,'end':start+len(kept)-1,'maxcol':maxcol,'row_map':{str(start+i):rn for i,(rn,_) in enumerate(kept)}}
    ss['tables']={table:ref}
eg=sheet_spec(exc,we['Huong dan'])
for cell,v in {
 'A3':'236 chuỗi giữ nguyên / 322 hóa đơn liên quan; 8 chuỗi đã chuyển sang bộ xử lý theo xác nhận ngày 19/10/2026.',
 'A4':'F0 năm 2024: 117; F0 tháng 1–6/2025: 119.',
 'A5':'Giữ toàn chuỗi theo tiêu chí công ty/hộ kinh doanh và ngoại lệ còn hiệu lực. Hai chuỗi MST cá nhân 478/2578,40/662 và6cặp cũ đã được chuyển sang xử lý theo xác nhận mới.',
 'A11':'Theo chốt 19/10: bỏ giữ riêng6cặp cũ; MST8823260350 và8743683287 là cá nhân. 4 hóa đơn chỉ đổi thông tin không đổi tiền được giữ riêng trong danh sách72.',
 'A12':'Bộ Đã điều chỉnh giữ đủ lịch sử 8 chuỗi đã chuyển; bộ import bổ sung13 hóa đơn. Không có hóa đơn thuộc236chuỗi còn giữ nguyên trong import.'}.items():change(eg,cell,v)

lst['sheets']['Tong hop']={'changes':[
 {'cell':'D5','value':'10.442 hóa đơn trước bổ sung + 85 hóa đơn đối chiếu.'},
 {'cell':'A8','value':'Bộ import trước đợt bổ sung'},
 {'cell':'D8','value':'8.410 chưa điều chỉnh + 2.032 đã điều chỉnh trước bổ sung.'},
 {'cell':'A9','value':'Đã bổ sung vào bộ import'},
 {'cell':'D9','value':'Đã bổ sung 13 hóa đơn / 8 chuỗi; bộ đã điều chỉnh nay có 2.045 hóa đơn.'},
 {'cell':'A21','value':'Đã cập nhật bộ import: 10.455 hóa đơn = 8.410 chưa điều chỉnh + 2.045 đã điều chỉnh.'}],
 'blocks':[],'formulas':[],'tables':{},'row_heights':{}}
compact['sheets']={new:adj['sheets'][old] for old,new in [('Import dieu chinh','Import dieu chinh'),('Dieu chinh','F0, F1, F2, ...')]}
compact['author_sheet_map']={'F0, F1, F2, ...':'Dieu chinh','Import dieu chinh':'Import dieu chinh'}
plan={'date':'2026-10-19','files':[adj,compact,exc,lst],'actions':actions,'root_checks':root_checks,'source_hashes':src['hashes'],
      'new_source_rows':new_source,'counts':{'added_invoices':13,'added_rows':206,'adjusted_invoices':2045,'adjusted_rows':41996,'adjusted_roots':1015,'all_invoices':10455,'all_rows':55564,'exceptions_roots':236,'exceptions_invoices':322},
      'retained_information_invoices':['1C25TUV/00006108','1C25TUV/00004203','1C25TUV/00004204','1C25TUV/00009186'],
      'untreated_hashes':{n:sha(ROOT/'outputs'/n) for n in ['HOA_DON_CHUA_DIEU_CHINH_F0_2024_T6_2025.xlsx','HOA_DON_CHUA_DIEU_CHINH_IMPORT_2_SHEET.xlsx']}}
save('plan.json',plan)
wa.close();we.close()
print(json.dumps({'counts':plan['counts'],'sequences':[[a['seq'],a['target'],len(a['lines']),a['money']] for a in actions],'zero_checks':root_checks},ensure_ascii=False,default=encode))
