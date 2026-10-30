"""Dựng hai tab báo cáo tháng/mã hàng và danh sách cần đối chiếu trong file hiện tại."""
from pathlib import Path
import json,math
ROOT=Path(__file__).resolve().parents[1]
P=ROOT/'scripts/data/monthly-goods-20261028'
d=json.loads((P/'prepared.json').read_text(encoding='utf8'))
MAIN=700000020;PENDING=700000021
def cell(v):
    if v is None:return {}
    if isinstance(v,(int,float)):return {'userEnteredValue':{'numberValue':v}}
    return {'userEnteredValue':{'formulaValue':v} if v.startswith('=') else {'stringValue':v}}
vat=json.loads((ROOT/'scripts/data/monthly-vat-quarters-20261030/prepared.json').read_text(encoding='utf8'))
vat_by_month={r['month']:r for r in vat['monthly_vat']}
combined_headers=['Tháng','DT AMIS','VAT AMIS','DT hóa đơn (dự tính)','VAT hóa đơn (dự tính)','Chênh DT','Chênh VAT']
main=[['HÀNG XUẤT BÁN NĂM 2024 — TỔNG HỢP THEO THÁNG VÀ QUÝ'],
['Đơn vị: đồng, doanh thu chưa VAT. Bán/trả theo ngày AMIS; HĐ cũ theo tháng hóa đơn gốc.'],
['1. DOANH THU VÀ VAT THEO THÁNG'],combined_headers]
for r in d['monthly']:
    rn=len(main)+1
    v=vat_by_month[r['month']]
    main.append([r['month'],r['rev_amis'],v['amis'],r['rev_invoice'],v['invoice'],f'=D{rn}-B{rn}',f'=E{rn}-C{rn}'])
assert len(main)==16
main.extend([['CẢ NĂM']+[f'=SUM({c}5:{c}16)' for c in 'BCDEFG'],[],['2. DOANH THU VÀ VAT THEO QUÝ'],['Quý']+combined_headers[1:]])
for q in range(1,5):
    a=5+(q-1)*3;b=a+2;rn=len(main)+1
    main.append([f'Quý {q}']+[f'=SUM({c}{a}:{c}{b})' for c in 'BCDE']+[f'=D{rn}-B{rn}',f'=E{rn}-C{rn}'])
main.extend([['CẢ NĂM']+[f'=SUM({c}21:{c}24)' for c in 'BCDEFG'],[],['3. GIÁ VỐN THEO THÁNG'],['Tháng','Sổ AMIS','Hóa đơn: phần đã có giá vốn','Còn cần bổ sung']])
for r in d['monthly']:
    refs=[x['invoice'].split('/')[-1].lstrip('0') for x in d['cost_pending'] if x['month']==r['month']]
    main.append([r['month'],r['cost_amis'],r['cost_invoice_known'],'Giá vốn HĐ '+', '.join(refs) if refs else None])
assert len(main)==40
main.extend([['CẢ NĂM','=SUM(B29:B40)','=SUM(C29:C40)','Còn thiếu giá vốn HĐ 63 và 111.'],
['Giá vốn HĐ ở cột C chỉ là phần đã đối chiếu; chưa bao gồm hai hóa đơn đang thiếu nguồn.'],[],
])
main[42]=['4. GIÁ VỐN THEO QUÝ']
main.append(['Quý','Sổ AMIS','Hóa đơn: phần đã có giá vốn','Còn cần bổ sung'])
for q in range(1,5):
    a=29+(q-1)*3;b=a+2
    main.append([f'Quý {q}',f'=SUM(B{a}:B{b})',f'=SUM(C{a}:C{b})','Thiếu giá vốn HĐ 63 và 111' if q==1 else None])
main.extend([['CẢ NĂM','=SUM(B45:B48)','=SUM(C45:C48)','Còn thiếu giá vốn HĐ 63 và 111'],[],[],
['CHI TIẾT MÃ HÀNG VÀ SỐ LƯỢNG — mở dấu + bên trái khi cần xem'],
['Tháng','Mã hàng','Tên hàng','ĐVT','SL AMIS','SL HĐ dự tính','Chênh SL','DT AMIS','DT HĐ dự tính','Chênh DT','GV AMIS','GV HĐ đã ghép','Chênh GV đã ghép','Tình trạng']])
assert len(main)==53
first_detail=len(main)+1
heights=[]
for r in d['stats']:
    rn=len(main)+1;missing=not r['cost_invoice_complete']
    status='Đã ghép nguồn AMIS'
    if r['code'].startswith('HD_'):status='Chưa ghép mã; giá vốn chưa phân bổ vào dòng HĐ này'
    elif r['code'].startswith('DIEU_'):status='Điều chỉnh tiền chưa phân bổ theo mã hàng; không phát sinh SL'
    if r['notes']:status+='; '+'; '.join(r['notes'])
    main.append([r['month'],r['code'],r['name'],r['unit'],r['qty_amis'],r['qty_invoice'],f'=F{rn}-E{rn}',r['rev_amis'],r['rev_invoice'],f'=I{rn}-H{rn}',r['cost_amis'],None if missing and not r['cost_invoice'] else r['cost_invoice'],None if missing else f'=L{rn}-K{rn}',status])
    heights.append({'row':rn,'height':min(110,max(42,math.ceil(len(r['name'])/38)*13+10))})
last_detail=len(main)
main.append(['TỔNG TIỀN',None,None,None,None,None,None]+[f'=SUM({c}{first_detail}:{c}{last_detail})' for c in 'HIJKL']+[f'=IF(COUNTA($D$29:$D$40)>0;"";L{last_detail+1}-K{last_detail+1})','Không cộng chung số lượng khác ĐVT; GV còn thiếu ở HĐ 111/63.'])
assert len(main[-1])==14
pending=[['HÀNG HÓA 2024 — NHỮNG VIỆC CẦN BỔ SUNG'],
['Đọc 5 việc chính dưới đây. Chi tiết từng dòng được thu gọn bên dưới.'],
['Việc cần làm','Số trường hợp','Liên quan','Cần bổ sung'],
['Giá vốn HĐ 63',1,'Tháng 2/2024','BHH hoặc phiếu xuất kho của hàng tặng KA100.'],
['Giá vốn HĐ 111',1,'Tháng 3/2024','Giá vốn đơn BHH2023005090, là đơn bán ghi nhận năm 2023.'],
['Xác định mã hàng / tách combo',d['counts']['unmapped_item_rows'],'Hóa đơn cũ giữ lại','Xác nhận mã hàng và thành phần combo; chưa tự chia giá hoặc số lượng.'],
['Phân bổ tiền điều chỉnh',d['counts']['unallocated_money_rows'],'Hóa đơn chỉ ghi tổng tiền','Xác nhận khoản điều chỉnh thuộc mã hàng nào.'],
['Đơn vị tính khác nhau',d['counts']['unit_mismatch_rows'],'Hóa đơn và sổ AMIS','Xác nhận cách quy đổi; hiện giữ nguyên ĐVT nguồn.'],[],
['CHI TIẾT TỪNG DÒNG — mở dấu + bên trái khi cần xem'],
['Cần kiểm tra','Tháng','HĐ / chứng từ','BHH liên quan','Tên hàng / khoản cần kiểm tra','Số lượng','Tiền hàng liên quan','Cần bổ sung / làm rõ','Nguồn']]
ordered=sorted(d['pending'],key=lambda r:({'Thiếu giá vốn':0,'Mã hàng/Combo':1,'ĐVT khác AMIS':2,'Phân bổ tiền':3}[r['kind']],r['month'],r['invoice']))
for r in ordered:pending.append([r['kind'],r['month'],r['invoice'],'\n'.join(r['bhh']),r['item'],r['qty'],r['revenue'],r['reason'],r['source']])
for label,sid,values,cols in [('main',MAIN,main,14),('pending',PENDING,pending,9)]:
    for i in range(0,len(values),40):
        rs=[{'values':[cell(r[j] if j<len(r) else None) for j in range(cols)]} for r in values[i:i+40]]
        req={'updateCells':{'start':{'sheetId':sid,'rowIndex':i,'columnIndex':0},'rows':rs,'fields':'userEnteredValue'}}
        (P/f'{label}-write-{i//40:02}.json').write_text(json.dumps(req,ensure_ascii=False,separators=(',',':')),encoding='utf8')
meta={'main_rows':len(main),'stats_first':first_detail,'stats_last':last_detail,'stats_count':len(d['stats']),'stats_total_row':last_detail+1,'main_chunks':math.ceil(len(main)/40),'pending_rows':len(pending),'pending_first':12,'pending_chunks':math.ceil(len(pending)/40),'pending_count':len(ordered),'heights':heights}
(P/'layout.json').write_text(json.dumps(meta,ensure_ascii=False),encoding='utf8')
print(json.dumps({k:v for k,v in meta.items() if k!='heights'}))
