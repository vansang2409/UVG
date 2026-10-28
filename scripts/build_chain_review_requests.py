"""Chuẩn bị yêu cầu ghi tab rà soát, không thay đổi các tab import."""
from pathlib import Path
import json
P=Path(__file__).resolve().parent/'data/revenue-chain-audit-20261028'
d=json.loads((P/'review-sheet.json').read_text(encoding='utf8'))
SID=700000010
def cell(v):
    if v is None:return {}
    if isinstance(v,(int,float)):return {'userEnteredValue':{'numberValue':v}}
    return {'userEnteredValue':{'stringValue':v}}
last=14+len(d['rows'])
top=[
    ['BÁO CÁO DOANH THU 2024 — TRƯỚC VÀ SAU XỬ LÝ'],
    ['Đơn vị: đồng, chưa VAT. Đã tính lại theo file vừa pull (9cb2cca). Số sau là dự tính khi hoàn tất các bước.'],
    ['Các bước xử lý','Số tiền (đ)','Giải thích'],
    ['1. Hóa đơn bán đang có',d['known_before'],'Hóa đơn gốc năm 2024, đã cộng các điều chỉnh và lấy bản thay thế cuối.'],
    ['2. Điều chỉnh về 0',d['adjustment_zero'],'Đã gồm phần giảm của 4 cặp 3/508, 498/2140, 574/2189, 478/2578.'],
    ['3. Thay thế bằng 0',d['replacement_zero'],'6 HĐ: 2601, 2610, 2602, 2603, 2604, 2617; KT xử lý riêng.'],
    ['4. Hóa đơn cũ còn giữ lại',d['retained'],'47.402.157.888 − 46.407.850.480 − 371.344.277 = 622.963.131đ.'],
    ['5. Cộng hóa đơn bán mới',d['new_sales'],'Theo AMIS, sau khi loại các đơn trong danh sách bạn gửi.'],
    ['6. Trừ hàng khách trả',d['new_returns'],'163 phiếu trả hàng, lập hóa đơn điều chỉnh riêng.'],
    ['SAU XỬ LÝ — DỰ TÍNH',d['known_after'],'622.963.131 + 49.533.443.555 − 225.865.926đ.'],
    ['Doanh thu theo sổ AMIS',d['amis_net'],'Doanh số bán 50.154.037.726đ − hàng trả lại 225.865.926đ.'],
    ['CHÊNH CẦN ĐỐI CHIẾU',d['gap'],'Số dự tính sau xử lý đang cao hơn sổ AMIS.'],
    ['Chênh còn lại: khác năm và tiền hàng',d['gap'],'HĐ 111 liên quan đơn năm 2023 (+3.888.889đ); các nhóm đã ghép lệch ròng −1.519.929đ.'],
    d['headers'],
]
out=[]
for r in top:out.append({'values':[cell(v) for v in r]})
for i,formula in {3:f'=SUM(G15:G{last})',4:f'=SUM(J15:J{last})',5:f'=SUM(K15:K{last})',6:'=SUM(B4:B6)',9:'=SUM(B7:B9)',11:'=B10-B11',12:'=B12'}.items():
    out[i]['values'][1]={'userEnteredValue':{'formulaValue':formula}}
for idx,r in enumerate(d['rows'],15):
    cells=[cell(v) for v in r]
    cells[11]={'userEnteredValue':{'formulaValue':f'=IF(N{idx}="Cần kiểm tra";"";SUM(G{idx};J{idx};K{idx}))'}}
    out.append({'values':cells})
for i in range(0,len(out),40):
    req={'updateCells':{'start':{'sheetId':SID,'rowIndex':i,'columnIndex':0},'rows':out[i:i+40],'fields':'userEnteredValue'}}
    (P/f'write-{i//40:02}.json').write_text(json.dumps(req,ensure_ascii=False,separators=(',',':')),encoding='utf8')
print(json.dumps({'chunks':(len(out)+39)//40,'rows':len(out),'sheet_id':SID}))
