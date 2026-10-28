"""Đóng gói import lớn: dùng dự phòng đã xác nhận do Artifact export lớn bị treo."""
from pathlib import Path
from datetime import datetime
import json
import sys
import os
import xlsxwriter

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/amis-2024-linked-20261026'
OUT=Path(os.environ.get('UVG_OUTPUT_DIR',ROOT/'outputs/amis-2024-theo-btl-bhh-20261026'))

def main():
    p=json.loads((WORK/'prepared.json').read_text(encoding='utf8'))
    OUT.mkdir(parents=True,exist_ok=True)
    file=OUT/'IMPORT_HOA_DON_MOI_AMIS_2024.xlsx'
    if file.exists() and '--replace' not in sys.argv:raise FileExistsError(file)
    w=xlsxwriter.Workbook(file,{'constant_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    base={'font_name':'Arial','font_size':10,'valign':'vcenter'}
    body=w.add_format(base);wrap=w.add_format({**base,'text_wrap':True});text=w.add_format({**base,'num_format':'@'})
    money=w.add_format({**base,'num_format':'#,##0;[Red](#,##0);0'});decimal=w.add_format({**base,'num_format':'#,##0.00########'})
    date=w.add_format({**base,'num_format':'dd/mm/yyyy'});head=w.add_format({**base,'bold':True,'bg_color':'#24435C','font_color':'white','text_wrap':True})
    s=w.add_worksheet('Hóa đơn GTGT');s.hide_gridlines(2)
    widths=[14,16,40,19,55,40,25,20,22,23,18,23,86,12,15,23,23,27,25,26,18,27,33]
    for c,x in enumerate(widths):s.set_column(c,c,x)
    notes=[r[0] for r in p['template_rows'][:7]]
    notes[6]='- Ngày chứng từ AMIS; TM/CK. Đã lọc danh sách loại trừ và bù trừ theo BHH gốc thực tế; chỉ nhập các dòng trong file mới này.'
    for i,v in enumerate(notes):
        if v:s.merge_range(i,0,i,22,v,wrap);s.set_row(i,26)
    headers=p['template_rows'][8]+['Số chứng từ AMIS','Tổng tiền hàng','Tổng tiền thanh toán','Dòng nguồn AMIS','Mã hàng hóa/dịch vụ','BTL đã bù trừ']
    s.write_row(8,0,headers,head);s.set_row(8,48)
    rr=9;count=0
    for n,d in enumerate((d for d in p['docs'] if d['category']=='Đưa vào import'),1):
        count=n
        for j,l in enumerate(d['lines']):
            r=l['values'];first=j==0
            values=[n,datetime.fromisoformat(d['date']) if first else None,d['buyer'] if first else None,d['tax'] if first else None,d['address'] if first else None,d['buyer'] if first else None,None,None,None,'TM/CK' if first else None,d['rate'] if first else None,d['sums']['vat'] if first else None,r[7],r[16],r[17],r[19],r[20],d['doc'],d['sums']['goods'] if first else None,d['sums']['payment'] if first else None,l['row'],r[12],', '.join(d['btls']) or None]
            for c,v in enumerate(values):
                f=date if c==1 else money if c in [11,16,18,19] else decimal if c in [14,15] else text if c in [3,17,21,22] else wrap if c in [2,4,5,12] else body
                if v is None or v=='':s.write_blank(rr,c,None,f)
                elif isinstance(v,datetime):s.write_datetime(rr,c,v,f)
                elif isinstance(v,(int,float)):s.write_number(rr,c,v,f)
                else:s.write_string(rr,c,str(v),f)
            s.set_row(rr,34);rr+=1
    s.autofilter(8,0,rr-1,22);s.freeze_panes(9,2)
    g=w.add_worksheet('Huong dan');g.hide_gridlines(2);g.set_column(0,0,33);g.set_column(1,1,125)
    g.write_row(0,0,['Mục','Nội dung'],head)
    guide=[['Cách nhập','Chọn sheet Hóa đơn GTGT, tiêu đề dòng 9; mẫu GTGT MTT một thuế suất. Mỗi STT là một hóa đơn.'],['Ngày và thanh toán','Lấy Ngày chứng từ AMIS, hình thức TM/CK theo xác nhận.'],['Tổng tiền',f'{count:,} hóa đơn / {rr-9:,} dòng. Tổng thanh toán sau bù trừ: {p["summary"]["Đưa vào import"]["net_payment"]:,} VND.'],['Bù trừ trả lại','Chỉ dùng liên kết BTL–BHH lấy trực tiếp AMIS. Đã loại 158 BHH trả hết; 2 BHH trả một phần được trừ đúng số lượng, tiền hàng và VAT.'],['Hai BHH đã trừ một phần','BHHT2024006103 và BHHT2024024887. Đơn giá giữ nguồn; tiền hàng/VAT còn lại = bán gốc - trả lại thực tế.'],['Cần kiểm tra','13 chứng từ và 7 mục loại trừ chưa rõ nằm trong KIEM_TRA_VA_LOAI_TRU_AMIS_2024.xlsx. Chưa đưa các chứng từ cần kiểm tra vào import.'],['Loại theo danh sách','115 chứng từ theo tab HD cty 2024, bao gồm mã tách/gộp đã ghi rõ. Không loại toàn bộ tab F0/F1/F2.'],['Người mua','Giữ tên, địa chỉ, MST nguồn. Email, điện thoại, CCCD để trống vì nguồn không có.'],['Thuế suất và tiền','8%, suy ra duy nhất từ tiền hàng/VAT trong nguồn; dung sai làm tròn 1đ/dòng. Tổng VAT chỉ ghi dòng đầu. Không tự cân tiền.'],['Cột đối chiếu','A:Q theo mẫu gốc. R:W là BHH, tổng tiền, dòng AMIS, mã hàng và BTL để đối chiếu.'],['Nguồn', 'So_chi_tiet_ban_hang_2024.xlsx; HD Xuất CTY 2024.xlsx; LIEN_KET_BTL_BHH_2024.xlsx; mẫu MTT người dùng cung cấp.'],['Thay thế bản cũ','Dùng bộ tại outputs/amis-2024-theo-btl-bhh-20261026. Bộ outputs/amis-2024-moi-20261026 đã hết hiệu lực.'],['Trạng thái','Chưa nhập, ký hoặc phát hành hóa đơn.']]
    guide=[(a,b.replace('13 chứng từ và 7 mục',f'{p["summary"]["Cần kiểm tra"]["documents"]} chứng từ và 7 mục')) for a,b in guide]
    guide.extend((c['doc'],c['basis']) for c in p['confirmations'])
    for i,row in enumerate(guide,1):g.write_string(i,0,row[0],body);g.write_string(i,1,row[1],wrap);g.set_row(i,44)
    w.close()
    # Bản bàn giao giữ đúng mẫu 17 cột; đối chiếu nằm trong workbook kiểm tra.
    from match_import_2024_template import main as match_template
    match_template()
    # Chuyển bảng trả chờ thành mẫu import điều chỉnh, giữ riêng hai BTL lệch giá.
    from prepare_return_adjustment_import import main as prepare_return_import
    prepare_return_import()
    (WORK/'export-route.json').write_text(json.dumps({'review':'Artifact Tool author/export/render','import':'Bundled XlsxWriter fallback; Artifact Tool large-workbook export was unavailable in this session (previous export hung >15 minutes and ~3.7GB RAM).','validation':'Independent openpyxl read-only; source hashes and per-line reconciliation'},ensure_ascii=False,indent=2),encoding='utf8')
    print(f'SAVED {file} ({count} invoices, {rr-9} lines)')

if __name__=='__main__':main()
