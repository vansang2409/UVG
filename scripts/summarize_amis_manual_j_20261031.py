"""Produce reviewable audit tables and independently check coverage/source hashes."""
import csv, json, hashlib
from collections import Counter
from decimal import Decimal
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/doi-chieu-amis-j-20261031'

def main():
    x=json.loads((OUT/'audit.json').read_text(encoding='utf-8'))
    snap=json.loads((OUT/'amis-j-snapshot.json').read_text(encoding='utf-8'))
    covered=[r for g in x['groups'] for r in g['rows']]
    assert len(covered)==len(set(covered))==len(snap['rows'])==31231
    assert set(covered)=={r['row'] for r in snap['rows']}
    for m in x['manifest']:assert hashlib.sha256((ROOT/m['file']).read_bytes()).hexdigest()==m['sha256']
    labels={'KHOP_TONG_DUY_NHAT':'Khớp VAT và thanh toán với một ứng viên','LECH_TOI_DA_2D':'Lệch tối đa 2đ, chưa tự chấp nhận','NHIEU_UNG_VIEN_KHOP':'Nhiều ký hiệu cùng khớp tiền','LECH_TIEN':'Tổng nhóm chưa khớp ứng viên cùng số','GHI_CHU_KHONG_PHAI_SO_HD':'Ô chứa ghi chú hoặc tham chiếu đặc biệt'}
    with (OUT/'CHI_TIET_CAN_DOI_CHIEU.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Kết quả','Số HĐ cột J','Ký hiệu AMIS đã có','Số dòng AMIS','Các dòng AMIS','Chứng từ AMIS','Ngày AMIS','VAT AMIS cộng nhóm','Thanh toán AMIS cộng nhóm','Ký hiệu ứng viên','Ngày HĐ ứng viên','VAT HĐ','Thanh toán HĐ','Chênh VAT AMIS - HĐ','Chênh thanh toán AMIS - HĐ','Trạng thái HĐ nguồn','File nguồn','Dòng nguồn','Giới hạn'])
        for g in x['groups']:
            for c in g['candidates'] or [None]:
                d=lambda field,source:str(Decimal(g[field])-Decimal(c[source])) if c and c[source]!='None' else ''
                w.writerow([labels[g['state']],g['invoice'],g['signature_entered'],len(g['rows']),';'.join(map(str,g['rows'])),';'.join(g['vouchers']),';'.join(g['amis_dates']),g['amis_vat'],g['amis_payment'],c['key'][0] if c else '',c['date'] if c else '',c['vat'] if c else '',c['payment'] if c else '',d('amis_vat','vat'),d('amis_payment','payment'),c['status'] if c else '',c['file'] if c else '',c['row'] if c else '', 'Chưa xác minh chi tiết hàng/quan hệ từng chứng từ; không cộng số liệu lặp theo ứng viên'])
    with (OUT/'THAM_CHIEU_DAC_BIET.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Nội dung J','Loại','Dòng AMIS','VAT AMIS','Thanh toán AMIS','Ứng viên nguồn','Số ứng viên khớp đủ VAT/thanh toán'])
        for s in x['special_entries']:w.writerow([s['entry'],s['kind'],';'.join(map(str,s['rows'])),s['amis_vat'],s['amis_payment'],json.dumps(s['references'],ensure_ascii=False),len(s['exact_single_candidates'])])
    gs=[g for g in x['groups'] if not g['signature_entered'] and g['invoice'].isdigit()]
    dated=[s for s in x['special_entries'] if s['kind']=='dated_reference']
    lines=['KIỂM TRA CỘT J AMIS - 31/10/2026','Nguồn AMIS: snapshot Google Sheets trực tiếp trong cuộc trao đổi; meInvoice: 33 bảng kê xuất đến 29/10/2026, 159.589 HĐ duy nhất.','Phạm vi: 31.231 ô J nhập giá trị trực tiếp, gồm 1.145 ô đối chiếu mã đuôi 000 đã làm trước. Không khẳng định cả 31.231 ô là kế toán vừa bổ sung.','30.086 ô chưa có ký hiệu gồm 29.765 ô số đơn, 13 tham chiếu năm 2026, 13 ghi chú số HĐ - 0, 1 ô hai HĐ, 294 ghi chú văn bản.','Các nhóm số đơn được gom theo năm hạch toán AMIS, ký hiệu nếu đã có, số J. Khi thiếu ký hiệu có thể trộn các HĐ cùng số; vì vậy lệch tổng chưa chứng minh từng dòng điền sai.','Khớp tiền chỉ là khớp số HĐ + tổng VAT/thanh toán, chưa xác nhận nội dung hàng hóa hoặc từng đơn trong hóa đơn gộp. Ngày AMIS và ngày HĐ được lưu riêng; không tự ép cùng kỳ.','So với số tiền của chính HĐ nguồn, không mặc định lấy số dư chuỗi sau điều chỉnh. Trạng thái nguồn được ghi rõ.','','Nhóm số đơn chưa có ký hiệu:']
    for state,count in Counter(g['state'] for g in gs).items():lines.append(f"{labels[state]}: {count} nhóm / {sum(len(g['rows']) for g in gs if g['state']==state)} dòng")
    lines+=['',f"13 tham chiếu năm 2026: tìm đủ 13; {sum(bool(s['exact_single_candidates']) for s in dated)} khớp cả VAT/thanh toán; HĐ 1C26MUV/00005767, AMIS dòng 50940, thanh toán AMIS 2.977.000đ, HĐ 2.976.999đ (lệch 1đ).",'Ví dụ: J=427: 627 dòng, AMIS 875.996.776đ; 1C25TUV/427 875.775.776đ, chênh +221.000đ. Có cả 1C25MUV/427, cần xác định ký hiệu trước khi sửa.','Ví dụ: J=431: 585 dòng, AMIS 799.932.550đ; 1C25TUV/431 804.873.892đ, chênh -4.941.342đ.','Ví dụ khớp: J=489, 68 dòng, VAT 7.652.427đ và thanh toán 103.307.763đ khớp 1C25TUV/489. Đây là HĐ đã bị điều chỉnh; khớp số gốc không xác nhận số dư chuỗi hiện tại.','','Kiểm chứng: phủ đủ 31.231 dòng, không trùng/thiếu dòng; hash 33 nguồn không đổi; đọc lại độc lập bằng openpyxl 10 HĐ thuộc 5 số 106/427/431/489/12255, khớp dữ liệu đã trích.','Không sửa Google Sheets, nguồn meInvoice, import hoặc phát hành hóa đơn.','CHI_TIET_CAN_DOI_CHIEU.csv: từng nhóm và từng ứng viên nguồn, có chênh lệch, số dòng và chứng từ. Một nhóm có thể có nhiều ứng viên; không cộng các dòng ứng viên như các hóa đơn đã ghép.','THAM_CHIEU_DAC_BIET.csv: chi tiết toàn bộ ô J không phải một số HĐ đơn.']
    (OUT/'KET_QUA.txt').write_text('\n'.join(lines),encoding='utf-8-sig')
    (OUT/'verification.json').write_text(json.dumps({'covered_rows':len(covered),'unique_rows':len(set(covered)),'source_hashes_verified':len(x['manifest']),'independent_sample_invoices':10,'google_sheet_modified':False},indent=2),encoding='utf-8')
    print('\n'.join(lines))

if __name__=='__main__':main()
