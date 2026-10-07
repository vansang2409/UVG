"""Tạo danh sách cần bổ sung bằng chứng; không quyết định phát hành hóa đơn."""
from pathlib import Path
from collections import Counter
import argparse
import csv
from datetime import datetime
from decimal import Decimal
import json

ROOT = Path(__file__).resolve().parents[1]
RULES = {
    'Thiếu liên kết tới hóa đơn nguồn': 'XML hóa đơn này và hóa đơn nguồn; danh sách điều chỉnh/thay thế đầy đủ, gồm các lần sửa tiếp theo.',
    'Thiếu số tiền trên bảng tổng': 'XML và bảng kê đầy đủ tiền hàng, thuế, tổng thanh toán của hóa đơn này.',
    'Tiền hàng cộng thuế không bằng tổng thanh toán': 'XML và bản hiển thị hóa đơn; kiểm tra tiền hàng, thuế, tổng thanh toán và làm tròn.',
}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'outputs')
    parser.add_argument('--json', type=Path, help='File trung gian tùy chọn để tạo Excel.')
    args = parser.parse_args()
    source = args.output_dir / 'uvg-doi-chieu-nhom' / 'HOA_DON_VA_KY_GOC.csv'
    dest = args.output_dir / 'uvg-kiem-tra-uu-tien'
    dest.mkdir(parents=True, exist_ok=True)
    with source.open(encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        original_headers = reader.fieldnames
        records = list(reader)
    counts = Counter()
    selected = []
    keys = set()
    extra = ['Nhóm kiểm tra ưu tiên', 'Tài liệu cần bổ sung', 'Tình trạng kiểm tra',
             'File bằng chứng đã bổ sung', 'Kết quả kiểm tra', 'Hướng xử lý đề xuất', 'Người kiểm tra']
    for row in records:
        flags = [flag for flag in RULES if flag in row['Cần kiểm tra']]
        if not flags:
            continue
        key = (row['Ký hiệu'], row['Số hóa đơn'])
        if key in keys:
            raise ValueError(f'Trùng khóa hóa đơn: {key}')
        keys.add(key)
        counts.update(flags)
        selected.append({**row, 'Nhóm kiểm tra ưu tiên': '; '.join(flags),
                         'Tài liệu cần bổ sung': ' '.join(RULES[flag] for flag in flags),
                         'Tình trạng kiểm tra': 'Chưa kiểm tra',
                         **{col: '' for col in extra[3:]}})
    selected.sort(key=lambda row: (datetime.strptime(row['Ngày phát hành'], '%d/%m/%Y'),
                                   row['Ký hiệu'], row['Số hóa đơn']))
    headers = original_headers + extra
    with (dest / 'HOA_DON_CAN_KIEM_TRA_UU_TIEN.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(selected)
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        rows = []
        for row in selected:
            values = [row[h] for h in headers]
            for col in ['Doanh thu', 'Thuế', 'Thanh toán', 'Dòng']:
                idx = headers.index(col)
                values[idx] = float(Decimal(values[idx])) if values[idx] != '' else None
            rows.append(values)
        args.json.write_text(json.dumps({'headers': headers, 'rows': rows, 'counts': dict(counts)},
                                       ensure_ascii=False), encoding='utf-8')
    summary = ['# UVG – danh sách kiểm tra ưu tiên', '',
               f'Có **{len(selected)} hóa đơn riêng biệt** được chọn từ {len(records)} dòng hóa đơn trong báo cáo hiện có.', '',
               '| Vấn đề | Số hóa đơn |', '|---|---:|']
    summary += [f'| {flag} | {counts[flag]} |' for flag in RULES]
    summary += ['', 'Các nhóm có thể trùng nhau. Không cộng các số nhóm để ra số hóa đơn riêng biệt.', '',
                'Đây là danh sách cần kiểm tra, chưa phải danh sách hóa đơn sai hoặc cần phát hành.',
                'Các hóa đơn ngoài danh sách vẫn có thể thiếu chi tiết sản phẩm, thiếu lượng hoặc lệch kỳ.', '',
                '## Cách dùng', '',
                '1. Mở HOA_DON_CAN_KIEM_TRA_UU_TIEN.xlsx, lọc theo Nhóm kiểm tra ưu tiên.',
                '2. Tra cứu bằng Ký hiệu + Số hóa đơn. File nguồn và Dòng giúp mở lại đúng dòng trong báo cáo nguồn.',
                '3. Bổ sung tài liệu ở cột Tài liệu cần bổ sung. Giữ XML và tài liệu liên quan trong data/Bo sung.',
                '4. Điền các cột từ Tình trạng kiểm tra trở đi. Ghi đường dẫn tương đối của bằng chứng để máy khác mở được.',
                '5. Chỉ ghi hướng xử lý khi đã đủ căn cứ. Đề xuất không đồng nghĩa với được phép phát hành.', '',
                'CSV giữ thông tin gốc dạng chữ, số hóa đơn có số 0 đầu. Nếu mở CSV trong Excel, dùng Data > From Text/CSV và chọn kiểu Text cho Ký hiệu, Số hóa đơn, Ký hiệu gốc, Số HĐ gốc.',
                'Giá trị tiền bị thiếu được giữ trống. Không coi ô trống là 0.', '',
                '## Tái tạo', '', '```powershell',
                'py scripts/prepare_priority_review.py', '```', '',
                'Lệnh trên ghi lại CSV và báo cáo thống kê từ kết quả đối chiếu hiện tại. Nó không đọc các ghi chú bạn đã nhập trong Excel.',
                'Excel là bản làm việc đã tạo riêng. Lưu bản có ghi chú với tên khác trước khi nhận hoặc tạo bản mới.']
    (dest / 'KIEM_TRA_UU_TIEN.md').write_text('\n'.join(summary) + '\n', encoding='utf-8')
    print(json.dumps({'unique_invoices': len(selected), 'counts': dict(counts)}, ensure_ascii=True))


if __name__ == '__main__':
    main()
