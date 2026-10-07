"""Lọc ngoại lệ giữ nguyên theo người mua và MST, giữ nguyên các file nguồn."""
import csv
import argparse
import hashlib
import json
import re
import unicodedata
from collections import Counter, defaultdict, deque
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/ngoai-le-meinvoice'


def txt(v):
    return '' if v is None else str(v).strip()


def normal(v):
    s = unicodedata.normalize('NFD', txt(v).lower()).replace('đ', 'd')
    return ' '.join(''.join(c for c in s if not unicodedata.combining(c)).split())


def is_company(v):
    return bool(re.search(r'\b(cong ty|cty|tnhh|doanh nghiep|dntn|co phan)\b', normal(v)))


def tax_shaped(v):
    # Chỉ nhận diện cấu trúc trong file, không xác nhận MST còn hoạt động.
    return bool(re.fullmatch(r'\d{10}|\d{13}|\d{10}-\d{3}', txt(v)))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--year', type=int, help='Năm hóa đơn gốc cần xử lý.')
    args = parser.parse_args()
    out = ROOT / f'outputs/ngoai-le-meinvoice-{args.year}' if args.year else OUT
    records = defaultdict(lambda: {'observations': [], 'headers': []})
    graph = defaultdict(set)
    parents = defaultdict(set)
    links = []
    for entry in json.loads((ROOT / 'SOURCE_MANIFEST.json').read_text(encoding='utf-8')):
        p = ROOT / entry['path']
        assert hashlib.sha256(p.read_bytes()).hexdigest() == entry['sha256'], p
    def observe(key, name, code, date, source, status='', money=None, header=False):
        assert key[0] and key[1], (source, key)
        obs = {'name': txt(name), 'code': txt(code), 'date': date.strftime('%d/%m/%Y') if isinstance(date, datetime) else txt(date),
               'source': source, 'status': status, 'money': money}
        records[key]['observations'].append(obs)
        if header:
            records[key]['headers'].append(obs)
    for p in sorted((ROOT / 'data/Du lieu Meinvoice').glob('*.xls')):
        if p.name.startswith('Mau'):
            continue
        with p.open('rb') as stream:
            wb = load_workbook(stream, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            for i, r in enumerate(ws.values, 1):
                if not isinstance(r[0], (int, float)):
                    continue
                source = f'{p.relative_to(ROOT).as_posix()} | {ws.title} | dòng {i}'
                if p.name.startswith('Bang_ke'):
                    observe((txt(r[1]), txt(r[2])), r[7], r[9], r[3], source, txt(r[20]),
                            [v if isinstance(v, (int, float)) else None for v in r[15:18]], True)
                elif p.name.startswith('Bảng kê chi tiết'):
                    observe((txt(r[1]), txt(r[2])), r[4], r[6], r[3], source, txt(r[23]))
                elif p.name.startswith('DS'):
                    old, new = (txt(r[1]), txt(r[3])), (txt(r[12]), txt(r[14]))
                    observe(old, r[5], r[6], r[2], source)
                    observe(new, r[16], r[17], r[13], source)
                    graph[old].add(new)
                    graph[new].add(old)
                    parents[new].add(old)
                    links.append([*old, *new, 'Điều chỉnh' if 'điều chỉnh' in p.name else 'Thay thế', source])
            wb.close()
    direct, review = set(), set()
    for key, rec in records.items():
        obs = rec['observations']
        if any(is_company(o['name']) or tax_shaped(o['code']) for o in obs):
            direct.add(key)
        # Khác MST, số định danh không nhận diện được, hoặc tổ chức thiếu MST cần xem lại.
        codes = {o['code'] for o in obs if o['code']}
        if len(codes) > 1 or any(not tax_shaped(c) for c in codes) or (key not in direct and any(re.search(r'\b(chi nhanh|ho kinh doanh|truong|vien|ngan hang|hop tac xa)\b', normal(o['name'])) for o in obs)):
            review.add(key)
    protected = set(direct)
    origins = defaultdict(set)
    visited = set()
    for start in sorted(records):
        if start in visited:
            continue
        component, queue = set(), deque([start])
        while queue:
            key = queue.popleft()
            if key in component:
                continue
            component.add(key)
            queue.extend(graph[key] - component)
        visited.update(component)
        seeds = component & direct
        if seeds:
            protected.update(component)
            for key in component:
                origins[key] = seeds
    headers = ['Ký hiệu', 'Số hóa đơn', 'Ngày hóa đơn', 'Tên người mua', 'MST/CCCD chủ hộ nguyên bản',
               'Trạng thái theo bảng tổng', 'Phân loại', 'Căn cứ nhận diện', 'Tiền hàng theo bảng tổng',
               'Thuế theo bảng tổng', 'Thanh toán theo bảng tổng', 'Các MST/định danh trong nguồn',
               'Ngoại lệ trực tiếp trong chuỗi', 'Cần rà soát', 'Vị trí nguồn nhận diện']
    def row(key, category):
        rec = records[key]
        obs = rec['observations']
        best = rec['headers'][-1] if rec['headers'] else obs[0]
        positive = [o for o in obs if is_company(o['name']) or tax_shaped(o['code'])]
        display = positive[0] if positive else best
        reasons = []
        if any(is_company(o['name']) for o in obs):
            reasons.append('Tên người mua nhận diện công ty/doanh nghiệp')
        if any(tax_shaped(o['code']) for o in obs):
            reasons.append('Cột MST/CCCD có giá trị dạng MST 10 hoặc 13 chữ số')
        if not reasons:
            reasons.append('Có liên kết DS với hóa đơn ngoại lệ; giữ riêng để kiểm tra chuỗi')
        codes = sorted({o['code'] for o in obs if o['code']})
        concerns = []
        if key in review:
            concerns.append('Kiểm tra định danh hoặc thông tin người mua khác nhau giữa các nguồn')
        if not rec['headers']:
            concerns.append('Không có trong bảng tổng; không suy ra tiền hoặc trạng thái hiện hành')
        if len({(o['name'], o['code']) for o in obs if o['name'] or o['code']}) > 1:
            concerns.append('Có khác biệt tên/định danh giữa các nguồn')
        return [*key, best['date'], display['name'], display['code'], best['status'] or 'Chưa có trạng thái bảng tổng',
                category, '; '.join(reasons), *(best['money'] or [None, None, None]), '; '.join(codes),
                '; '.join('/'.join(k) for k in sorted(origins[key])), '; '.join(concerns),
                '; '.join(dict.fromkeys(o['source'] for o in (positive or obs)))]
    def sorted_rows(keys, category):
        return [row(k, category) for k in sorted(keys)]
    related_later = set()
    root_year_keys = set()
    if args.year:
        def invoice_year(key):
            rec = records[key]
            chosen = rec['headers'][-1] if rec['headers'] else rec['observations'][0]
            return datetime.strptime(chosen['date'], '%d/%m/%Y').year
        # Chỉ lấy chuỗi có gốc đúng năm yêu cầu, bằng liên kết DS rõ ràng.
        root_year_keys = {k for k in protected if not parents[k] and invoice_year(k) == args.year}
        scope = set()
        queue = deque(root_year_keys)
        while queue:
            key = queue.popleft()
            if key in scope:
                continue
            scope.add(key)
            queue.extend(graph[key] - scope)
        direct = direct & scope
        related_later = {k for k in scope if invoice_year(k) != args.year}
        direct -= related_later
        protected = scope - related_later
        review &= scope
    direct_rows = sorted_rows(direct, 'Ngoại lệ theo thông tin nguồn')
    chain_rows = sorted_rows(protected - direct, 'Liên quan chuỗi ngoại lệ – cần giữ để rà soát')
    review_rows = sorted_rows(review, 'Cần kiểm tra định danh')
    out.mkdir(parents=True, exist_ok=True)
    datasets = [('Ngoai le', 'HOA_DON_NGOAI_LE.csv', direct_rows),
                ('Lien quan chuoi', 'LIEN_QUAN_CHUOI_NGOAI_LE.csv', chain_rows),
                ('Can ra soat', 'CAN_RA_SOAT.csv', review_rows)]
    if args.year:
        datasets.append(('Lich su sau 2024', 'LICH_SU_LIEN_QUAN_SAU_2024.csv', sorted_rows(related_later, f'Lịch sử liên quan hóa đơn gốc {args.year}')))
    for name, filename, rows in datasets:
        with (out / filename).open('w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            w.writerow(headers)
            w.writerows(rows)
    summary = {'unique_invoices_all_sources': len(records), 'unique_invoices_header': sum(bool(r['headers']) for r in records.values()),
               'direct_exceptions': len(direct), 'related_chain_only': len(protected-direct), 'protected_with_chains': len(protected),
               'direct_without_header': sum(not records[k]['headers'] for k in direct),
               'review': len(review), 'direct_status': dict(Counter(r[5] for r in direct_rows)),
               'direct_with_company_name': sum(any(is_company(o['name']) for o in records[k]['observations']) for k in direct),
               'direct_with_tax_shaped_id': sum(any(tax_shaped(o['code']) for o in records[k]['observations']) for k in direct)}
    summary.update({'root_year': args.year, 'roots_in_year': len(root_year_keys), 'later_history': len(related_later)})
    payload = {'headers': headers, 'sheets': [{'name': name, 'rows': rows} for name, _, rows in datasets if rows], 'summary': summary}
    (out / 'workbook_data.json').write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
    (out / 'THONG_KE.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    report = ['# UVG – hóa đơn ngoại lệ giữ nguyên', '', 'Ngày lọc: 07/10/2026. Nguồn: tám bảng kê meInvoice trong data/Du lieu Meinvoice, không dùng file mẫu điều chỉnh.', '',
        f"- Ngoại lệ trực tiếp theo thông tin nguồn: **{len(direct)} hóa đơn**.",
        f"- Hóa đơn chỉ liên quan qua chuỗi DS: **{len(protected-direct)} hóa đơn**, tách riêng để rà soát, chưa tự kết luận người mua thuộc ngoại lệ.",
        f"- Hóa đơn cần rà soát định danh: **{len(review)}**; có thể trùng danh sách ngoại lệ.", '',
        '## Quy tắc', '',
        'Giữ ngoại lệ nếu bất kỳ bảng tổng, bảng chi tiết hoặc bảng DS nào ghi tên người mua nhận diện công ty/doanh nghiệp, hoặc cột MST/CCCD chủ hộ có giá trị dạng MST 10 chữ số, 13 chữ số hoặc 10 chữ số + dấu gạch + 3 chữ số. Đây là nhận diện cấu trúc dữ liệu, chưa tra cứu xác thực MST. Không mặc định mọi số định danh là MST.', '',
        'Tên người mua không phải công ty nhưng có dạng MST vẫn thuộc ngoại lệ theo yêu cầu người dùng. Nhóm này có thể gồm cá nhân, hộ kinh doanh hoặc đơn vị khác; chưa suy ra loại pháp nhân chỉ từ tên.', '',
        'Lọc tất cả trạng thái, kể cả đã hủy/đã bị thay thế để giữ đầy đủ lịch sử. Số lượng ngoại lệ không đồng nghĩa số hóa đơn hiện hành hoặc doanh thu còn hiệu lực. Không cộng các hóa đơn gốc và bản thay thế để ra doanh thu.', '',
        'Không tự suy ra liên kết từ mã đơn sàn, tên, MST hoặc số tiền. Bảng liên quan chuỗi chỉ lần theo liên kết rõ ràng trong ba bảng DS. Nếu nguồn khác nhau về người mua/định danh, giữ dấu vết và cờ rà soát.', '',
        'Tiền và trạng thái lấy từ dòng bảng tổng chọn theo thứ tự tên file như script đối chiếu cũ; không có bảng tổng thì để trống. Tên/định danh trình bày từ dòng có căn cứ ngoại lệ và ghi vị trí nguồn ở cuối bảng.', '',
        '## File kết quả', '', 'HOA_DON_NGOAI_LE.csv và HOA_DON_NGOAI_LE.xlsx: danh sách ngoại lệ trực tiếp. LIEN_QUAN_CHUOI_NGOAI_LE.csv: các hóa đơn chỉ liên quan chuỗi. CAN_RA_SOAT.csv: định danh cần kiểm tra.', '',
        '11 file nguồn đã kiểm tra SHA-256 khớp SOURCE_MANIFEST.json và giữ nguyên. Chưa đối chiếu AMIS để loại doanh thu ngoại lệ, chưa sửa/import/ký/phát hành hóa đơn. Bảng tổng bao phủ 2024–06/2025. Ngoại lệ 1C25TUV/00012764 ngày 15/07/2025 có trong DS bị điều chỉnh 2024.xls, dòng 574, nhưng không có bảng tổng. Bộ DS có ít nhất một lần sửa sau 06/2025; chưa chứng minh lịch sử đã đầy đủ.', '',
        'Tái chạy: `python scripts/filter_invoice_exceptions.py`. Các CSV là kết quả tái tạo, không dùng để giữ ghi chú thủ công.', '']
    if args.year:
        report = [f'# UVG – ngoại lệ hóa đơn gốc năm {args.year}', '',
                  f'- {len(direct)} hóa đơn ngoại lệ có ngày trong năm {args.year}, thuộc {len(root_year_keys)} gốc theo liên kết hiện có.',
                  f'- {len(related_later)} hóa đơn phát sinh sau năm {args.year}, có liên kết DS với các gốc trên, nằm ở bảng Lich su sau 2024.',
                  '- Không đưa hóa đơn gốc năm 2025 vào danh sách chính. Giữ các lần sửa sau 2024 để không bỏ mất lịch sử.',
                  '- Các khoản gốc và thay thế không được cộng tùy ý thành doanh thu. Chưa xác minh XML hoặc lịch sử đầy đủ.',
                  '- Ngoại lệ nhận diện bằng tên công ty/doanh nghiệp hoặc định danh dạng MST trong nguồn, chưa tra cứu xác thực MST.',
                  '- File nguồn giữ nguyên, đã kiểm tra SHA-256 của 11 file. Số hóa đơn/MST giữ kiểu chữ, tiền thiếu giữ trống.',
                  f'- Tái chạy: `python scripts/filter_invoice_exceptions.py --year {args.year}`.', '']
    (out / 'NGOAI_LE_MEINVOICE.md').write_text('\n'.join(report), encoding='utf-8')
    assert len(direct_rows) == len(direct)
    assert not (direct & (protected-direct))
    print(json.dumps(summary, ensure_ascii=True))


if __name__ == '__main__':
    main()
