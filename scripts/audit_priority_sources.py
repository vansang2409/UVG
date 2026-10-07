"""Rà soát 321 hóa đơn bằng file có sẵn, không sửa nguồn hay quyết định phát hành."""
from pathlib import Path
from collections import defaultdict, Counter
from decimal import Decimal
import csv
import hashlib
import json
import re
import sys
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/uvg-kiem-tra-uu-tien/ra-soat-2026-10-07'


def numeric(value):
    # Ô trống phải giữ trống; không quy đổi sang số 0.
    return Decimal(str(value)) if isinstance(value, (int, float)) else None


def text(value):
    return '' if value is None else str(value)


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    source = ROOT / 'outputs/uvg-kiem-tra-uu-tien/HOA_DON_CAN_KIEM_TRA_UU_TIEN.csv'
    with source.open(encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames
        priority = list(reader)
    keys = {(r['Ký hiệu'], r['Số hóa đơn']) for r in priority}
    assert len(keys) == len(priority) == 321
    # Kiểm tra toàn bộ file nguồn bằng checksum của bản bàn giao.
    manifest = json.loads((ROOT / 'SOURCE_MANIFEST.json').read_text(encoding='utf-8'))
    for entry in manifest:
        path = ROOT / entry['path']
        assert path.stat().st_size == entry['bytes'], path
        assert hashlib.sha256(path.read_bytes()).hexdigest() == entry['sha256'], path
    totals, lines, parents, related = (defaultdict(list) for _ in range(4))
    evidence = []
    invoice_numbers = defaultdict(set)
    for path in sorted((ROOT / 'data/Du lieu Meinvoice').glob('*.xls')):
        if path.name.startswith('Mau'):
            continue
        with path.open('rb') as stream:
            wb = load_workbook(stream, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            for rownum, r in enumerate(ws.values, 1):
                if not isinstance(r[0], (int, float)):
                    continue
                ref = f'{path.relative_to(ROOT).as_posix()} | {ws.title} | dòng {rownum}'
                if path.name.startswith('Bang_ke'):
                    key = (text(r[1]), text(r[2]))
                    if key[1].isdigit():
                        invoice_numbers[int(key[1])].add(key)
                    if key in keys:
                        totals[key].append((tuple(numeric(r[j]) for j in (15, 16, 17)), text(r[20]), ref))
                elif path.name.startswith('Bảng kê chi tiết'):
                    key = (text(r[1]), text(r[2]))
                    if key in keys:
                        lines[key].append((tuple(numeric(r[j]) for j in (15, 17, 18)), text(r[11]), text(r[10]), ref))
                elif path.name.startswith('DS'):
                    old, new = (text(r[1]), text(r[3])), (text(r[12]), text(r[14]))
                    kind = 'Điều chỉnh' if 'điều chỉnh' in path.name else 'Thay thế'
                    parents[new].append((old, kind, ref))
                    if old in keys:
                        related[old].append((tuple(numeric(r[j]) for j in (9, 10, 11)), new, kind, ref))
                    if new in keys:
                        related[new].append((tuple(numeric(r[j]) for j in (20, 21, 22)), old, kind, ref))
            wb.close()
    extra = ['Số bảng tổng tìm thấy', 'Số dòng chi tiết', 'Liên kết nguồn tìm thấy',
             'Tiền hàng dòng chi tiết đủ số', 'Thuế dòng chi tiết đủ số',
             'Thanh toán dòng chi tiết đủ số', 'Số dòng thiếu tiền hàng hoặc thuế',
             'Tiền hàng cộng thuế trừ thanh toán', 'Căn cứ file có sẵn', 'Kết quả rà soát file',
             'Bằng chứng cần lấy tiếp', 'Số HĐ nhắc trong diễn giải - chưa xác minh',
             'Ứng viên theo số - chưa phải liên kết', 'Diễn giải và vị trí nguồn']
    results, counts = [], Counter()
    for original in priority:
        row = dict(original)
        key = (row['Ký hiệu'], row['Số hóa đơn'])
        hs, ds, ps, rs = totals[key], lines[key], parents[key], related[key]
        expected = (f"data/Du lieu Meinvoice/{row['File nguồn']} | dòng {row['Dòng']}")
        assert any(expected.split(' | ')[0] in h[2] and h[2].endswith(f"dòng {row['Dòng']}") for h in hs), key
        primary = next(h for h in hs if h[2].startswith(expected.split(' | ')[0]) and h[2].endswith(f"dòng {row['Dòng']}"))
        amounts = primary[0]
        # Không cộng dòng thiếu giá trị vào một tổng được trình bày là đầy đủ.
        sums = [sum((d[0][j] for d in ds), Decimal(0)) if ds and all(d[0][j] is not None for d in ds) else None for j in range(3)]
        missing = sum(any(v is None for v in d[0][:2]) for d in ds)
        residual = amounts[0] + amounts[1] - amounts[2] if all(v is not None for v in amounts) else None
        notes, needed = [], []
        if 'Thiếu liên kết' in row['Nhóm kiểm tra ưu tiên']:
            if ps:
                notes.append('Tìm thấy liên kết trực tiếp trong bảng DS; cần xác minh XML và toàn chuỗi.')
                counts['Tìm thấy liên kết'] += 1
            else:
                notes.append('Không tìm thấy liên kết nguồn trong cả ba bảng DS hiện có.')
                counts['299 chưa tìm thấy liên kết'] += 1
            needed.append('XML hóa đơn và hóa đơn tham chiếu; DS điều chỉnh/thay thế đầy đủ kể cả sau 06/2025.')
        if 'Thiếu số tiền' in row['Nhóm kiểm tra ưu tiên']:
            complete = {h[0] for h in hs if all(v is not None for v in h[0])}
            notes.append('Các ô bảng tổng theo thứ tự tiền hàng/thuế/thanh toán: ' + '/'.join('TRỐNG' if v is None else str(v) for v in amounts) + '.')
            notes.append(f'{len(complete)} bộ giá trị đầy đủ ở bảng tổng khác; {len(rs)} dòng trong bảng liên kết; {len(ds)} dòng chi tiết.')
            counts['Thiếu tiền có bảng tổng khác đầy đủ'] += bool(complete)
            needed.append('XML xác nhận tiền hàng, thuế, thanh toán; không điền trống bằng 0.')
        if residual is not None and residual != 0:
            notes.append(f'Bảng tổng: tiền hàng + thuế - thanh toán = {residual} đồng.')
            counts['8 lệch số học xác nhận lại'] += 1
            needed.append('XML và bản hiển thị để giải thích chênh lệch, chiết khấu và làm tròn.')
        if ds:
            notes.append(f'Có {len(ds)} dòng chi tiết, {missing} dòng thiếu tiền hàng hoặc thuế.')
        else:
            notes.append('Không có dòng chi tiết theo đúng ký hiệu + số hóa đơn.')
            counts['Không có chi tiết'] += 1
        versions = {(h[0], h[1]) for h in hs}
        if len(versions) > 1:
            notes.append('Bảng tổng có nhiều phiên bản giá trị/trạng thái; cần đối chiếu thời điểm xuất.')
            counts['Nhiều phiên bản bảng tổng'] += 1
        refs = list(dict.fromkeys([h[2] for h in hs] + [d[3] for d in ds] + [p[2] for p in ps] + [r[3] for r in rs]))
        mentions = sorted({int(n) for d in ds for n in re.findall(r'h[oóò]a\s+đơn\s+số\s+(\d+)', d[1], re.I)})
        candidates = sorted({k for n in mentions for k in invoice_numbers[n]})
        if 'Thiếu liên kết' in row['Nhóm kiểm tra ưu tiên'] and mentions:
            counts['Thiếu liên kết có số HĐ trong diễn giải'] += 1
            notes.append('Diễn giải có số hóa đơn tham chiếu nhưng chưa xác định ký hiệu; các ứng viên chỉ dùng tra cứu.')
        row.update(dict(zip(extra, [len(hs), len(ds), '; '.join('/'.join(p[0]) for p in ps),
            *['' if v is None else str(v) for v in sums], missing,
            '' if residual is None else str(residual), '; '.join(refs), ' '.join(notes), ' '.join(needed),
            '; '.join(str(n).zfill(8) for n in mentions), '; '.join('/'.join(k) for k in candidates),
            '; '.join(f'{d[1]} [{d[3]}]' for d in ds)])))
        # Chỉ ghi kết quả rà soát file, không đổi kết luận nghiệp vụ hoặc ghi chú cũ.
        results.append(row)
        for kind, entries in [('Bảng tổng', hs), ('Chi tiết', ds)]:
            for e in entries:
                evidence.append([*key, kind, e[-1], *['' if v is None else str(v) for v in e[0]]])
        for amounts2, other, kind, ref in rs:
            evidence.append([*key, f'{kind}: liên quan {other[0]}/{other[1]}', ref, *['' if v is None else str(v) for v in amounts2]])
    OUT.mkdir(parents=True, exist_ok=True)
    for filename, fields, records in [('KET_QUA_RA_SOAT_321.csv', headers + extra, results)]:
        with (OUT / filename).open('w', encoding='utf-8-sig', newline='') as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(records)
    with (OUT / 'BANG_CHUNG_TU_FILE.csv').open('w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Ký hiệu', 'Số hóa đơn', 'Loại nguồn', 'Vị trí nguồn', 'Tiền hàng', 'Thuế', 'Thanh toán'])
        w.writerows(evidence)
    summary = {'invoices': len(results), 'source_checksums_verified': len(manifest), 'counts': dict(counts),
               'supplemental_files': len(list((ROOT / 'data/Bo sung').rglob('*'))) if (ROOT / 'data/Bo sung').exists() else 0}
    (OUT / 'THONG_KE.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    money = [r for r in results if 'Thiếu số tiền' in r['Nhóm kiểm tra ưu tiên'] or 'Tiền hàng' in r['Nhóm kiểm tra ưu tiên']]
    with (OUT / 'LO_22_CAN_LAY_XML.csv').open('w', encoding='utf-8-sig', newline='') as f:
        fields = ['Ký hiệu', 'Số hóa đơn', 'Ngày phát hành', 'Nhóm kiểm tra ưu tiên',
                  'Kết quả rà soát file', 'Bằng chứng cần lấy tiếp', 'Căn cứ file có sẵn']
        w = csv.DictWriter(f, fieldnames=fields, extrasaction='ignore')
        w.writeheader()
        w.writerows(money)
    assert len(money) == 22
    mentioned = [r for r in results if 'Thiếu liên kết' in r['Nhóm kiểm tra ưu tiên'] and r[extra[-3]]]
    def fmt(v):
        return f'{Decimal(v):,.0f}'.replace(',', '.') if v != '' else 'Trống'
    report = ['# UVG – rà soát file của 321 hóa đơn ưu tiên', '',
        'Ngày kiểm tra: 07/10/2026. Phạm vi: bộ file đã bàn giao, 2024–06/2025. Đây là kiểm tra file, chưa xác minh XML hoặc lịch sử CQT.', '',
        '## Kết quả', '',
        '- Đã rà 321/321 khóa ký hiệu + số hóa đơn và truy lại đúng dòng bảng tổng được báo cáo trích dẫn.',
        '- 11/11 file nguồn khớp dung lượng và SHA-256 của SOURCE_MANIFEST.json.',
        '- Không có file bổ sung trong data/Bo sung. Chưa thể xác minh lịch sử sau kỳ xuất file.',
        '- 299/299 hóa đơn thiếu liên kết chưa tìm được hóa đơn nguồn trong ba bảng DS. Có 39 hóa đơn nhắc số hóa đơn trong diễn giải; chỉ lập ứng viên tra cứu, chưa xác nhận ký hiệu hoặc liên kết.',
        '- 14/14 hóa đơn thiếu tiền chưa có bảng tổng khác cung cấp đủ ba giá trị. Ô trống giữ nguyên; phần tiền có sẵn được trình bày trong kết quả rà soát.',
        '- 8/8 trường hợp lệch số học được xác nhận lại từ bảng tổng. Cả 8 có một dòng chi tiết thiếu thành tiền; chưa dùng chi tiết để giải thích được chênh lệch.',
        '- 4 hóa đơn không có dòng chi tiết theo đúng khóa. Các bảng tổng của 321 hóa đơn không có phiên bản khác nhau về ba giá trị tiền/trạng thái trong bộ file.', '',
        '## Tám trường hợp lệch tiền', '',
        'Chênh lệch = tiền hàng + thuế − tổng thanh toán, đơn vị đồng. Đây là phép kiểm tra số học, không phải số tiền cần điều chỉnh.', '',
        '| Ký hiệu/số hóa đơn | Tiền hàng | Thuế | Thanh toán | Chênh lệch |',
        '|---|---:|---:|---:|---:|']
    for r in money:
        if 'Tiền hàng' in r['Nhóm kiểm tra ưu tiên']:
            report.append(f"| {r['Ký hiệu']}/{r['Số hóa đơn']} | {fmt(r['Doanh thu'])} | {fmt(r['Thuế'])} | {fmt(r['Thanh toán'])} | {fmt(r['Tiền hàng cộng thuế trừ thanh toán'])} |")
    report += ['',
        'Ưu tiên XML 1C25TUV/00000649: tiền hàng và thuế âm, tổng thanh toán dương; cần xác minh dấu và nội dung điều chỉnh. Tiếp theo 00000561 lệch 136.000 đồng; các trường hợp 00001819, 00001167, 00002145 lệch 2.000, 1.000, 900 đồng. Ba trường hợp 00000626, 00001821, 00002619 lệch tuyệt đối 1 đồng; chưa kết luận là làm tròn khi chưa có XML.', '',
        '## Mười bốn hóa đơn thiếu tiền', '',
        'Bảng sau giữ đúng ô bảng tổng; khác với bảng ưu tiên cũ vốn để trống cả bộ ba khi bất kỳ giá trị nào thiếu.', '',
        '| Ký hiệu/số hóa đơn | Tiền hàng nguồn | Thuế nguồn | Thanh toán nguồn | Số dòng chi tiết |',
        '|---|---:|---:|---:|---:|']
    for r in money:
        if 'Thiếu số tiền' in r['Nhóm kiểm tra ưu tiên']:
            key = (r['Ký hiệu'], r['Số hóa đơn'])
            h = next(h for h in totals[key] if h[2].startswith(f"data/Du lieu Meinvoice/{r['File nguồn']} |") and h[2].endswith(f"dòng {r['Dòng']}"))
            report.append(f"| {key[0]}/{key[1]} | " + ' | '.join('Trống' if v is None else fmt(str(v)) for v in h[0]) + f" | {len(lines[key])} |")
    report += ['',
        '1C25TUV/00009186 có hai dòng ghi chú điều chỉnh tên hàng, đều trống tiền. 1C24TUV/00000063 có diễn giải hàng tặng không thu tiền. Đây là mô tả có sẵn, chưa đủ căn cứ để điền tổng thanh toán bằng 0 hoặc kết luận nội dung đúng.', '',
        '## Tra cứu 39 hóa đơn có số tham chiếu trong diễn giải', '',
        'Đọc các cột Số HĐ nhắc trong diễn giải và Ứng viên theo số trong KET_QUA_RA_SOAT_321.csv. Số tham chiếu được nhận diện từ cụm “hóa/hoá đơn số”; không dùng mã đơn sàn. Ứng viên gồm mọi ký hiệu trong bảng tổng có cùng số, có thể thuộc nhiều năm. Không chọn tự động ngay cả khi chỉ có một ứng viên trong bộ file.', '',
        '## File kết quả và bước tiếp', '',
        '- KET_QUA_RA_SOAT_321.csv: đủ 321 hóa đơn; giữ nguyên 21 cột cũ, thêm kết quả và căn cứ file phía cuối. Trạng thái/ghi chú cũ chưa được sửa; kết quả lần này nằm ở cột Kết quả rà soát file.',
        '- BANG_CHUNG_TU_FILE.csv: vị trí file/sheet/dòng và ba giá trị tiền của từng nguồn tìm thấy; không chứa tên người mua, địa chỉ hoặc MST.',
        '- LO_22_CAN_LAY_XML.csv: 8 lệch tiền + 14 thiếu tiền để lấy XML trước. Với mỗi hóa đơn lấy XML gốc của hóa đơn đang tra, XML tham chiếu và toàn bộ các lần điều chỉnh/thay thế, kể cả sau 06/2025.',
        '- Sau lô 22, lấy XML cho 39 hóa đơn có số tham chiếu trong diễn giải, rồi các hóa đơn thiếu liên kết còn lại.',
        '- Lưu file nhận được trong data/Bo sung/meInvoice; ghi đường dẫn tương đối. Cần kiểm tra MST người bán và ký hiệu/số/ngày trên XML trước khi dùng.',
        '- Mở CSV bằng Data > From Text/CSV, UTF-8; chọn Text cho cột ký hiệu/số hóa đơn để giữ số 0 đầu.', '',
        'Bảng Excel ưu tiên và CSV gốc không bị ghi đè. Chưa cập nhật tổng đối chiếu, chưa chốt Giữ nguyên/Xuất mới/Điều chỉnh/Thay thế, chưa import/ký/phát hành.', '',
        'Tái chạy: `python scripts/audit_priority_sources.py`. Script chỉ đọc dữ liệu gốc và ghi các file trong thư mục ra-soat-2026-10-07; không nhập lại ghi chú từ những file kết quả này.', '']
    (OUT / 'KET_QUA_KIEM_TRA.md').write_text('\n'.join(report), encoding='utf-8')
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == '__main__':
    main()
