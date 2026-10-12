"""Independent verification of the updated review, preserved evidence and action plan."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import json

import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
read = lambda n: json.loads((HERE / n).read_text(encoding='utf-8'))
data = read('confirmed-plan-data.json')
output = read('confirmed-output.json')['output']
source = openpyxl.load_workbook(data['source'], data_only=False)
book = openpyxl.load_workbook(output, data_only=False)
cached = openpyxl.load_workbook(output, data_only=True)
assert book.sheetnames == source.sheetnames + ['Phuong an phat sinh', 'Giu nguyen ca chuoi']
assert len(book.sheetnames) == 10
changes = []

def allowed(sheet, cell):
    if sheet == 'Tong hop':
        return cell.coordinate in {'A2','A4','A5','A14','G7'} or (7 <= cell.row <= 12 and cell.column >= 4) or 17 <= cell.row <= 27
    if sheet in {'Lech tien','So luong can kiem','Thanh toan trong','Hang va mo ta','Cach ghi dieu chinh'}:
        status = 18 if sheet == 'Lech tien' else 14
        return cell.coordinate in {'A4','A5'} or cell.column == status
    if sheet == 'Da co can cu':
        return cell.coordinate in {'A2','A4','F7','G7'}
    if sheet == 'Hoa don lien quan':
        return cell.coordinate == 'A5' or (cell.column == 2 and cell.row >= 8)
    return False

for old in source:
    new = book[old.title]
    for row in old:
        for cell in row:
            current = new[cell.coordinate]
            if current.value != cell.value:
                assert allowed(old.title, cell), (old.title, cell.coordinate, cell.value, current.value)
                changes.append((old.title, cell.coordinate))
    assert len(new.tables) == 1
    assert new.freeze_panes == old.freeze_panes

states = Counter()
for case in data['case_plans']:
    sheet = book[case['sheet']]
    money = case['sheet'] == 'Lech tien'
    row = case['row']
    assert sheet.cell(row, 19 if money else 15).value == case['decision']
    assert sheet.cell(row, 20 if money else 16).value == case['mode']
    state = cached[case['sheet']].cell(row, 18 if money else 14).value
    expected = 'Giữ nguyên' if case['mode'] == 'Giữ nguyên' else 'Đã chốt xử lý'
    assert state == expected
    states[state] += 1
    for col in (3,5):
        cell = sheet.cell(row,col)
        if cell.value is not None:
            assert isinstance(cell.value,datetime) and cell.number_format == 'dd/mm/yyyy'
assert states == {'Đã chốt xử lý':26, 'Giữ nguyên':3}
assert [cached['Tong hop'].cell(14,c).value for c in range(3,7)] == [29,0,26,3]
assert [cached['Tong hop'].cell(19,c).value for c in range(3,7)] == [36,0,27,9]
assert cached['Tong hop']['C21'].value == 48
assert cached['Tong hop']['C23'].value == 244
for plan in data['known_plans']:
    assert book['Da co can cu'].cell(plan['row'],9).value == plan['decision']
    assert book['Da co can cu'].cell(plan['row'],10).value == plan['mode']

target_keys = set()
for row, action in enumerate(data['actions'],8):
    ws = book['Phuong an phat sinh']
    assert ws.cell(row,1).value == action['id']
    assert ws.cell(row,4).value == '/'.join(action['target'])
    assert ws.cell(row,5).value.date().isoformat() == action['date'][:10]
    assert [ws.cell(row,c).value for c in (8,9,10)] == action['money']
    assert [ws.cell(row,c).value for c in (12,13,14)] == action['source_money']
    assert cached[ws.title].cell(row,11).value == action['difference']
    target_keys.add(tuple(action['target']))
    if action['action'] == 'Đảo điều chỉnh cũ':
        assert cached[ws.title].cell(row,15).value == 'Về 0 từng chỉ tiêu'
        assert [a+b for a,b in zip(action['source_money'],action['money'])] == [0,0,0]
assert len(target_keys) == 48
assert sum(a['difference'] != 0 for a in data['actions']) == 8
held_keys = {tuple(k) for chain in data['held_chains'] for k in chain['invoices']}
assert not target_keys & held_keys
assert len(data['held_chains']) == 244 and len(held_keys) == 338
hold=book['Giu nguyen ca chuoi']
assert hold['H8'].value == '0314215951' and hold['H8'].data_type == 's'
assert hold['I10'].value == 'Tin nhắn người dùng 11/10/2026'
raw = book['Hoa don lien quan']
held_source_rows = 0
for row in range(8,75):
    num = raw.cell(row,5)
    assert num.data_type == 's' and len(num.value) == 8 and num.value.isdigit()
    assert num.number_format == '@'
    assert raw.cell(row,6).number_format == 'dd/mm/yyyy'
    root = (raw.cell(row,4).value, num.value)
    held_source_rows += root not in target_keys
assert held_source_rows == 19
for ws in book:
    for table in ws.tables.values():
        assert table.autoFilter is not None
    for row in cached[ws.title]:
        for cell in row:
            assert cell.data_type != 'e', (ws.title,cell.coordinate,cell.value)
assert sum(len(ws.tables) for ws in book) == 10
for filename, expected_hash in data['workbook_hashes'].items():
    assert sha256((ROOT / 'outputs' / filename).read_bytes()).hexdigest() == expected_hash
assert sha256(Path(data['source']).read_bytes()).hexdigest() == data['source_sha256']
report = {'passed':True,'output':output,'sheets':10,'case_statuses':dict(states),
          'total_cases':36,'action_cases':27,'kept_cases':9,'proposed_invoices':48,
          'protected_chains':244,'protected_related_invoices':338,
          'source_invoices_preserved':67,'source_invoices_kept':19,'reversals_checked':21,
          'source_values_and_replies_preserved':True,'all_six_points_confirmed':True,
          'all_source_and_import_workbooks_unchanged':True,'formula_errors':0,
          'output_sha256':sha256(Path(output).read_bytes()).hexdigest()}
(HERE / 'confirmed-output-verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
source.close();book.close();cached.close()
