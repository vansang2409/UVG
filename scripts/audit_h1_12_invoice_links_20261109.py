"""Read all local meInvoice exports/XML for the twelve exact marketplace order IDs."""
from pathlib import Path
from zipfile import ZipFile, BadZipFile
from io import BytesIO
import json, hashlib, re

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs/import-2025-20261109'

def main():
    p = json.loads((OUT/'prepared.json').read_text(encoding='utf-8'))
    codes = ['BHHS2025010077','BHHS2025010076','BHHS2025010073','BHHS2025010072','BHHS2025010071','BHHS2025010068','BHHS2025010067','BHHS2025010066','BHHS2025010065','BHHS2025010064','BHHS2025010062','BHHS2025010060']
    orders = {p['docs'][code]['order']: code for code in codes}
    assert len(orders) == 12
    pattern = re.compile(b'|'.join(re.escape(k.encode()) for k in orders))
    controls, hits, unreadable = [], [], []
    for path in sorted((ROOT/'data').rglob('*')):
        if not path.is_file() or 'meinv' not in str(path).lower() or path.suffix.lower() not in ['.xls','.xlsx','.xml'] or path.name.startswith('~$'):
            continue
        data = path.read_bytes()
        name = path.relative_to(ROOT).as_posix()
        controls.append({'file': name, 'bytes':len(data), 'sha256':hashlib.sha256(data).hexdigest()})
        if path.suffix.lower() == '.xml':
            streams = [(name, data)]
        else:
            try:
                with ZipFile(BytesIO(data)) as z:
                    streams = [(member,z.read(member)) for member in z.namelist() if member == 'xl/sharedStrings.xml' or member.startswith('xl/worksheets/') and member.endswith('.xml')]
            except BadZipFile:
                unreadable.append(name)
                continue
        for member, raw in streams:
            for order in sorted(set(pattern.findall(raw.upper()))):
                hits.append({'file':name,'member':member,'order':order.decode(),'sale':orders[order.decode()]})
    result = {'exact_order_ids':orders, 'files_checked':len(controls), 'controls':controls, 'hits':hits, 'unreadable_files':unreadable, 'scope':'All local .xls/.xlsx/.xml under data paths containing meinv; no invoice inferred from amount or number alone'}
    (OUT/'h1-12-invoice-source-audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    ui = json.loads((OUT/'source/h1-12-amis-live.json').read_text(encoding='utf-8'))
    assert {r['doc'] for r in ui['records']} == set(codes)
    assert not hits and not unreadable, 'Investigate matches or unreadable sources before classifying links as unresolved'
    original = json.loads((OUT/'source/h1-12-google-evidence.json').read_text(encoding='utf-8'))
    original_rows = {r[2]:r for r in original['original']['values']}
    cache = json.loads((OUT/'source-index.json').read_text(encoding='utf-8'))
    reviewed = []
    for rec in ui['records']:
        code=rec['doc']; d=p['docs'][code]; old=original_rows[code]
        assert rec['symbol'] is None and rec['invoice_number'] is None and rec['status']=='Đã lập hóa đơn Chưa phát hành'
        assert d['date'] == rec['date'] == '2025-06-27' and d['order'] == old[8]
        assert not cache['raw_exact'].get(d['order'])
        key=old[10]+'/'+str(old[9]).zfill(8); inv=cache['ledger'][key]
        reviewed.append({'doc':code,'order':d['order'],'date':d['date'],'goods':d['goods'],'vat':d['vat'],'pay':d['pay'],'snapshot_invoice':key,'snapshot_invoice_date':inv['date'],'snapshot_invoice_pay':inv['resolved_money'][2],'snapshot_invoice_source':inv['source'],'amis_status':rec['status'],'amis_symbol':None,'amis_invoice_number':None,'amis_url':rec['url'],'confirmed_invoice':None,'conclusion':'T1–T6 tạo mới; HĐ thực tế chưa xác định. Không dùng liên kết Sheet chưa xác minh để xử lý HĐ tháng 7 đang giữ.'})
    review={'checked_on':ui['checked_on'],'records':reviewed,'source_files_checked':len(controls),'exact_order_matches':0,'limitation':ui['limitation'],'policy':'Tháng 1–6 tạo mới toàn bộ trừ danh sách công ty; phần xác định/xử lý HĐ cũ theo dõi riêng, không tự gán hóa đơn.'}
    (OUT/'h1-12-invoice-link-review.json').write_text(json.dumps(review,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'files_checked':len(controls),'hits':hits,'unreadable_files':unreadable},ensure_ascii=False))

if __name__ == '__main__':
    main()
