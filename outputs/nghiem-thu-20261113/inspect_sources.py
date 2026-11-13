from pathlib import Path
from zipfile import ZipFile
from itertools import islice
import sys, json, io
sys.path.insert(0, str(Path('D:/UVG/scripts')))
from prepare_invoice_reports import rows
sys.stdout.reconfigure(encoding='utf-8')
root=Path('D:/UVG/data/1311')
for p in sorted(root.rglob('*')):
    if not p.is_file(): continue
    print('\nSOURCE',p.name,flush=True)
    if p.suffix=='.zip':
        with ZipFile(p) as z:
            entries=[x for x in z.infolist() if not x.is_dir()]
            print('PARTS',[(x.filename,x.file_size) for x in entries],flush=True)
            if entries:
                for rn,r in islice(rows(io.BytesIO(z.read(entries[0]))),8):
                    print(rn,json.dumps(r,ensure_ascii=False),flush=True)
    else:
        for rn,r in islice(rows(p),8):
            print(rn,json.dumps(r,ensure_ascii=False),flush=True)
