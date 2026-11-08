"""Package only the five independently verified Excel deliverables."""
from pathlib import Path
import hashlib,json,zipfile,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def main():
    v=json.loads((OUT/'verification.json').read_text(encoding='utf-8'));assert v['passed']
    files=sorted(p for p in (OUT/'deliverables').glob('*.xlsx') if not p.name.startswith('~$'))
    expected={r['file']:r['sha256'] for r in v['exported_files']}
    assert len(files)==5 and {p.name for p in files}==set(expected)
    for p in files:assert hashlib.sha256(p.read_bytes()).hexdigest()==expected[p.name]
    target=OUT/'BO_IMPORT_HD_2025_20261109.zip';tmp=target.with_suffix('.zip.tmp')
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as z:
        for p in files:z.write(p,p.name)
    with zipfile.ZipFile(tmp) as z:
        assert z.testzip() is None and set(z.namelist())==set(expected)
        for name in z.namelist():assert hashlib.sha256(z.read(name)).hexdigest()==expected[name]
    tmp.replace(target)
    result={'passed':True,'file':target.name,'members':expected,'sha256':hashlib.sha256(target.read_bytes()).hexdigest(),'size':target.stat().st_size}
    (OUT/'zip-verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
