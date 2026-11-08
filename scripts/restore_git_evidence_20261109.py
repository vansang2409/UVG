"""Verify or restore the authorized business-data snapshot into a chosen folder.

Usage: python scripts/restore_git_evidence_20261109.py --verify-only
       python scripts/restore_git_evidence_20261109.py --output D:/UVG-restored
Existing files must have identical hashes; different files are never overwritten.
"""
from pathlib import Path, PurePosixPath
import argparse, hashlib, json, shutil, tempfile, zipfile

ROOT=Path(__file__).resolve().parents[1]
BUNDLE=ROOT/'data/git-evidence-20261109'

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--verify-only',action='store_true');parser.add_argument('--output',type=Path);args=parser.parse_args()
    if not args.verify_only and args.output is None:parser.error('Choose --verify-only or an explicit --output folder')
    manifest=json.loads((BUNDLE/'manifest.json').read_text(encoding='utf-8'))
    with tempfile.TemporaryFile() as joined:
        whole=hashlib.sha256()
        for part in manifest['parts']:
            assert Path(part['file']).name==part['file']
            b=(BUNDLE/part['file']).read_bytes();assert len(b)==part['size'] and hashlib.sha256(b).hexdigest()==part['sha256'];whole.update(b);joined.write(b)
        assert whole.hexdigest()==manifest['zip_sha256'];joined.seek(0)
        with zipfile.ZipFile(joined) as z:
            records=json.loads(z.read('manifest.json'));assert records==manifest['files']
            for record in records:
                rel=PurePosixPath(record['path']);assert not rel.is_absolute() and '..' not in rel.parts and ':' not in str(rel)
                with z.open('objects/'+record['sha256']) as f:
                    h=hashlib.sha256();size=0
                    for b in iter(lambda:f.read(1024*1024),b''):h.update(b);size+=len(b)
                assert h.hexdigest()==record['sha256'] and size==record['size']
                if not args.verify_only:
                    dest=args.output.resolve()/Path(*rel.parts);assert dest.resolve().is_relative_to(args.output.resolve())
                    if dest.exists():assert hashlib.sha256(dest.read_bytes()).hexdigest()==record['sha256'],f'Existing different file: {dest}'
                    else:
                        dest.parent.mkdir(parents=True,exist_ok=True)
                        with z.open('objects/'+record['sha256']) as f, dest.open('xb') as out:shutil.copyfileobj(f,out)
    print(json.dumps({'passed':True,'files':len(manifest['files']),'zip_sha256':manifest['zip_sha256'],'restored':not args.verify_only}))

if __name__=='__main__':main()
