"""Preserve ignored business evidence as deduplicated, verified Git-sized ZIP parts.

No source files are changed. Runtime dependencies, locks and inspection caches
are excluded. Restore is deliberately separate and never overwrites other data.
"""
from pathlib import Path
import hashlib, json, os, re, subprocess, sys, zipfile

ROOT=Path(__file__).resolve().parents[1]
DEST=ROOT/'data/git-evidence-20261109'
CHUNK=40*1024*1024
ALLOWED={'.xlsx','.xls','.csv','.tsv','.json','.xml','.zip','.png','.pdf','.docx','.txt','.md','.mjs','.py','.html'}
SECRET=[('private_key',rb'-----BEGIN (?:RSA |OPENSSH |EC )?PRIVATE KEY-----'),('github_token',rb'gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{50,}'),('aws_key',rb'AKIA[A-Z0-9]{16}'),('google_api_key',rb'AIza[0-9A-Za-z_-]{35}'),('slack_token',rb'xox[baprs]-[0-9A-Za-z-]{20,}')]

def git_paths(*args):
    return [ROOT/p.decode('utf-8') for p in subprocess.check_output(['git','ls-files','-z',*args],cwd=ROOT).split(b'\0') if p]

def sha(p):
    h=hashlib.sha256()
    with p.open('rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
    return h.hexdigest()

def main():
    ignored=git_paths('--others','--ignored','--exclude-standard')
    candidates=[]
    for p in ignored:
        rel=p.relative_to(ROOT)
        if rel.parts[0] not in {'data','outputs','scripts','Du lieu Amis','Du lieu Meinvoice'}:continue
        if any(x in {'node_modules','.venv','__pycache__','_build'} for x in rel.parts):continue
        if p.name.startswith('~$') or p.is_symlink() or p.suffix.lower() not in ALLOWED:continue
        if p.resolve().is_relative_to(ROOT.resolve()) and p.is_file():candidates.append(p)
    # Check newly supplied source/evidence without revealing any matched value.
    new=git_paths('--others','--exclude-standard')
    findings=[]
    for p in sorted(set(candidates+new)):
        if p.suffix.lower() not in {'.json','.xml','.txt','.md','.py','.mjs','.js','.html','.csv','.tsv'}:continue
        content=p.read_bytes()
        for category,pattern in SECRET:
            if re.search(pattern,content):findings.append({'path':str(p.relative_to(ROOT)),'category':category})
    assert not findings,json.dumps(findings)
    DEST.mkdir(parents=True,exist_ok=True)
    records=[];objects={}
    for p in candidates:
        digest=sha(p);records.append({'path':p.relative_to(ROOT).as_posix(),'size':p.stat().st_size,'sha256':digest});objects.setdefault(digest,p)
    target=DEST/'evidence.zip.tmp'
    with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED,compresslevel=6,allowZip64=True) as z:
        for digest,p in sorted(objects.items()):z.write(p,'objects/'+digest)
        z.writestr('manifest.json',json.dumps(records,ensure_ascii=False,indent=2).encode('utf-8'))
    with zipfile.ZipFile(target) as z:
        assert z.testzip() is None
        for digest in objects:
            with z.open('objects/'+digest) as f:
                h=hashlib.sha256()
                for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
                assert h.hexdigest()==digest
    digest=sha(target);parts=[]
    with target.open('rb') as f:
        i=1
        for block in iter(lambda:f.read(CHUNK),b''):
            p=DEST/f'evidence.zip.part{i:03d}';p.write_bytes(block)
            parts.append({'file':p.name,'size':len(block),'sha256':sha(p)});i+=1
    proof={'passed':True,'authorization':'User explicitly allowed pushing all work data on 2026-11-09','files':records,'unique_objects':len(objects),'source_bytes':sum(x['size'] for x in records),'zip_bytes':target.stat().st_size,'zip_sha256':digest,'parts':parts,'secret_pattern_findings':0,'exclusions':'Dependencies, Python caches, editor locks, inspection NDJSON/logs/temp build caches; originals remain local.'}
    (DEST/'manifest.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8')
    target.unlink() # Exact task-owned temporary archive, after every part and object verified.
    print(json.dumps({k:proof[k] for k in ['passed','unique_objects','source_bytes','zip_bytes','zip_sha256','parts','secret_pattern_findings']},ensure_ascii=False))

if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8');main()
