"""Replace the three authorized reports only after independent staged checks pass."""
from pathlib import Path
import hashlib,json,os,shutil
HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[2]
plan=json.loads((HERE/'edit-plan.json').read_text(encoding='utf8'))
verification=json.loads((HERE/'staged-verification.json').read_text(encoding='utf8'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
assert verification['passed'] and len(verification['files'])==3
assert verification['new_invoices']==48 and verification['formula_errors']==0
verified={f['name']:f for f in verification['files']}
for f in plan['files']:
    assert sha(ROOT/'outputs'/f['name'])==f['sha256'],('current workbook changed',f['name'])
    assert sha(HERE/'before'/f['name'])==f['sha256']
    assert sha(HERE/'staged'/f['name'])==verified[f['name']]['sha256']
replaced=[]
try:
    for f in plan['files']:
        dest=ROOT/'outputs'/f['name'];temp=dest.with_suffix('.xlsx.updating')
        assert temp.parent.resolve()==dest.parent.resolve() and not temp.exists()
        shutil.copyfile(HERE/'staged'/f['name'],temp)
        os.replace(temp,dest)
        assert sha(dest)==verified[f['name']]['sha256']
        replaced.append(f)
except Exception:
    for f in replaced:shutil.copyfile(HERE/'before'/f['name'],ROOT/'outputs'/f['name'])
    raise
final_report={**verification,'mode':'final','verification_method':'Independent staged content checks, followed by SHA-256 equality for all promoted files','software_imported':False,'signed':False,'issued':False}
final_report['files']=[{**verified[f['name']],'path':(Path('outputs')/f['name']).as_posix(),'backup_path':(HERE/'before'/f['name']).relative_to(ROOT).as_posix(),'backup_sha256':f['sha256']} for f in plan['files']]
(HERE/'final-verification.json').write_text(json.dumps(final_report,ensure_ascii=False,indent=2),encoding='utf8')
print('Updated all three reports; recoverable originals retained in before/.')
