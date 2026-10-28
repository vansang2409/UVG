"""Promote only independently verified files, retaining recoverable snapshots."""
from pathlib import Path
from hashlib import sha256
import json,os,shutil
HERE=Path(__file__).resolve().parent;ROOT=HERE.parents[2]
p=json.loads((HERE/'plan.json').read_text(encoding='utf8'))
report=json.loads((HERE/'staged-verification.json').read_text(encoding='utf8'))
sha=lambda path:sha256(path.read_bytes()).hexdigest()
assert report['passed'] and report['counts']==p['counts'] and len(report['files'])==4
expected={f['name']:f['sha256'] for f in report['files']}
for f in p['files']:
    destination=(ROOT/f['destination']).resolve()
    assert destination.is_relative_to((ROOT/'outputs').resolve())
    assert sha(destination)==f['before_sha256'],('current workbook changed',f['name'])
    assert sha(HERE/'before'/f['name'])==f['before_sha256']
    assert sha(HERE/'staged'/f['name'])==expected[f['name']]
replaced=[]
try:
    for f in p['files']:
        destination=ROOT/f['destination'];temp=destination.with_suffix('.xlsx.updating')
        assert not temp.exists()
        shutil.copyfile(HERE/'staged'/f['name'],temp)
        os.replace(temp,destination);replaced.append(f)
        assert sha(destination)==expected[f['name']]
except Exception:
    for f in replaced:shutil.copyfile(HERE/'before'/f['name'],ROOT/f['destination'])
    raise
rounding=json.loads((HERE/'rounding-review.json').read_text(encoding='utf8'))
for r in rounding['lines']:
    action=next(a for a in p['actions'] if a['seq']==r['sequence'])
    line=action['lines'][r['target_line']-1]
    assert r['import_row']==action['start']+r['target_line']-1
    assert (line['qty'],line['price'],line['goods'])==(r['qty'],r['positive_unit_price'],r['issued_line_amount_reversed'])
report.update(mode='final',verification_method='Independent saved-file checks followed by SHA-256 equality for every promoted workbook',approved_plan_date=p['date'],source_rounding_review=rounding)
(HERE/'final-verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
print('Updated four workbooks. All final SHA-256 hashes match independently verified files.')
print(json.dumps(report['counts'],ensure_ascii=False))
