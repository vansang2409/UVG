"""Locate original detailed lines of every unresolved invoice in the orphan audit."""
from pathlib import Path
from collections import defaultdict
import sys,json,hashlib
from prepare_invoice_reports import rows,invoice_key
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def main():
    a=json.loads((OUT/'unmapped-meinvoice-audit.json').read_text(encoding='utf-8'))
    wanted={k for r in a['components'] if r['status']=='review' for k in r['keys']};found=defaultdict(list);headers={};manifest=[]
    for name in ['Bang_ke_chi_tiet_HD_da_su_dung_2025.xls','Bang_ke_chi_tiet_HD_da_su_dung_2025_MTT.xls']:
        p=ROOT/'data/Du lieu Meinvoice'/name;mt='_MTT' in name
        manifest.append({'file':str(p.relative_to(ROOT)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
        for rn,r in rows(p):
            if rn<=6:headers[name+':'+str(rn)]=r;continue
            key=invoice_key(r.get('F' if mt else 'D'),r.get('B'));k='/'.join(key) if key else None
            if k in wanted:found[k].append({'file':str(p.relative_to(ROOT)),'row':rn,'values':r})
    result={'headers':headers,'manifest':manifest,'found':dict(found),'missing':sorted(wanted-set(found))}
    (OUT/'unmapped-invoice-detail-review.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'wanted':len(wanted),'found':len(found),'missing':len(result['missing']),'examples':dict(list(found.items())[:2])},ensure_ascii=False))
if __name__=='__main__':main()
