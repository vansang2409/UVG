"""Reuse the repository's verified invoice-chain reader; do not edit sources."""
from pathlib import Path
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path('D:/UVG')
OUT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT/'scripts'))
original = ROOT/'scripts/classify_2025_20261105.py'
prefix = original.read_text(encoding='utf-8').split("amis=read(SRC/'plan25AMIS.json')", 1)[0]
assert 'def chain(k)' in prefix and 'amis=read' not in prefix
# Only the index-cache location changes; all invoice input paths stay as authored.
prefix = prefix.replace("OUT=ROOT/'outputs/phan-loai-2025-20261105'", "OUT=Path("+repr(str(OUT))+ ")")
namespace = {'__file__': str(original), '__name__': 'invoice_chain_read_only'}
exec(compile(prefix, str(original), 'exec'), namespace)
print(json.dumps({'stage':'index_loaded','invoices':len(namespace['ledger']),'edges':sum(len(v) for v in namespace['children'].values())},ensure_ascii=False),flush=True)

def normkey(symbol, number):
    return namespace['kval'](symbol, number)

rows = json.loads((OUT/'hd-cty-before-chain.json').read_text(encoding='utf-8'))['values']
local = []
for i, r in enumerate(rows, 1):
    if i < 5 or len(r)<4 or not r[3]:
        continue
    k = normkey(str(r[1] or '')+str(r[2] or ''), r[3])
    roots = sorted(namespace['roots'](k)) if k else []
    ch = namespace['chain'](roots[0]) if len(roots)==1 else None
    local.append({'row':i,'invoice':k,'roots':roots,'chain':ch,'invoice_exists':k in namespace['ledger']})
(OUT/'full-chain-resolution.json').write_text(json.dumps(local,ensure_ascii=False,indent=2),encoding='utf-8')
counts = {}
for r in local:
    status = 'missing_invoice' if not r['invoice_exists'] else 'ambiguous_roots' if len(r['roots'])!=1 else 'missing_money' if any(x is None for x in r['chain']['balance']) else 'has_issues' if r['chain']['issues'] else 'resolved'
    counts[status] = counts.get(status,0)+1
print(json.dumps({'stage':'resolution','counts':counts,'changed':[{ 'row':r['row'],'invoice':r['invoice'],'roots':r['roots'],'balance':r['chain']['balance'],'keys':r['chain']['keys'],'issues':r['chain']['issues']} for r in local if r['chain'] and (len(r['chain']['keys'])>1 or r['chain']['issues'])]},ensure_ascii=False),flush=True)
