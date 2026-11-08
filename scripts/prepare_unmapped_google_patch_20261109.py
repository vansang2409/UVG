"""Create append-only value blocks, preserving existing import rows and notes."""
from pathlib import Path
import json,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109';BASE=OUT/'archive/before-unmapped-meinvoice-20261109/google-upload';DEST=OUT/'google-unmapped-patch'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def save(p,v):p.write_text(json.dumps(v,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
def main():
    DEST.mkdir(exist_ok=True);m=read(OUT/'google-upload/manifest.json');plan={'spreadsheetId':m['spreadsheetId'],'sheets':[],'chunks':[]}
    for sid in [700000101,700000102,700000104]:
        s=next(s for s in m['sheets'] if s['sheetId']==sid);new=read(OUT/f'google-upload/{sid}-matrix.json');old=read(BASE/f'{sid}-matrix.json') if sid!=700000104 else []
        assert new[:len(old)]==old
        plan['sheets'].append({**{k:s[k] for k in ['sheetId','title','rows','columns']},'old_rows':len(old)})
        for start in range(len(old),len(new),250):
            name=f'{sid}-{start}.json';save(DEST/name,{'sheetId':sid,'title':s['title'],'start':start,'rows':new[start:start+250]});plan['chunks'].append(name)
    save(DEST/'manifest.json',plan);print(json.dumps(plan,ensure_ascii=False))
if __name__=='__main__':main()
