"""Stage narrow, guarded name-field writes and the 14 approved new rows."""
from pathlib import Path
import json,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def save(p,x):p.write_text(json.dumps(x,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
def cell(x):return {} if x is None else {'userEnteredValue':{'numberValue':x} if isinstance(x,(int,float)) else {'stringValue':x}}
def main():
    before=OUT/'archive/before-service-quantity-fix-20261109/google-upload';current=OUT/'google-upload';dest=OUT/'google-buyer-service-patch';dest.mkdir(exist_ok=True)
    manifest=read(current/'manifest.json');old_manifest=read(before/'manifest.json');logs=[];changed=0
    for tab in manifest['sheets']:
        sid=tab['sheetId'];old=read(before/f'{sid}-matrix.json');new=read(current/f'{sid}-matrix.json')
        assert len(new[0])==len(old[0]) and old[8]==new[8]
        for i,row in enumerate(old):
            for j,value in enumerate(row):
                if value!=new[i][j]:
                    assert (i>=9 and j in [2,5]) or (sid==0 and i==1 and j==0),(sid,i,j)
                    changed+=1
        if sid==0:assert len(old)==35011 and len(new)==35025 and len(new)-len(old)==14
        else:assert len(old)==len(new)
        for start in range(9,len(old),1000):
            end=min(start+1000,len(old));requests=[]
            for col in [2,5]:
                if not any(old[i][col]!=new[i][col] for i in range(start,end)):continue
                requests.append({'updateCells':{'range':{'sheetId':sid,'startRowIndex':start,'endRowIndex':end,'startColumnIndex':col,'endColumnIndex':col+1},'rows':[{'values':[cell(new[i][col])]} for i in range(start,end)],'fields':'userEnteredValue'}})
            if not requests:continue
            filename=f'{sid}-{start}.json'
            save(dest/filename,{'sheetId':sid,'title':tab['title'],'start':start,'end':end,'expected':[[old[i][j] for j in [0,2,3,4,5]] for i in range(start,end)],'requests':requests})
            logs.append({'file':filename,'sheetId':sid,'start':start,'end':end})
    old=read(before/'0-matrix.json');new=read(current/'0-matrix.json')
    save(dest/'append.json',{'old_note':old[1][0],'new_note':new[1][0],'last_old_row':old[-1],'rows':new[len(old):]})
    result={'spreadsheetId':manifest['spreadsheetId'],'batches':logs,'old_rows':len(old),'new_rows':len(new),'columns':23,'changed_existing_cells':changed,'append_rows':14,'source_baseline':str(before.relative_to(ROOT))}
    save(dest/'manifest.json',result);print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
