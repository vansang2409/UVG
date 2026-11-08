"""Consolidate monthly sales files into one import and update dependent invoice sequences."""
from pathlib import Path
from collections import defaultdict
import csv,json,copy,sys
from import_2025_buyer_fields import apply_buyer_fields

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/import-2025-20261109'
TARGET='IMPORT_HD_MOI_2025.xlsx'
sys.stdout.reconfigure(encoding='utf-8')

def main():
    path=OUT/'authoring.json';p=json.loads(path.read_text(encoding='utf-8'))
    old=[o for o in p['outputs'] if o['kind']=='new']
    if len(old)==1 and old[0]['file']==TARGET:
        if (OUT/'buyer-field-confirmation.json').exists():
            stats=apply_buyer_fields(p,OUT,ROOT)
            path.write_text(json.dumps(p,ensure_ascii=False),encoding='utf-8')
            print(json.dumps({'already_consolidated':True,'buyer_fields':stats},ensure_ascii=False))
        else:print('Already consolidated')
        return
    assert len(old)==7
    backup=OUT/'source/authoring-monthly.json'
    if not backup.exists():backup.write_text(json.dumps(p,ensure_ascii=False),encoding='utf-8')
    records=[];seqs={};oldmap={}
    for o in old:
        for row in o['rows']:
            code=row[18]
            if code not in seqs:
                seqs[code]=len(seqs)+1;oldmap[code]=(o['file'],row[0])
            assert oldmap[code]==(o['file'],row[0])
            records.append([seqs[code],*row[1:]])
    assert len(seqs)==p['summary']['new_invoices']
    assert len(records)==p['summary']['new_lines']
    history=OUT/'invoice-sequences.json'
    previous=json.loads(history.read_text(encoding='utf-8')) if history.exists() else {}
    assert set(previous)<=set(seqs),'Previously assigned invoices disappeared; do not silently resequence'
    assert sorted(previous.values())==list(range(1,len(previous)+1))
    stable=dict(previous)
    for code in seqs:
        if code not in stable:stable[code]=len(stable)+1
    seqs=stable
    for row in records:row[0]=seqs[row[18]]
    records.sort(key=lambda row:row[0])
    history.write_text(json.dumps(seqs,ensure_ascii=False),encoding='utf-8')
    merged=copy.deepcopy(old[0]);merged.update(file=TARGET,rows=records,invoices=len(seqs))
    count_text=f'{len(seqs):,}'.replace(',','.')
    merged['notes'][0]=f'Gộp toàn bộ bán mới năm 2025 trong một sheet; STT hóa đơn liên tục 1–{count_text}. Cột Ngày chứng từ AMIS giữ thông tin kỳ gốc.'
    p['outputs']=[merged]+[o for o in p['outputs'] if o['kind']!='new']
    report=next(o for o in p['outputs'] if o['kind']=='report')
    updated=0
    reference_tables=report['sheets']+[o for o in p['outputs'] if o['kind']=='return_adjustment']
    for sh in reference_tables:
        h=sh['headers']
        if 'File import bán mới' not in h:continue
        fi=h.index('File import bán mới');si=h.index('STT HĐ bán mới');ci=h.index('Chứng từ bán gốc')
        for row in sh['rows']:
            if not row[fi]:continue
            code=row[ci];assert (row[fi],row[si])==oldmap[code]
            row[fi]=TARGET;row[si]=seqs[code];updated+=1
    for filename in ['LIEN_KET_BTL_BHH_2025.csv']:
        fpath=OUT/filename
        with fpath.open(encoding='utf-8-sig',newline='') as f:
            r=csv.reader(f);h=next(r);rs=list(r)
        fi=h.index('File import bán mới');si=h.index('STT HĐ bán mới');ci=h.index('Chứng từ bán gốc')
        for row in rs:
            if not row[fi]:continue
            code=row[ci];assert (row[fi],int(row[si]))==oldmap[code]
            row[fi]=TARGET;row[si]=seqs[code]
        with fpath.open('w',encoding='utf-8-sig',newline='') as f:
            w=csv.writer(f);w.writerow(h);w.writerows(rs)
    grouped=defaultdict(list)
    for i,row in enumerate(records,10):grouped[row[18]].append((i,row))
    with (OUT/'LIEN_KET_IMPORT_MOI.csv').open(encoding='utf-8-sig',newline='') as f:
        reader=csv.reader(f);h=next(reader);roots={r[2]:r[5] for r in reader}
    with (OUT/'LIEN_KET_IMPORT_MOI.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(h)
        for code,ls in grouped.items():w.writerow([TARGET,seqs[code],code,ls[0][0],len(ls),roots[code]])
    p['summary']['delivery_files']=len(p['outputs'])
    p['summary']['sales_import_file']=TARGET
    if (OUT/'buyer-field-confirmation.json').exists():apply_buyer_fields(p,OUT,ROOT)
    path.write_text(json.dumps(p,ensure_ascii=False),encoding='utf-8')
    (OUT/'final-summary.json').write_text(json.dumps(p['summary'],ensure_ascii=False),encoding='utf-8')
    (OUT/'consolidation.json').write_text(json.dumps({'previous_sales_files':[o['file'] for o in old],'new_file':TARGET,'invoices':len(seqs),'rows':len(records),'updated_report_references':updated,'money_unchanged':True},ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'files':len(p['outputs']),'invoices':len(seqs),'rows':len(records),'updated_report_references':updated}))

if __name__=='__main__':main()
