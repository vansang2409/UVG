"""Reconcile the newly issued returns and duplicate replacements, read-only."""
from collections import Counter, defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path, PurePosixPath
from zipfile import ZipFile
import hashlib,json,sys

from openpyxl import load_workbook
from prepare_invoice_reports import rows,invoice_key,invoice_day,amount

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/post-final-issue-20261101'
SOURCE=ROOT/'data/Meinv_SAU_DIEU_CHINH_THAY_THE_20261101'
PREVIOUS=ROOT/'scripts/data/new-sales-issued-20261101'
ARCHIVE=Path(r'C:\Users\thuong ngo\Downloads\Bang_ke_hoa_don_da_su_dung_1793524594152.zip')
LINK_FILES={
    'adjustment':Path(r'C:\Users\thuong ngo\Downloads\Danh_sach_hoa_don_dieu_chinh.xlsx'),
    'replacement':Path(r'C:\Users\thuong ngo\Downloads\Danh_sach_hoa_don_thay_the.xlsx'),
}
PLAN=Path(r'C:\Users\thuong ngo\Downloads\CHECK HOA DON XUAT MOI 2024 31.10.2026\IMPORT_DIEU_CHINH_163_PHIEU_TRA_HANG_2024_MOI_NHAT.xlsx')


def native(v):
    if isinstance(v,Decimal):return int(v) if v==int(v) else str(v)
    if isinstance(v,(date,Path)):return str(v)
    raise TypeError(type(v))


def dump(name,value):
    (WORK/name).write_text(json.dumps(value,ensure_ascii=False,indent=2,default=native)+'\n',encoding='utf-8')


def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def text(v):return '' if v is None else str(v).strip()
def key(series,number):return '/'.join(invoice_key(series,number))


def sums(records):
    return [sum((r['money'][i] for r in records if r['money'][i] is not None),Decimal(0)) for i in range(3)]


def read_summary(path):
    headers={};records={};indices=[];footer=[];period=None
    for rn,r in rows(path):
        if rn==3:period=r.get('A')
        if rn==5:headers={v:k for k,v in r.items() if v}
        if rn<6:continue
        get=lambda label:r.get(headers.get(label))
        if amount(r.get('A')) is None:
            if r.get('A')=='Tổng cộng':footer=[amount(get(c)) for c in ['Doanh số bán chưa thuế','Tiền thuế GTGT','Tổng tiền']]
            continue
        k=key(get('Ký hiệu'),get('Số hóa đơn'))
        assert k not in records
        indices.append(int(amount(r['A'])))
        records[k]={'key':k,'date':invoice_day(get('Ngày hóa đơn')),'order':text(get('Mã đơn hàng')),
                    'money':[amount(get(c)) for c in ['Doanh số bán chưa thuế','Tiền thuế GTGT','Tổng tiền']],
                    'cqt':text(get('Mã của CQT')),'cqt_status':text(get('Trạng thái gửi CQT')),'status':text(get('Trạng thái hóa đơn')),
                    'source_file':path.relative_to(ROOT).as_posix(),'source_row':rn}
    assert indices==list(range(1,len(indices)+1))
    assert sums(records.values())==footer,(path.name,'Footer differs',sums(records.values()),footer)
    return records,{'file':path.relative_to(ROOT).as_posix(),'invoices':len(records),'period':period,'money':footer}


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    WORK.mkdir(parents=True,exist_ok=True);SOURCE.mkdir(parents=True,exist_ok=True)
    manifest=[];summary_paths=[];link_paths={}
    def save(name,raw,kind):
        path=(SOURCE/name).resolve()
        assert path.parent==SOURCE.resolve()
        if path.exists():assert path.read_bytes()==raw,('Existing source changed',name)
        else:path.write_bytes(raw)
        manifest.append({'file':path.relative_to(ROOT).as_posix(),'kind':kind,'bytes':len(raw),'sha256':sha(path)})
        return path
    with ZipFile(ARCHIVE) as z:
        assert z.testzip() is None
        for e in z.infolist():
            if e.is_dir():continue
            name=PurePosixPath(e.filename.replace('\\','/'))
            assert len(name.parts)==1 and name.suffix.lower() in {'.xls','.xlsx'}
            summary_paths.append(save(name.name,z.read(e),'summary'))
    for kind,path in LINK_FILES.items():link_paths[kind]=save(path.name,path.read_bytes(),kind)
    actual={};file_controls=[]
    for path in summary_paths:
        found,control=read_summary(path)
        assert not actual.keys()&found.keys(),'Duplicate invoice across summary parts'
        actual.update(found);file_controls.append(control)
    links={}
    for kind,path in link_paths.items():
        items=[];indices=[]
        for rn,r in rows(path):
            if rn<8 or amount(r.get('A')) is None:continue
            indices.append(int(amount(r['A'])))
            old=key(r.get('B'),r.get('D') if kind=='adjustment' else r.get('E'))
            new=key(r.get('I'),r.get('K') if kind=='adjustment' else r.get('J'))
            items.append({'old':old,'new':new,'old_date':invoice_day(r.get('C')),
                          'new_date':invoice_day(r.get('J') if kind=='adjustment' else r.get('K')),
                          'old_cqt':r.get('E') if kind=='adjustment' else r.get('D'),
                          'reason':text(r.get('L') if kind=='adjustment' else r.get('H')),
                          'source_file':path.relative_to(ROOT).as_posix(),'source_row':rn})
        assert indices==list(range(1,len(indices)+1))
        assert len({x['old'] for x in items})==len(items) and len({x['new'] for x in items})==len(items)
        links[kind]=items
    previous=json.loads((PREVIOUS/'summary.json').read_text(encoding='utf-8'))
    resolved=json.loads((PREVIOUS/'resolved-coverage.json').read_text(encoding='utf-8'))
    before={}
    for e in previous['source_manifest']:
        recs,_=read_summary(ROOT/e['file']);before.update(recs)
    # Preserve the exact current import file used by the user as the plan source.
    plan_hash=sha(PLAN)
    wb=load_workbook(PLAN,read_only=True,data_only=True)
    sheet=wb['Import dieu chinh'];plans={}
    expected_doc=json.loads((ROOT/'scripts/data/return-adjustment-issued-20261101/prepared.json').read_text(encoding='utf-8'))
    doc_by_old={a['invoice']:a['doc'] for a in expected_doc['audit']}
    current=None
    for rn,r in enumerate(sheet.iter_rows(min_row=10,max_row=228,max_col=26,values_only=True),10):
        if current is None or r[0]!=current['sequence']:
            old=key(r[9],r[10])
            assert old not in plans
            current={'sequence':r[0],'old':old,'doc':doc_by_old[old],'order':text(r[24]),'reason':r[13],
                     'money':[Decimal(0),Decimal(0),Decimal(0)],'lines':0,'rows':[]}
            plans[old]=current
        assert key(r[9],r[10])==current['old'] and r[13]==current['reason'] and text(r[24])==current['order']
        current['money'][0]+=Decimal(str(r[18]))
        for i,c in [(1,21),(2,22)]:
            if r[c] is not None:current['money'][i]+=Decimal(str(r[c]))
        current['lines']+=1;current['rows'].append(rn)
    wb.close()
    assert len(plans)==163 and sum(p['lines'] for p in plans.values())==219
    assert sums(plans.values())==[Decimal(-225865926),Decimal(-18069275),Decimal(-243935201)]
    adjustments=[];replacement_results=[];issues=[]
    if {x['old'] for x in links['adjustment']}!=set(plans):issues.append({'type':'Adjustment parent set differs'})
    adjustment_by_old={x['old']:x for x in links['adjustment']}
    for old,p in plans.items():
        link=adjustment_by_old.get(old)
        if link is None:continue
        new=actual.get(link['new']);parent=actual.get(old)
        if new is None or parent is None:
            issues.append({'type':'Adjustment invoice missing from summary','old':old,'new':link['new']});continue
        checks={'money':new['money']==p['money'],'order':new['order']==parent['order']==p['order'],
                'date':new['date']==link['new_date']==date(2026,11,1),'parent_date':parent['date']==link['old_date'],
                'parent_cqt':parent['cqt']==link['old_cqt'],'reason':link['reason']==p['reason'],'kind':new['status']=='Hoá đơn điều chỉnh'}
        if not all(checks.values()):issues.append({'type':'Adjustment check failed','old':old,'new':new['key'],'checks':checks})
        adjustments.append({'doc':p['doc'],'order':parent['order'],'sequence':p['sequence'],'old':old,'new':new['key'],
                            'new_date':new['date'],'expected':p['money'],'actual':new['money'],'delta':[None if new['money'][i] is None else new['money'][i]-p['money'][i] for i in range(3)],
                            'parent_status':parent['status'],'cqt_status':new['cqt_status'],'cqt':new['cqt'],
                            'reason':link['reason'],'checks':checks,'source_file':new['source_file'],'source_row':new['source_row']})
    pairs={p['invoices'][1]['key']:p for p in resolved['duplicate_pairs']}
    replacement_by_old={x['old']:x for x in links['replacement']}
    if set(replacement_by_old)!=set(pairs):issues.append({'type':'Replacement parent set differs'})
    for old,p in sorted(pairs.items()):
        link=replacement_by_old.get(old)
        if link is None:continue
        new=actual.get(link['new']);parent=actual.get(old);retained=actual.get(p['invoices'][0]['key'])
        if new is None or parent is None or retained is None:
            issues.append({'type':'Replacement invoice missing from summary','old':old});continue
        checks={'new_goods_vat_zero':new['money'][:2]==[Decimal(0),Decimal(0)],'old_replaced_status':parent['status']=='Hóa đơn đã bị thay thế',
                'new_replacement_status':new['status']=='Hoá đơn thay thế','cqt_valid':new['cqt_status']=='HĐ hợp lệ',
                'old_date':parent['date']==link['old_date'],'new_date':new['date']==link['new_date']==date(2026,11,1),
                'old_cqt':parent['cqt']==link['old_cqt'],'old_amount_unchanged':parent['money']==[Decimal(v) for v in p['expected']],
                'retained_unchanged':retained['money']==[Decimal(v) for v in p['expected']] and retained['status']=='Hoá đơn mới'}
        if not all(checks.values()):issues.append({'type':'Replacement check failed','old':old,'new':new['key'],'checks':checks})
        replacement_results.append({'doc':p['doc'],'retained':retained['key'],'old':old,'new':new['key'],'new_date':new['date'],
                                    'before':parent['money'],'after':new['money'],'parent_status':parent['status'],'cqt_status':new['cqt_status'],'cqt':new['cqt'],
                                    'checks':checks,'source_file':new['source_file'],'source_row':new['source_row']})
    adjustments.sort(key=lambda a:int(a['new'].split('/')[1]))
    new_keys={x['new'] for items in links.values() for x in items}
    assert len(new_keys)==2163
    added=set(actual)-set(before)
    if added!=new_keys or not set(before)<=set(actual):issues.append({'type':'Unexpected new or missing invoices','extra':sorted(added-new_keys),'missing':sorted(new_keys-added)})
    old_value_changes=[k for k in before if k not in actual or any(actual[k][f]!=before[k][f] for f in ['date','order','money','cqt'])]
    if old_value_changes:issues.append({'type':'Old invoice source values changed','keys':old_value_changes})
    # One original copy per planned 2024 sale, after excluding the replaced copy.
    chosen=[]
    prior_detail=json.loads((PREVIOUS/'detail-summary.json').read_text(encoding='utf-8'))
    payment_evidence={p['invoice']:p['detail_rows'][0]['payment'] for p in prior_detail['blank_payment_evidence']}
    for p in resolved['coverage']+resolved['blank_payments']:
        surviving=[x['key'] for x in p['invoices'] if x['key'] not in pairs]
        assert len(surviving)==1
        rec=actual[surviving[0]]
        money=rec['money'].copy()
        if money[2] is None and rec['key'] in payment_evidence:
            assert money[:2]==[Decimal(0),Decimal(0)] and payment_evidence[rec['key']]==0
            money[2]=Decimal(0)
        chosen.append({'key':rec['key'],'money':money})
    base_money=sums(chosen)
    assert len(chosen)==34739 and base_money==[Decimal(v) for v in previous['plan']['money']]
    adjustment_money=[sum(a['actual'][i] for a in adjustments) for i in range(3)]
    unknown_replacement_payment=sum(r['after'][2] is None for r in replacement_results)
    net_goods=base_money[0]+adjustment_money[0]+sum(r['after'][0] for r in replacement_results)
    net_vat=base_money[1]+adjustment_money[1]+sum(r['after'][1] for r in replacement_results)
    payment_known=base_money[2]+adjustment_money[2]+sum(r['after'][2] for r in replacement_results if r['after'][2] is not None)
    pending=[a for a in adjustments if a['cqt_status']!='HĐ hợp lệ']
    summary_result={'source_manifest':manifest,'archive_sha256':sha(ARCHIVE),'plan_file':str(PLAN),'plan_sha256':plan_hash,
        'file_controls':file_controls,'export_invoices':len(actual),'export_money':sums(actual.values()),'link_counts':{k:len(v) for k,v in links.items()},
        'new_invoices':len(added),'prior_invoice_values_unchanged':not old_value_changes,
        'adjustments':{'invoices':len(adjustments),'money':adjustment_money,'all_reference_money_reason_checks_passed':all(all(x['checks'].values()) for x in adjustments),
                       'cqt_statuses':dict(Counter(a['cqt_status'] for a in adjustments)), 'new_min':min(a['new'] for a in adjustments),'new_max':max(a['new'] for a in adjustments),
                       'original_statuses':dict(Counter(a['parent_status'] for a in adjustments)), 'pending_min':min(a['new'] for a in pending) if pending else None,'pending_max':max(a['new'] for a in pending) if pending else None},
        'replacements':{'invoices':len(replacement_results),'old_min':min(pairs),'old_max':max(pairs),'new_min':min(r['new'] for r in replacement_results),'new_max':max(r['new'] for r in replacement_results),
                        'all_checks_passed':all(all(r['checks'].values()) for r in replacement_results),'old_amount_removed':sums([{'money':r['before']} for r in replacement_results]),
                        'new_goods_vat':[sum(r['after'][i] for r in replacement_results) for i in range(2)],'cqt_statuses':dict(Counter(r['cqt_status'] for r in replacement_results)),
                        'new_payment_blank_count':unknown_replacement_payment},
        'revenue_scope':'Only the 2024 new-sales import after returns and duplicate replacements; excludes 273 unrelated invoices and historic invoices retained separately.',
        'base_sales_money':base_money,'new_sales_net_goods':net_goods,'new_sales_net_vat':net_vat,
        'new_sales_net_payment':None if unknown_replacement_payment else payment_known,'new_sales_net_payment_known_part':payment_known,
        'new_sales_net_payment_if_replacement_totals_zero':payment_known,'unknown_replacement_payment_invoices':unknown_replacement_payment,
        'issues':issues,'missing_latest_detail_export':True,
        'limits':['63 invoice status values may still be awaiting update/CQT response; not automatically treated as rejected.',
                  'No current detailed export: cannot verify new product-level details, note flags or the 2000 blank replacement total-payment fields.',
                  'Blank payment is not changed to zero; the previous two promotional zero invoices use their own existing detailed-source evidence.'],
        'status_help':'https://helpv4.meinvoice.vn/kb/mot-so-van-de-lien-quan-den-hoa-don-khoi-tao-tu-may-tinh-tien-theo-nghi-dinh-254-2026-nd-cp/'}
    dump('summary.json',summary_result);dump('adjustments.json',adjustments);dump('replacements.json',replacement_results);dump('pending-cqt.json',pending)
    assert all(sha(ROOT/e['file'])==e['sha256'] for e in manifest) and sha(PLAN)==plan_hash
    print(json.dumps({k:v for k,v in summary_result.items() if k not in ['source_manifest','file_controls']},ensure_ascii=False,indent=2,default=native))


if __name__=='__main__':main()
