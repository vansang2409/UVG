"""Read-only reconciliation of the latest issued detail with returns/replacements."""
from collections import Counter, defaultdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path, PurePosixPath
from zipfile import ZipFile
import hashlib, json, sys

from openpyxl import load_workbook
from prepare_invoice_reports import rows, amount, invoice_day, invoice_key
from audit_final_issue_20261101 import read_summary

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/post-final-detail-20261101'
PRIOR=ROOT/'scripts/data/post-final-issue-20261101'
SOURCE=ROOT/'data/Meinv_SAU_DIEU_CHINH_THAY_THE_20261101/CHI_TIET'
ARCHIVE=Path(r'C:\Users\thuong ngo\Downloads\Bang_ke_chi_tiet_hoa_don_da_su_dung_20261101_161720.zip')

def text(v):return '' if v is None else str(v).strip()
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def native(v):
    if isinstance(v,Decimal):return int(v) if v==int(v) else str(v)
    if isinstance(v,(date,Path)):return str(v)
    raise TypeError(type(v))
def dump(name,v):(WORK/name).write_text(json.dumps(v,ensure_ascii=False,indent=2,default=native)+'\n',encoding='utf-8')
def key(a,b):return '/'.join(invoice_key(a,b))
def complete_sum(lines,field):
    vals=[x[field] for x in lines]
    return sum(vals,Decimal(0)) if vals and all(v is not None for v in vals) else None
def money(lines):return [complete_sum(lines,f) for f in ['goods','vat','payment']]
def products(lines):return Counter(tuple(x[f] for f in ['name','unit','qty','price','goods']) for x in lines)
def rounded_products(lines):return Counter((x['name'],x['unit'],x['qty'],x['price'].quantize(Decimal(1),rounding=ROUND_HALF_UP),x['goods']) for x in lines)

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    WORK.mkdir(parents=True,exist_ok=True);SOURCE.mkdir(parents=True,exist_ok=True)
    prior=json.loads((PRIOR/'summary.json').read_text(encoding='utf-8'))
    adjustments=json.loads((PRIOR/'adjustments.json').read_text(encoding='utf-8'))
    replacements=json.loads((PRIOR/'replacements.json').read_text(encoding='utf-8'))
    manifest=[]
    with ZipFile(ARCHIVE) as z:
        assert z.testzip() is None
        for e in z.infolist():
            if e.is_dir():continue
            name=PurePosixPath(e.filename.replace('\\','/'))
            assert len(name.parts)==1 and name.suffix.lower() in {'.xls','.xlsx'}
            path=(SOURCE/name.name).resolve();assert path.parent==SOURCE.resolve()
            raw=z.read(e)
            if path.exists():assert path.read_bytes()==raw
            else:path.write_bytes(raw)
            manifest.append({'file':path.relative_to(ROOT).as_posix(),'sha256':sha(path),'bytes':len(raw)})
    details=defaultdict(list);controls=[]
    for e in manifest:
        headers={};indices=[];period=None;footer=None;file_lines=[]
        for rn,r in rows(ROOT/e['file']):
            if rn==3:period=r.get('A')
            if rn==5:
                headers={v:k for k,v in r.items() if v}
                assert {'Ký hiệu','Số hóa đơn','Tên hàng','Tính chất HHDV','Thành tiền','Tổng tiền TT'}<=headers.keys()
            if rn<6:continue
            get=lambda label:r.get(headers.get(label))
            if amount(r.get('A')) is None:
                if r.get('A')=='Tổng cộng':footer=[amount(get(f)) for f in ['Thành tiền','Tiền thuế GTGT','Tổng tiền TT']]
                continue
            indices.append(int(amount(r['A'])))
            k=key(get('Ký hiệu'),get('Số hóa đơn'))
            rec={'invoice':k,'date':invoice_day(get('Ngày hóa đơn')),'name':text(get('Tên hàng')),'unit':text(get('ĐVT')),
                 'qty':amount(get('Số lượng')),'price':amount(get('Đơn giá')),'goods':amount(get('Thành tiền')),
                 'rate':text(get('Thuế suất (%)')),'vat':amount(get('Tiền thuế GTGT')),'payment':amount(get('Tổng tiền TT')),
                 'kind':text(get('Tính chất HHDV')),'promotion':text(get('Hàng KM')),
                 'cqt_status':text(get('Trạng thái gửi CQT')),'invoice_status':text(get('Trạng thái HĐ')),'issue_status':text(get('Trạng thái phát hành')),
                 'source_file':e['file'],'source_row':rn}
            details[k].append(rec);file_lines.append(rec)
        assert indices==list(range(1,len(indices)+1)),(e['file'],'STT')
        known=[sum((r[f] for r in file_lines if r[f] is not None),Decimal(0)) for f in ['goods','vat','payment']]
        assert known==footer,(e['file'],'Footer',known,footer)
        controls.append({'file':e['file'],'period':period,'headers':list(headers),'lines':len(indices),'footer_known_sums':footer})

    summary_records={}
    for e in prior['source_manifest']:
        if e['kind']=='summary':
            found,_=read_summary(ROOT/e['file']);summary_records.update(found)
    issues=[]
    missing=sorted(set(summary_records)-set(details))
    if missing:issues.append({'type':'missing_from_detail','invoices':missing})
    headers_differences=[]
    for k,s in summary_records.items():
        ls=details.get(k,[])
        if not ls:continue
        observed=money(ls)
        for i,v in enumerate(observed):
            if v is not None and s['money'][i] is not None and v!=s['money'][i]:
                headers_differences.append({'invoice':k,'field':['goods','vat','payment'][i],'summary':s['money'][i],'detail':v})
        if {l['date'] for l in ls}!={s['date']}:issues.append({'type':'date_difference','invoice':k})
    if headers_differences:issues.append({'type':'header_money_difference','items':headers_differences})

    plan_path=Path(prior['plan_file']);assert sha(plan_path)==prior['plan_sha256']
    plan=defaultdict(list)
    wb=load_workbook(plan_path,read_only=True,data_only=True)
    for rn,r in enumerate(wb['Import dieu chinh'].iter_rows(min_row=10,max_row=228,max_col=26,values_only=True),10):
        old=key(r[9],r[10]);plan[old].append({'name':text(r[14]),'unit':text(r[15]),'qty':amount(r[16]),'price':amount(r[17]),'goods':amount(r[18]),'sku':text(r[25]),'rate':amount(r[20]),'plan_row':rn})
    wb.close()
    assert len(plan)==163 and sum(map(len,plan.values()))==219
    ad_results=[];ad_lines=[];price_rounding=[]
    for a in adjustments:
        ls=details.get(a['new'],[]);pl=plan[a['old']]
        exact_products=products(ls)==products(pl)
        price_rounding_only=not exact_products and rounded_products(pl)==products(ls)
        checks={'present':bool(ls),'line_count':len(ls)==len(pl),'product_values_or_integer_price_rounding':exact_products or price_rounding_only,
                'money':money(ls)==[amount(v) for v in a['expected']],
                'date':{l['date'] for l in ls}=={date(2026,11,1)},
                'rate_8':bool(ls) and all(l['rate']=='8%' for l in ls),
                'kind':bool(ls) and all(l['kind']=='Hàng hóa, dịch vụ' for l in ls),
                'invoice_status':bool(ls) and all(l['invoice_status']=='Hoá đơn điều chỉnh' for l in ls)}
        result={**a,'detail_lines':len(ls),'expected_lines':len(pl),'detail_money':money(ls),'detail_cqt_statuses':sorted({l['cqt_status'] for l in ls}),
                'detail_parent_statuses':sorted({l['invoice_status'] for l in details.get(a['old'],[])}),
                'product_values_exact':exact_products,'price_rounding_only':price_rounding_only,
                'detail_checks':checks,'sku_check':'Không có cột mã hàng hóa/dịch vụ trong bảng kê chi tiết',
                'detail_source_rows':[l['source_row'] for l in ls]}
        if not all(checks.values()):
            result['expected_products']=list(products(pl).items());result['actual_products']=list(products(ls).items())
            issues.append({'type':'adjustment_detail_check','invoice':a['new'],'checks':checks})
        ad_results.append(result)
        expected_by_values=defaultdict(list)
        for p in pl:
            price=p['price'].quantize(Decimal(1),rounding=ROUND_HALF_UP) if price_rounding_only else p['price']
            expected_by_values[(p['name'],p['unit'],p['qty'],price,p['goods'])].append(p)
        for l in ls:
            matches=expected_by_values.get(tuple(l[f] for f in ['name','unit','qty','price','goods']),[])
            expected=matches.pop(0) if matches else None
            changed_price=bool(expected and expected['price']!=l['price'])
            line={**l,'old':a['old'],'order':a['order'],'doc':a['doc'],'planned_sku':expected['sku'] if expected else None,
                  'planned_price':expected['price'] if expected else None,'plan_row':expected['plan_row'] if expected else None,'unit_price_rounded':changed_price}
            ad_lines.append(line)
            if changed_price:price_rounding.append(line)

    re_results=[]
    for r in replacements:
        ls=details.get(r['new'],[])
        oldnum=r['old'].split('/')[-1]
        expected_reason=f'Thay thế hóa đơn số {oldnum} ngày 01/11/2026 do xuất trùng đơn hàng'
        checks={'one_line':len(ls)==1,'note_kind':bool(ls) and all(l['kind']=='Ghi chú/diễn giải' for l in ls),
                'reason_name':bool(ls) and all(l['name']==expected_reason for l in ls),
                'no_nonzero_qty_price_money':bool(ls) and all(l[f] is None or l[f]==0 for l in ls for f in ['qty','price','goods','vat','payment']),
                'summary_goods_vat_zero':r['after'][:2]==[0,0],
                'date':{l['date'] for l in ls}=={date(2026,11,1)},
                'invoice_status':bool(ls) and all(l['invoice_status']=='Hoá đơn thay thế' for l in ls)}
        result={**r,'detail_lines':len(ls),'detail_money':money(ls),'detail_checks':checks,
                'detail_cqt_statuses':sorted({l['cqt_status'] for l in ls}),'detail_fields':ls[0] if len(ls)==1 else ls,
                'computed_payment_from_summary_goods_plus_vat':sum(r['after'][:2]),
                'source_payment_is_blank':all(l['payment'] is None for l in ls)}
        if not all(checks.values()):issues.append({'type':'replacement_detail_check','invoice':r['new'],'checks':checks})
        re_results.append(result)
    old_lines=json.loads((ROOT/'scripts/data/new-sales-issued-20261101/detail-lines.json').read_text(encoding='utf-8'))
    old_fields=['name','unit','qty','price','goods','rate','vat','payment','kind','promotion']
    old_changes=[]
    for k,expected in old_lines.items():
        ca=Counter(tuple(x[f] if f not in ['qty','price','goods','vat','payment'] else amount(x[f]) for f in old_fields) for x in expected)
        cb=Counter(tuple(x[f] for f in old_fields) for x in details.get(k,[]))
        if ca!=cb:old_changes.append(k)
    if old_changes:issues.append({'type':'prior_product_lines_changed','invoices':old_changes})
    ad_status=Counter(a['detail_cqt_statuses'][0] if len(a['detail_cqt_statuses'])==1 else 'Mixed' for a in ad_results)
    re_status=Counter(r['detail_cqt_statuses'][0] if len(r['detail_cqt_statuses'])==1 else 'Mixed' for r in re_results)
    status_changed=[{'invoice':a['new'],'earlier_summary':a['cqt_status'],'latest_detail':a['detail_cqt_statuses']} for a in ad_results if a['detail_cqt_statuses']!=[a['cqt_status']]]
    summary={'archive':{'file':str(ARCHIVE),'sha256':sha(ARCHIVE)},'source_manifest':manifest,'file_controls':controls,
             'all_detail_invoices':len(details),'all_detail_lines':sum(map(len,details.values())),
             'target_summary_invoices':len(summary_records),'missing_target_invoices':missing,
             'outside_summary_invoices':len(set(details)-set(summary_records)),
             'all_prior_37012_invoice_product_lines_unchanged':not old_changes,'header_money_differences':headers_differences,
             'adjustments':{'invoices':len(ad_results),'lines':len(ad_lines),'checks_passed':sum(all(a['detail_checks'].values()) for a in ad_results),
                            'money':[sum((a['detail_money'][i] for a in ad_results if a['detail_money'][i] is not None),Decimal(0)) for i in range(3)],
                            'latest_cqt_statuses':dict(ad_status),'status_updates':status_changed,'sku_exported':False,
                            'exact_product_invoices':sum(a['product_values_exact'] for a in ad_results),
                            'price_rounding_invoices':sum(a['price_rounding_only'] for a in ad_results),'price_rounding_lines':len(price_rounding)},
             'replacements':{'invoices':len(re_results),'lines':sum(r['detail_lines'] for r in re_results),'checks_passed':sum(all(r['detail_checks'].values()) for r in re_results),
                             'latest_cqt_statuses':dict(re_status),'notes':sum(r['detail_checks']['note_kind'] for r in re_results),
                             'payment_blank_in_detail':sum(r['source_payment_is_blank'] for r in re_results),
                             'rate_values':dict(Counter(l['rate'] for r in re_results for l in details[r['new']])),
                             'computed_payment_from_summary_goods_plus_vat':sum(r['computed_payment_from_summary_goods_plus_vat'] for r in re_results)},
             'base_sales_money':prior['base_sales_money'],'net_new_sales_goods':prior['new_sales_net_goods'],'net_new_sales_vat':prior['new_sales_net_vat'],
             'net_new_sales_payment_calculated':prior['new_sales_net_payment_if_replacement_totals_zero'],
             'payment_basis':'Payment is calculated from verified goods and VAT; source payment cells for note-only replacements remain blank, not overwritten as numeric zero.',
             'scope':'Bộ bán mới 2024 sau trả hàng, không cộng 273 HĐ ngoài bộ hoặc các HĐ lịch sử giữ riêng.',
             'issues':issues,'sku_limit':'Bảng kê không có cột Mã hàng hóa/dịch vụ. Mã nhóm HHDV không phải SKU.',
             'rate_limit':'Dòng ghi chú của 2.000 HĐ thay thế không xuất thuế suất; không xác minh được 8% từ bảng chi tiết.'}
    dump('summary.json',summary);dump('adjustments.json',ad_results);dump('adjustment-lines.json',ad_lines);dump('replacements.json',re_results);dump('price-rounding.json',price_rounding)
    assert sha(plan_path)==prior['plan_sha256']
    assert all(sha(ROOT/e['file'])==e['sha256'] for e in manifest)
    printable={k:v for k,v in summary.items() if k not in ['source_manifest','file_controls']}
    printable['adjustments']={k:v for k,v in summary['adjustments'].items() if k!='status_updates'}
    printable['adjustment_status_update_count']=len(status_changed)
    printable['issue_count']=len(issues);printable['issues']=issues[:4]
    print(json.dumps(printable,ensure_ascii=False,indent=2,default=native))

if __name__=='__main__':main()
