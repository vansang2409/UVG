"""Independent openpyxl readback of issued detail, plan and updated report."""
from collections import Counter,defaultdict
from datetime import date,datetime,timedelta
from decimal import Decimal,ROUND_HALF_UP
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
import hashlib,json,sys
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/post-final-detail-20261101'
PRIOR=ROOT/'scripts/data/post-final-issue-20261101'
REPORT=ROOT/'outputs/kiem-tra-sau-dieu-chinh-thay-the-20261101/KET_QUA_KIEM_TRA_SAU_PHAT_HANH_20261101.xlsx'

def txt(v):return '' if v is None else str(v).strip()
def dec(v):return None if v in (None,'') else Decimal(str(v))
def key(a,b):return txt(a).upper()+'/'+txt(b).removesuffix('.0').zfill(8)
def day(v):
    if isinstance(v,datetime):return v.date()
    if isinstance(v,date):return v
    for f in ['%d/%m/%Y','%Y-%m-%d']:
        try:return datetime.strptime(txt(v),f).date()
        except ValueError:pass
    return (datetime(1899,12,30)+timedelta(days=float(v))).date()
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p,**kwargs):return load_workbook(BytesIO(p.read_bytes()),read_only=True,**kwargs)
def native(v):
    if isinstance(v,Decimal):return int(v) if v==int(v) else str(v)
    if isinstance(v,date):return str(v)
    raise TypeError(type(v))

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    p=json.loads((WORK/'summary.json').read_text(encoding='utf-8'))
    prior=json.loads((PRIOR/'summary.json').read_text(encoding='utf-8'))
    adjustments=json.loads((WORK/'adjustments.json').read_text(encoding='utf-8'))
    replacements=json.loads((WORK/'replacements.json').read_text(encoding='utf-8'))
    ad_by_new={a['new']:a for a in adjustments};re_by_new={r['new']:r for r in replacements}
    plan_path=Path(prior['plan_file']);assert sha(plan_path)==prior['plan_sha256']
    source_hashes={e['file']:e['sha256'] for e in p['source_manifest']+prior['source_manifest']}
    assert sha(Path(p['archive']['file']))==p['archive']['sha256']
    for f,h in source_hashes.items():assert sha(ROOT/f)==h
    plan=defaultdict(list);wb=read(plan_path,data_only=True)
    for r in wb['Import dieu chinh'].iter_rows(min_row=10,max_row=228,max_col=26,values_only=True):
        plan[key(r[9],r[10])].append((txt(r[14]),txt(r[15]),dec(r[16]),dec(r[17]),dec(r[18])))
    wb.close();assert len(plan)==163 and sum(map(len,plan.values()))==219
    ls=defaultdict(list);target=set(ad_by_new)|set(re_by_new);all_keys=set();line_count=0
    footer_sums=[Decimal(0)]*3
    for e in p['source_manifest']:
        wb=read(ROOT/e['file'],data_only=True);s=wb.worksheets[0]
        headers=next(s.iter_rows(min_row=5,max_row=5,values_only=True));h={v:i for i,v in enumerate(headers) if v}
        assert 'Mã hàng hóa/dịch vụ' not in h
        for rn,r in enumerate(s.iter_rows(min_row=6,max_col=26,values_only=True),6):
            get=lambda name:r[h[name]]
            if not isinstance(r[0],(int,float)):
                if r[0]=='Tổng cộng':
                    assert [dec(get(f)) for f in ['Thành tiền','Tiền thuế GTGT','Tổng tiền TT']]==footer_sums
                continue
            line_count+=1;k=key(get('Ký hiệu'),get('Số hóa đơn'));all_keys.add(k)
            for i,f in enumerate(['Thành tiền','Tiền thuế GTGT','Tổng tiền TT']):
                value=dec(get(f))
                if value is not None:footer_sums[i]+=value
            if k not in target:continue
            ls[k].append({'row':rn,'name':txt(get('Tên hàng')),'unit':txt(get('ĐVT')),'qty':dec(get('Số lượng')),
                          'price':dec(get('Đơn giá')),'goods':dec(get('Thành tiền')),'vat':dec(get('Tiền thuế GTGT')),'payment':dec(get('Tổng tiền TT')),
                          'rate':txt(get('Thuế suất (%)')),'kind':txt(get('Tính chất HHDV')),'date':day(get('Ngày hóa đơn')),
                          'cqt_status':txt(get('Trạng thái gửi CQT')),'status':txt(get('Trạng thái HĐ'))})
        wb.close()
    assert line_count==115248 and len(all_keys)==59409
    assert set(ls)==target and len(target)==2163
    exact_count=0;rounded_invoices=[];rounded_lines=0;ad_sums=[Decimal(0)]*3
    for k,a in ad_by_new.items():
        actual=ls[k];expected=plan[a['old']]
        assert len(actual)==len(expected)
        actual_products=Counter((l['name'],l['unit'],l['qty'],l['price'],l['goods']) for l in actual)
        if Counter(expected)==actual_products:exact_count+=1
        else:
            rounded=Counter((n,u,q,price.quantize(Decimal(1),rounding=ROUND_HALF_UP),g) for n,u,q,price,g in expected)
            assert rounded==actual_products,(k,'Other product difference')
            rounded_invoices.append(k)
            rounded_lines+=sum(price!=price.quantize(Decimal(1),rounding=ROUND_HALF_UP) for _,_,_,price,_ in expected)
        for l in actual:
            assert l['rate']=='8%' and l['kind']=='Hàng hóa, dịch vụ' and l['date']==date(2026,11,1)
            assert l['cqt_status']=='HĐ hợp lệ' and l['status']=='Hoá đơn điều chỉnh'
        actual_money=[]
        for i,f in enumerate(['goods','vat','payment']):
            assert all(l[f] is not None for l in actual)
            value=sum((l[f] for l in actual),Decimal(0));ad_sums[i]+=value;actual_money.append(value)
        assert actual_money==[dec(v) for v in a['expected']]
    assert exact_count==157 and len(rounded_invoices)==6 and rounded_lines==7
    assert ad_sums==[Decimal(-225865926),Decimal(-18069275),Decimal(-243935201)]
    for k,r in re_by_new.items():
        assert len(ls[k])==1
        l=ls[k][0]
        expected=f"Thay thế hóa đơn số {r['old'].split('/')[-1]} ngày 01/11/2026 do xuất trùng đơn hàng"
        assert l['name']==expected and l['kind']=='Ghi chú/diễn giải'
        assert l['date']==date(2026,11,1) and l['cqt_status']=='HĐ hợp lệ' and l['status']=='Hoá đơn thay thế'
        assert all(l[f] is None for f in ['qty','price','goods','vat','payment']) and l['rate']==''
        assert r['after']==[0,0,None]

    output=load_workbook(REPORT,data_only=True);base=load_workbook(WORK/'before.xlsx',data_only=True)
    assert output.sheetnames==base.sheetnames==['Tong hop','Dieu chinh 163','Thay the 2000']
    s=output['Tong hop']
    assert [s[f'B{i}'].value for i in range(4,12)]==[163,163,163,0,2000,2000,2000,2000]
    assert [s[f'{c}18'].value for c in 'BCD']==[49307577629,3944606512,53252184141]
    assert s['D17'].value==0 and all(s[f'{c}19'].value is None for c in 'ABCDE')
    ad=output['Dieu chinh 163'];re=output['Thay the 2000']
    for rn,a in enumerate(adjustments,5):
        row=next(ad.iter_rows(min_row=rn,max_row=rn,max_col=21,values_only=True))
        assert row[3:5]==(a['old'],a['new'])
        assert row[15]=='HĐ hợp lệ' and row[16]=='Hóa đơn đã bị điều chỉnh'
        assert [row[i] for i in [7,10,13]]==a['expected'] and all(row[i]==0 for i in [8,11,14])
        assert ('đơn giá làm tròn' in row[18])==a['price_rounding_only']
    assert [ad[f'{c}174'].value for c in 'FG']==[16499.33,16499]
    for rn,r in enumerate(replacements,5):
        row=next(re.iter_rows(min_row=rn,max_row=rn,max_col=17,values_only=True))
        assert row[1:4]==(r['retained'],r['old'],r['new']) and row[8:12]==(0,0,None,'HĐ hợp lệ')
        assert row[13]=='Đúng 1 dòng ghi chú, đúng nội dung'
    changed=Counter()
    for sheet,rows_count,cols_count,allowed in [('Dieu chinh 163',167,21,{16,17,19,21}),('Thay the 2000',2004,17,{12,14,15,17})]:
        arows=base[sheet].iter_rows(min_row=5,max_row=rows_count,max_col=cols_count,values_only=True)
        brows=output[sheet].iter_rows(min_row=5,max_row=rows_count,max_col=cols_count,values_only=True)
        for rn,(ar,br) in enumerate(zip(arows,brows),5):
            for col,(av,bv) in enumerate(zip(ar,br),1):
                if av!=bv:
                    assert col in allowed,(sheet,rn,col,av,bv)
                    changed[sheet]+=1
    base.close();output.close()
    # Normal-mode metadata read checks native filters, tables and panes.
    wb=load_workbook(REPORT,data_only=False);old=load_workbook(WORK/'before.xlsx',data_only=False)
    for name in old.sheetnames:
        assert wb[name].freeze_panes==old[name].freeze_panes
        for t in old[name].tables.values():assert wb[name].tables[t.name].ref==t.ref
    assert wb['Dieu chinh 163'].tables['RoundedReturnPrices'].ref=='A173:I180'
    assert wb['Tong hop']['D17'].value=="=SUM('Thay the 2000'!I5:J2004)"
    assert wb['Tong hop']['D18'].value=='=SUM(D15:D17)'
    for sheet in wb:
        for row in sheet:
            for c in row:assert c.data_type!='e',(sheet.title,c.coordinate,c.value)
    wb.close();old.close()
    for f,h in source_hashes.items():assert sha(ROOT/f)==h
    assert sha(plan_path)==prior['plan_sha256']
    result={'passed':True,'independent_reader':'openpyxl','detail_invoices':len(all_keys),'detail_lines':line_count,
            'adjustments':163,'adjustment_lines':219,'adjustments_cqt_valid':163,'replacements_cqt_valid':2000,'correct_note_lines':2000,
            'exact_product_invoices':exact_count,'rounded_unit_price_invoices':rounded_invoices,'rounded_unit_price_lines':rounded_lines,
            'source_payment_blanks_preserved':2000,'net_goods':49307577629,'net_vat':3944606512,'payment_calculated_not_raw':53252184141,
            'sku_not_exported':True,'note_rate_not_exported':True,'unrelated_report_values_preserved':True,'tables_and_freeze_panes_preserved':True,
            'report_changed_cells':dict(changed),'sha256':sha(REPORT)}
    (WORK/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=native)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False,indent=2,default=native))

if __name__=='__main__':main()
