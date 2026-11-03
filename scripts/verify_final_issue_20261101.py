"""Independent openpyxl checks of fresh exports, import plan and audit report."""
from collections import Counter, defaultdict
from decimal import Decimal
from io import BytesIO
from pathlib import Path
import hashlib,json,sys

from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
WORK=ROOT/'scripts/data/post-final-issue-20261101'
OUTPUT=ROOT/'outputs/kiem-tra-sau-dieu-chinh-thay-the-20261101/KET_QUA_KIEM_TRA_SAU_PHAT_HANH_20261101.xlsx'


def money(v):return None if v in (None,'') else Decimal(str(v))
def key(a,b):return str(a).strip()+'/'+str(b).strip().zfill(8)


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    summary=json.loads((WORK/'summary.json').read_text(encoding='utf-8'))
    assert not summary['issues']
    actual={};links={}
    for e in summary['source_manifest']:
        path=ROOT/e['file']
        assert hashlib.sha256(path.read_bytes()).hexdigest()==e['sha256']
        w=load_workbook(BytesIO(path.read_bytes()),read_only=True,data_only=True)
        s=w.active
        if e['kind']=='summary':
            items=[];control=None;indices=[]
            for rn,r in enumerate(s.iter_rows(min_row=6,values_only=True),6):
                if r[0]=='Tổng cộng':control=[money(r[i]) for i in [15,16,17]]
                if not isinstance(r[0],(int,float)):continue
                k=key(r[1],r[2]);assert k not in actual
                indices.append(r[0])
                a={'date':r[3],'cqt':r[4],'order':r[5] or '', 'money':[money(r[i]) for i in [15,16,17]],'cqt_status':r[18],'status':r[20]}
                actual[k]=a;items.append(a)
            assert indices==list(range(1,len(indices)+1))
            assert [sum((a['money'][i] for a in items if a['money'][i] is not None),Decimal(0)) for i in range(3)]==control
        else:
            items=[]
            for r in s.iter_rows(min_row=8,values_only=True):
                if not isinstance(r[0],(int,float)):continue
                if e['kind']=='adjustment':items.append({'old':key(r[1],r[3]),'new':key(r[8],r[10]),'old_cqt':r[4],'reason':r[11]})
                else:items.append({'old':key(r[1],r[4]),'new':key(r[8],r[9]),'old_cqt':r[3]})
            links[e['kind']]=items
        w.close()
    assert len(actual)==39175 and len(links['adjustment'])==163 and len(links['replacement'])==2000
    plan_path=Path(summary['plan_file'])
    assert hashlib.sha256(plan_path.read_bytes()).hexdigest()==summary['plan_sha256']
    w=load_workbook(plan_path,read_only=True,data_only=True)
    plan=defaultdict(lambda:{'money':[Decimal(0),Decimal(0),Decimal(0)],'reasons':set(),'orders':set()})
    for r in w.active.iter_rows(min_row=10,max_row=228,max_col=26,values_only=True):
        a=plan[key(r[9],r[10])]
        for i,c in enumerate([18,21,22]):
            if r[c] is not None:a['money'][i]+=money(r[c])
        a['reasons'].add(r[13]);a['orders'].add(r[24])
    w.close()
    assert {a['old'] for a in links['adjustment']}==set(plan)
    adjustment_keys=set()
    for link in links['adjustment']:
        a=actual[link['new']];old=actual[link['old']];p=plan[link['old']]
        assert a['money']==p['money'] and a['order']==old['order'] and p['orders']=={a['order']}
        assert p['reasons']=={link['reason']} and link['old_cqt']==old['cqt']
        assert a['status']=='Hoá đơn điều chỉnh' and a['date']=='01/11/2026' and a['cqt']
        adjustment_keys.add(link['new'])
    assert adjustment_keys=={f'1C26MUV/{n:08}' for n in range(150013,150176)}
    statuses=Counter(actual[k]['cqt_status'] for k in adjustment_keys)
    assert statuses=={'HĐ hợp lệ':100,'Đã gửi CQT':63}
    pending={k for k in adjustment_keys if actual[k]['cqt_status']=='Đã gửi CQT'}
    assert pending=={f'1C26MUV/{n:08}' for n in range(150113,150176)}
    adjustment_total=[sum(actual[k]['money'][i] for k in adjustment_keys) for i in range(3)]
    assert adjustment_total==[-225865926,-18069275,-243935201]
    assert {a['old'] for a in links['replacement']}=={f'1C26MUV/{n:08}' for n in range(133274,135274)}
    assert {a['new'] for a in links['replacement']}=={f'1C26MUV/{n:08}' for n in range(148013,150013)}
    old_money=[Decimal(0),Decimal(0),Decimal(0)]
    for link in links['replacement']:
        a,old=actual[link['new']],actual[link['old']]
        assert a['money']==[Decimal(0),Decimal(0),None]
        assert a['cqt_status']=='HĐ hợp lệ' and a['status']=='Hoá đơn thay thế' and a['cqt']
        assert old['status']=='Hóa đơn đã bị thay thế' and old['cqt']==link['old_cqt']
        retained=actual['1C26MUV/'+str(int(link['old'].split('/')[1])-2000).zfill(8)]
        assert retained['status']=='Hoá đơn mới' and retained['money']==old['money']
        for i in range(3):old_money[i]+=old['money'][i]
    assert old_money==[2883691447,230695306,3114386753]
    assert summary['new_sales_net_goods']==49533443555-225865926==49307577629
    assert summary['new_sales_net_vat']==3962675787-18069275==3944606512
    assert summary['new_sales_net_payment'] is None and summary['unknown_replacement_payment_invoices']==2000
    # Read back all authored rows, identifiers, cached formulas and blank fields.
    w=load_workbook(OUTPUT,read_only=True,data_only=True)
    assert w.sheetnames==['Tong hop','Dieu chinh 163','Thay the 2000']
    s=w['Tong hop']
    assert [s[f'B{r}'].value for r in range(4,12)]==[163,163,100,63,2000,2000,2000,2000]
    assert [s['B18'].value,s['C18'].value,s['D18'].value,s['D19'].value]==[49307577629,3944606512,None,53252184141]
    assert s['D17'].value is None
    plan_rows=list(w['Dieu chinh 163'].iter_rows(min_row=5,max_row=167,max_col=21,values_only=True))
    assert len(plan_rows)==163 and {r[4] for r in plan_rows}==adjustment_keys
    for r in plan_rows:
        a=actual[r[4]];p=plan[r[3]]
        assert r[2]==a['order'] and isinstance(r[2],str)
        assert [r[7],r[10],r[13]]==[int(x) for x in a['money']]
        assert [r[8],r[11],r[14]]==[0,0,0] and r[15]==a['cqt_status'] and r[17] in p['reasons']
    replacement_rows=list(w['Thay the 2000'].iter_rows(min_row=5,max_row=2004,max_col=17,values_only=True))
    assert len(replacement_rows)==2000 and {r[3] for r in replacement_rows}=={x['new'] for x in links['replacement']}
    for r in replacement_rows:
        assert [r[8],r[9],r[10]]==[0,0,None] and r[11]=='HĐ hợp lệ' and r[12]=='Hóa đơn đã bị thay thế'
    w.close()
    result={'passed':True,'independent_reader':'openpyxl','source_invoices':39175,'adjustment_links':163,'replacement_links':2000,
            'adjustment_money':[-225865926,-18069275,-243935201],'adjustment_cqt_valid':100,'adjustment_cqt_sent':63,
            'replacements_cqt_valid':2000,'replacement_goods_vat_zero':2000,'replacement_payment_blanks_preserved':2000,
            'old_duplicate_value':[2883691447,230695306,3114386753],'net_new_2024_goods':49307577629,'net_new_2024_vat':3944606512,
            'net_payment_not_claimed_verified':True,'report_rows_and_formulas_verified':True,'sha256':hashlib.sha256(OUTPUT.read_bytes()).hexdigest()}
    (WORK/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))


if __name__=='__main__':main()
