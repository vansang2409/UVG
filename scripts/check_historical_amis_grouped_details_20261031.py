"""Use historical line data only when its voucher totals equal the live AMIS snapshot."""
import subprocess,io,json,sys,hashlib,csv,re
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal as D
from openpyxl import load_workbook
P=Path(__file__).resolve().parents[1];O=P/'outputs/doi-chieu-amis-j-20261031'
REV='9c50b0424a242695fa31182ab826613bf0fa699f';FILE='data/Du lieu Amis/So_chi_tiet_ban_hang T1-6.2025.xlsx'

def main():
    b=subprocess.check_output(['git','show',f'{REV}:{FILE}'],cwd=P)
    w=load_workbook(io.BytesIO(b),read_only=True,data_only=True);ws=w.worksheets[0]
    hist=defaultdict(list);prefix=Counter()
    for rn,r in enumerate(ws.iter_rows(min_row=5,values_only=True),5):
        if not r[2] or not isinstance(r[34],(int,float)):continue
        v=str(r[2]);prefix[v[:4]]+=1
        hist[v].append({'row':rn,'date':str(r[0])[:10],'signature':str(r[3] or ''),'invoice':str(r[5] or ''),'sku':str(r[12] or ''),'name':str(r[13] or ''),'unit':str(r[16] or ''),'qty':str(r[17] or 0),'net':str(D(str(r[20] or 0))-D(str(r[25] or 0))-D(str(r[27] or 0))-D(str(r[28] or 0))),'vat':str(r[32] or 0),'payment':str(r[34])})
    w.close()
    live=json.loads((O/'amis-all-snapshot.json').read_text(encoding='utf-8'));ix={r[2]:r for r in live[1:]}
    checks={}
    for v,ls in hist.items():
        sums={k:sum(D(l[k]) for l in ls) for k in ('net','vat','payment')}
        current=ix.get(v)
        checks[v]={'matches_live':bool(current and sums['vat']==D(str(current[6])) and sums['payment']==D(str(current[7]))),'sums':{k:str(z) for k,z in sums.items()},'rows':ls}
    a=json.loads((O/'multi-voucher-audit.json').read_text(encoding='utf-8'));cause=json.loads((O/'cause-investigation.json').read_text(encoding='utf-8'));ci={g['invoice']:g for g in cause['groups']}
    results=[]
    for g in a['groups']:
        vs=g['vouchers'];ok=[v for v in vs if checks.get(v,{}).get('matches_live')];bad=[v for v in vs if v in checks and not checks[v]['matches_live']];missing=[v for v in vs if v not in checks]
        gs=[l for v in ok for l in hist[v]]
        # Exact unit and family are diagnostics only, not exact SKU assignment.
        def model(name):
            m=re.search(r'\b(K[AM]\s*\d+[A-Z]*)',name.upper())
            return re.sub(r'\s','',m[1]) if m else 'KHONG_NHAN_DIEN'
        am=defaultdict(lambda:{'qty':D(0),'payment':D(0),'vat':D(0)})
        for l in gs:
            z=am[(model(l['name']),l['unit'].strip().casefold())];z['qty']+=D(l['qty']);z['payment']+=D(l['payment']);z['vat']+=D(l['vat'])
        source=ci[g['invoice']]['source_lines'];mi=defaultdict(lambda:{'qty':D(0),'payment':D(0),'vat':D(0)})
        for l in source:
            z=mi[model(l['name'])];z['qty']+=D(l['qty'] or 0);z['payment']+=D(l['payment']) if l['payment']!='None' else 0;z['vat']+=D(l['vat']) if l['vat']!='None' else 0
        results.append({'invoice':g['invoice'],'vouchers':len(vs),'historical_matches_live':len(ok),'historical_value_changed':bad,'no_historical_details':missing,'usable_historical_lines':len(gs),'amis_model_unit':[{'model':k[0],'unit':k[1],**{f:str(v) for f,v in z.items()}} for k,z in am.items()],'meinvoice_model':[{'model':k,**{f:str(v) for f,v in z.items()}} for k,z in mi.items()],'historical_explicit_invoice_refs':list({(l['signature'],l['invoice']) for l in gs if l['signature'] or l['invoice']})})
    summary={'historical_revision':REV,'historical_file':FILE,'historical_sha256':hashlib.sha256(b).hexdigest(),'historical_vouchers':len(hist),'prefix_counts':dict(prefix),'groups_all_vouchers_match_live':sum(g['historical_matches_live']==g['vouchers'] for g in results),'groups_with_changed_historical_values':sum(bool(g['historical_value_changed']) for g in results),'groups_with_missing_historical_vouchers':sum(bool(g['no_historical_details']) for g in results),'matched_live_vouchers':sum(g['historical_matches_live'] for g in results)}
    (O/'historical-amis-detail-check.json').write_text(json.dumps({'summary':summary,'groups':results,'voucher_checks':checks},ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    for n in ['00000427','00000431','00000465','00000428']:
        g=next(g for g in results if g['invoice']==n);print(json.dumps(g,ensure_ascii=False))

if __name__=='__main__':main()
