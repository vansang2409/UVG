"""Prepare a source-grounded draft revenue report, shaped like the supplied 2024 tab."""
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal
from datetime import date
import json,csv,hashlib,sys
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109';WORK=OUT/'report-2025'
sys.stdout.reconfigure(encoding='utf-8')
def read(p):return json.loads(p.read_text(encoding='utf-8'))
def n(v):return None if v is None or str(v).strip()=='' else Decimal(str(v))
def formula_number(v):
    value=n(v)
    return str(int(value)) if value==int(value) else format(value,'f').replace('.',',')
def native(v):
    if isinstance(v,Decimal):return int(v) if v==int(v) else float(v)
    if isinstance(v,list):return list(map(native,v))
    if isinstance(v,dict):return {k:native(x) for k,x in v.items()}
    return v
def save(path,v):
    temporary=path.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(native(v),ensure_ascii=False,separators=(',',':')),encoding='utf-8')
    temporary.replace(path)
def main():
    WORK.mkdir(exist_ok=True);a=read(OUT/'authoring.json');p=read(OUT/'prepared.json');u=read(OUT/'unmapped-meinvoice-audit.json')
    script=ROOT/'scripts/classify_2025_20261105.py';prefix=script.read_text(encoding='utf-8').split("amis=read(SRC/'plan25AMIS.json')",1)[0].replace("OUT=ROOT/'outputs/phan-loai-2025-20261105'",'OUT=Path('+repr(str(OUT))+')')
    ns={'__file__':str(script),'__name__':'report_source'};exec(compile(prefix,str(script),'exec'),ns);ledger=ns['ledger'];print(json.dumps({'stage':'loaded_ledger','invoices':len(ledger)}),flush=True)
    byroot=defaultdict(list)
    for d in p['decisions']:
        for root in d['roots']:byroot[root].append(d)
    comps={r['root']:r for r in u['components'] if len(r['roots'])==1}
    scope={root:r for root,r in comps.items() if str(ledger.get(root,{}).get('date','')).startswith('2025-')}
    for root,c in p['chains'].items():
        if root not in scope:scope[root]=comps.get(root,{'root':root,'keys':c['keys'],'status':'mapped_or_candidate','issues':[]})
    company={r['root'] for r in u['components'] if r['status']=='company'}|{root for root,ds in byroot.items() if any(d['action']=='company' for d in ds)}
    planned={r['root']:r for r in a['zero_audit']+a['unmapped_zero_audit']};assert set(planned)<=set(scope)
    with (OUT/'DOI_CHIEU_TOAN_BO_65493_CHUNG_TU.csv').open(encoding='utf-8-sig') as f:final={r['Chứng từ AMIS']:r['Phân loại'] for r in csv.DictReader(f)}
    keepgoods=sum(n(p['docs'][code]['goods']) for code,label in final.items() if label=='Giữ nguyên, không import')
    pendinggoods=sum(n(p['docs'][code]['goods']) for code,label in final.items() if label=='Chờ xác minh')
    pending_count=sum(label=='Chờ xác minh' for label in final.values())
    exchanges=a['summary'].get('psc2_exchanges',{})
    exchange_goods=sum(n(row['excluded_goods']) for row in exchanges.values())
    gift_policy=OUT/'unmapped-gift-confirmation.json'
    gift_decision=read(gift_policy) if gift_policy.exists() else {}
    kept_gifts=set(gift_decision['invoice_roots']) if gift_decision.get('accountant_decision')=='keep' else set()
    assert not kept_gifts.intersection(planned)
    adj=defaultdict(Decimal);adjseq=defaultdict(set)
    for o in a['outputs']:
        if o['kind']=='adjustment':
            for row in o['rows']:
                if row[19] is not None:adj[row[25]]+=n(row[19]);adjseq[row[25]].add(row[0])
    roots=[];detail=[];expected=[];states=Counter();G=J=K=L=O=Decimal(0);basis=[]
    for root,rec in sorted(scope.items(),key=lambda x:(str(ledger.get(x[0],{}).get('date','')),x[0])):
        ch=p['chains'].get(root) or rec if rec.get('balance') is not None else p['chains'].get(root)
        if ch is None or ch.get('balance') is None:ch=ns['chain'](root)
        keys=ch.get('keys',rec['keys']);active=ch.get('active',[]);excluded=ch.get('excluded',[k for k in keys if k not in active]);balance=ch.get('balance',[None]*3);issues=ch.get('issues',[])
        if root in comps and comps[root].get('balance_evidence'):basis.append({'root':root,'basis':comps[root]['balance_evidence']})
        ds=byroot.get(root,[]);pending=any(d['action']=='review' and d['doc'] not in exchanges for d in ds);gift=rec.get('status')=='kt_review';complete=all(n(v) is not None for v in balance)
        status='';hold=root in company or gift or pending or bool(issues) or not complete
        if root in planned:assert complete and not hold,(root,issues)
        if root in company:status='Công ty — ngoài phạm vi'
        elif gift:
            status='Biếu tặng — KT chốt giữ, theo dõi riêng' if root in kept_gifts else 'Biếu tặng — chờ KT, chưa cộng DT'
        elif issues or not complete:status='Chờ nguồn/liên kết — chưa cộng DT'
        elif pending:status='Chờ đối chiếu AMIS — chưa cộng DT'
        elif root in planned:status='Về0 theo dự thảo, chưa phát hành'
        elif all(n(v)==0 for v in balance):status='Nguồn đã về0, không tạo thêm'
        elif ds and all(d['action']=='keep' for d in ds):status='Giữ theo kết quả đối soát tiền'
        else:status='Chờ xác định phạm vi — chưa cộng DT';hold=True
        states[status]+=1;kind=planned[root]['kind'] if root in planned else 'adjustment' if any(ns['parents'][k] for k in keys) else 'original'
        kindlabel={'adjustment':'Điều chỉnh','replacement':'Thay thế','original':'F0'}[kind]
        if hold:
            g=h=i=l=None;j=k=Decimal(0);raw=n(balance[0]) if n(balance[0]) is not None else n(ledger.get(root,{}).get('money',[None])[0]);pendingvalue=raw
        else:
            g,h,i=map(n,balance);j=adj[root];k=-g if root in planned and planned[root]['kind']=='replacement' else Decimal(0);l=g+j+k;pendingvalue=None
            if root in planned:assert l==0,root
            else:assert j==0 and k==0
            G+=g;J+=j;K+=k;L+=l
        if pendingvalue is not None:O+=pendingvalue
        day=ledger.get(root,{}).get('date');serial=(date.fromisoformat(day)-date(1899,12,30)).days if day else None
        links=[]
        for old in sorted(keys):
            for child in sorted(ns['children'][old]&set(keys)):links.append(old+' — '+('thay thế' if child in ns['replacement'][old] else 'điều chỉnh')+' → '+child)
        sources=[]
        for key in keys:
            src=ledger.get(key,{}).get('source',{});sources.append(key+' | '+str(src.get('file',''))+' | dòng '+str(src.get('row','')))
        note='; '.join(issues)
        if ds:note=('AMIS '+str(len(ds))+' CT; '+note).strip('; ')
        if root in adjseq:note+='; dự thảo điều chỉnh STT '+', '.join(map(str,sorted(adjseq[root])))
        if root in planned and planned[root]['kind']=='replacement':note+='; dự thảo thay thế STT '+str(planned[root]['sequence'])
        if hold:note+='; giá trị chờ/ngoài phạm vi ở cộtO không cộng vào tổng dự kiến'
        if day and not day.startswith('2025'):note+='; gốc khác năm, có liên kết AMIS2025'
        row=[root,serial,kindlabel,len(keys),'\n'.join(active),'\n'.join(excluded) or None,g,h,i,j,k,l,None,status,pendingvalue,'\n'.join(links) or None,'\n'.join(sources),note.strip('; ') or None]
        assert all(not isinstance(v,str) or len(v)<=50000 for v in row),root
        rootkeep=sum(n(d['goods']) for d in ds if final[d['doc']]=='Giữ nguyên, không import');haskeep=any(final[d['doc']]=='Giữ nguyên, không import' for d in ds)
        if haskeep:
            row[12]=(l or Decimal(0))-rootkeep;row[17]=(row[17] or '')+'; tiền hàng AMIS giữ '+str(rootkeep)+'; chênh nhóm '+str(row[12])
        roots.append({'root':root,'state':status,'hold':hold,'balance':balance,'change_adjustment':j,'change_replacement':k,'after':l,'keys':keys,'mapped_docs':[d['doc'] for d in ds],'amis_keep_goods':rootkeep,'has_amis_keep':haskeep});detail.append(row);expected.append(row.copy())
    assert G+J+K==L and J==sum(adj.values()),'root totals'
    sales=sum(n(d['goods']) for d in p['docs'].values());companysales=sum(n(p['docs'][d['doc']]['goods']) for d in p['decisions'] if d['action']=='company')
    newgoods=n(a['summary']['new_goods']);ret=next(o for o in a['outputs'] if o['kind']=='return_adjustment');bad=set(a['summary']['returns']['issues'])
    allreturns=sum(n(r[18]) for r in ret['rows']);readyreturns=sum(n(r[18]) for r in ret['rows'] if r[25] not in bad);heldreturns=allreturns-readyreturns
    amisnet=sales-companysales+allreturns;projected=L+newgoods+readyreturns;gap=projected-amisnet
    titles=['BÁO CÁO PHƯƠNG ÁN HÓA ĐƠN THEO AMIS2025 — 09/11/2026','Đơn vị: đồng, chưa VAT. Nguồn meInvoice đến01/11/2026; các file mới là dự thảo, chưa nhập/ký/phát hành. Khoản chờ/công ty tách cộtO.']
    summaries=[['Các bước xử lý','Số tiền (đ)','Giải thích'],['1. Hóa đơn cũ trước xử lý — phạm vi đủ dữ liệu',G,'Tổng cộtG; gồm gốc2025 và các chuỗi gốc khác năm đã nối AMIS2025. Các khoản chờ/công ty ở cộtO không cộng.'],['2. Điều chỉnh về0 — dự kiến',J,'25 HĐ/18 chuỗi; đảo7 khoản điều chỉnh cũ còn tiền rồi giảm chi tiết18 F0. Chưa phát hành.'],['3. Thay thế về0 — giảm DT dự kiến',K,'2.039 F0, mỗi HĐ thay thế có tiền0; cộtK ghi tác động loại doanh thu HĐ cũ, không phải tiền trên HĐ mới.'],['4. Hóa đơn cũ giữ lại — dự kiến',L,'G+J+K của phần đủ dữ liệu; nguồn chưa đủ căn cứ/16HĐ biếu tặng/đơn công ty tách riêng.'],['5. Bán mới AMIS — dự thảo',newgoods,'30.244 HĐ/35.016 dòng theo file bán mới; tiền bán gốc trước trả lại, đã loại396CT công ty và95CT chưa đủ dữ liệu.'],['6. Trả hàng — dự kiến đủ điều kiện số liệu',readyreturns,'319/321 phiếu trong dự thảo;2 phiếu có ngoại lệ tiền/SL giữ ngoài số dự kiến. Tất cả vẫn chờ HĐ tham chiếu thực tế.'],['DOANH THU THEO PHƯƠNG ÁN ĐÃ CHUẨN BỊ',projected,'Cũ giữ lại + bán mới − trả lại đủ số liệu. Không phải kết quả đã phát hành.'],['Doanh thu AMIS2025 — cùng phạm vi meInvoice',amisnet,'Bán gốc '+str(sales)+' −396CT công ty '+str(companysales)+' + trả lại321phiếu '+str(allreturns)+'.'],['CHÊNH CẦN ĐỐI CHIẾU',gap,'Dự kiến trừ AMIS cùng phạm vi; còn95CT AMIS chờ,16HĐ biếu tặng,2ngoại lệ nguồn và2phiếu trả cần KT chốt. Không tự ép chênh bằng0.'],['Nguồn chờ/ngoài phạm vi — không cộng DT',O,'Tổng cộtO; giá trị nguồn để kiểm tra, không phải phần giải thích định lượng toàn bộ chênh B12. Phần trả chờ thêm '+str(heldreturns)+'đ.']]
    headers=['HĐ gốc','Ngày gốc','Loại chuỗi','Số HĐ trước xử lý','HĐ trước xử lý được cộng','HĐ trước xử lý bị loại','DT trước xử lý tính trong PV','VAT trước','Thanh toán trước','Điều chỉnh dự kiến','Thay thế: giảm DT dự kiến','DT cũ giữ lại dự kiến','Kiểm phương án về0','Trạng thái/đánh giá','Giá trị chờ/ngoài PV (không cộng DT)','Liên kết trong chuỗi','Nguồn kiểm tra','Ghi chú']
    matrix=[[titles[0]]+[None]*17,[titles[1]]+[None]*17]+[r+[None]*15 for r in summaries]+[headers]+detail
    if exchange_goods:
        sale_count=a['summary']['new_invoices'];line_count=a['summary']['new_lines']
        matrix[7][2]=f'{sale_count} HĐ / {line_count} dòng dự thảo trước trả lại. Đã loại 396 CT công ty, {pending_count} CT chờ và {exchange_goods}đ PSC2 của hai chứng từ đổi sản phẩm.'
        matrix[11][2]=f'Dự kiến trừ AMIS cùng phạm vi. Xem các nhóm chênh tại dòng 13, gồm {pending_count} CT AMIS chờ và khoản đổi sản phẩm PSC2 đã loại.'
    if kept_gifts:
        assert states['Biếu tặng — KT chốt giữ, theo dõi riêng']==len(kept_gifts)
        matrix[6][2]=f'G+J+K của phần đủ dữ liệu. {len(kept_gifts)} HĐ biếu tặng KT chốt giữ, theo dõi riêng ở cột O cùng các nhóm ngoài phạm vi.'
    matrix[12][0]='Giải thích chênh theo nhóm';matrix[12][1]=gap
    matrix[12][2]=f'DT cũ giữ − AMIS giữ {L-keepgoods}đ. AMIS {pending_count} CT chờ −{pendinggoods}đ. Trả chờ chưa giảm +{readyreturns-allreturns}đ.'
    if exchange_goods:matrix[12][2]+=f' Trừ PSC2 đổi sản phẩm −{exchange_goods}đ.'
    matrix[12][2]+=f' Tổng {gap}đ. Giá trị cột O tách riêng.'
    matrix[13][12]='Chênh so AMIS giữ / kiểm về0'
    assert (L-keepgoods)-pendinggoods+(readyreturns-allreturns)-exchange_goods==gap
    values=[r.copy() for r in matrix];end=len(matrix)
    formula={3:f'=SUM(G15:G{end})',4:f'=SUM(J15:J{end})',5:f'=SUM(K15:K{end})',6:f'=SUM(L15:L{end})',9:'=SUM(B7:B9)',11:'=B10-B11',12:f'=SUM(O15:O{end})'}
    for i,f in formula.items():matrix[i][1]=f
    matrix[12][1]=f'=B7-{formula_number(keepgoods)}-{formula_number(pendinggoods)}+B9-({formula_number(allreturns)})'
    if exchange_goods:matrix[12][1]+=f'-{formula_number(exchange_goods)}'
    for i,r in enumerate(roots,14):
        matrix[i][11]=f'=IF(G{i+1}="";"";SUM(G{i+1};J{i+1};K{i+1}))'
        if r['has_amis_keep']:matrix[i][12]=f'=N(L{i+1})-{formula_number(r["amis_keep_goods"])}'
        elif r['root'] in planned and not r['hold']:matrix[i][12]=f'=L{i+1}';values[i][12]=0
    save(WORK/'matrix.json',matrix);save(WORK/'expected-values.json',values);save(WORK/'roots.json',roots)
    chunks=[]
    for start in range(0,end,250):
        name=f'{start}.json';save(WORK/name,{'start':start,'rows':matrix[start:start+250],'expected':values[start:start+250]});chunks.append({'file':name,'start':start,'rows':min(250,end-start)})
    inputs=[OUT/'authoring.json',OUT/'prepared.json',OUT/'unmapped-meinvoice-audit.json',OUT/'source-index.json',OUT/'source/report-2024-reference.json']
    if kept_gifts:inputs.append(gift_policy)
    summary={'old_before':G,'adjustment':J,'replacement_effect':K,'old_kept':L,'new_sales':newgoods,'returns_ready':readyreturns,'returns_all':allreturns,'returns_held':heldreturns,'amis_sales_all':sales,'company_sales':companysales,'amis_net_scope':amisnet,'projected':projected,'gap':gap,'held_source_goods':O}
    manifest={'spreadsheetId':'1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA','sheetId':700000010,'title':'Ra soat chuoi 2025','rows':end,'columns':18,'detail_rows':len(roots),'chunks':chunks,'states':dict(states),'summary':summary,'sources':[{'file':str(f.relative_to(ROOT)),'sha256':hashlib.sha256(f.read_bytes()).hexdigest()} for f in inputs]}
    manifest['summary'].update(amis_keep_goods=keepgoods,amis_pending_goods=pendinggoods,kept_group_difference=L-keepgoods,return_hold_effect=readyreturns-allreturns)
    if exchange_goods:manifest['summary']['psc2_excluded_goods']=exchange_goods
    if kept_gifts:manifest['summary']['gift_invoices_kept']=len(kept_gifts)
    save(WORK/'manifest.json',manifest)
    with (WORK/'BAO_CAO_CHUOI_2025.csv').open('w',encoding='utf-8-sig',newline='') as f:csv.writer(f).writerows(values)
    print(json.dumps(native({'rows':end,'states':dict(states),'summary':summary}),ensure_ascii=False),flush=True)
if __name__=='__main__':main()
