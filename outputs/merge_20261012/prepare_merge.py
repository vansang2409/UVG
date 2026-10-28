from inspect_sources import *
from decimal import Decimal
import hashlib

BASE = ['Ngày hóa đơn','Số hóa đơn','Họ & Tên','Mã số thuế','Địa chỉ','Email','Tên người mua','Mã đơn hàng','Trạng thái đơn hàng','Ngày đặt hàng','Ngày khách nhận được hàng','Ngày hoàn thành thanh toán','Giá sản phẩm','Giảm giá của shop','Giá sau giảm (có VAT)','Giá trị trước thuế','Trả hàng/hoàn tiền','Mã SP / SKU ID Kiotviet','SL (Kiotviet)']
PRODUCTS = [x for i in range(1,6) for x in (f'Tên hàng {i}',f'SKU hàng {i}',f'SL{i}',f'Đơn giá {i} (trước thuế)')]
EXTRAS = ['Sản phẩm tặng kèm','Mã hóa đơn Kiotviet']
PROV = ['Sàn','Loại hóa đơn','Thư mục nguồn','Tệp nguồn','Sheet nguồn','Dòng nguồn','Ghi chú dữ liệu']
ERRORS = {'#N/A','#REF!','#VALUE!','#DIV/0!','#NAME?','#NUM!','#NULL!','#SPILL!','#CALC!'}
ALIASES = {
 'ngay hd':'Ngày hóa đơn','ngay hoa don':'Ngày hóa đơn','so hoa don':'Số hóa đơn','ho & ten':'Họ & Tên','ma so thue':'Mã số thuế','dia chi':'Địa chỉ','email':'Email',
 'ten nguoi mua':'Tên người mua','ma don hang':'Mã đơn hàng','order id':'Mã đơn hàng','trang thai don hang':'Trạng thái đơn hàng',
 'ngay dat hang':'Ngày đặt hàng','ngay khach nhan duoc hang':'Ngày khách nhận được hàng','ngay khach nhan hang':'Ngày khách nhận được hàng',
 'ngay hoan thanh thanh toan':'Ngày hoàn thành thanh toán','ngay thanh toan':'Ngày hoàn thành thanh toán','gia san pham':'Giá sản phẩm','gia sp':'Giá sản phẩm',
 'ma giam gia':'Giảm giá của shop','giam gia/khuyen mai cua shop':'Giảm giá của shop','gia sau khi giam vc cua shop(co vat)':'Giá sau giảm (có VAT)',
 'gia sau giam (bao gom vat)':'Giá sau giảm (có VAT)','gia the hien tren hoa don (da tru vat)':'Giá trị trước thuế','gia tri truoc thue':'Giá trị trước thuế',
 'tra hang/hoan tien':'Trả hàng/hoàn tiền','tra hang (neu co)':'Trả hàng/hoàn tiền','tra hang(neu co)':'Trả hàng/hoàn tiền',
 'check map sku id kiotviet':'Mã SP / SKU ID Kiotviet','ma sp (sku id)':'Mã SP / SKU ID Kiotviet','sl':'SL (Kiotviet)','sl (kiotviet)':'SL (Kiotviet)',
 'ten hang':'Tên hàng 1','san pham 1':'Tên hàng 1','san pham tang kem':'Sản phẩm tặng kèm','ma hoa don kiotviet':'Mã hóa đơn Kiotviet'
}
for i in range(1,6):
    for src,dst in [(f'ten hang {i}',f'Tên hàng {i}'),(f'sku hang {i}',f'SKU hàng {i}'),(f'sl{i}',f'SL{i}'),(f'don gia {i}',f'Đơn giá {i} (trước thuế)'),(f'don gia {i} (truoc thue)',f'Đơn giá {i} (trước thuế)')]: ALIASES[src]=dst

def col(ref): return re.sub(r'\d','',ref)
def colnum(c):
    n=0
    for t in c: n=n*26+ord(t)-64
    return n
def sortfile(p):
    month=int(re.search(r'Tháng (\d+)',str(p)).group(1))
    dt=re.findall(r'(\d{2})\.(\d{2})',p.name)
    return (month,'Tiktok' in p.name,int(dt[-1][0]) if dt else 0,p.name)
def map_headers(raw,platform,kind):
    mapping={}; fixes=[]; occurrences=collections.Counter()
    for c,h in sorted(raw.items(),key=lambda t:colnum(t[0])):
        s=plain(h)
        if not s or s.startswith('#'): continue
        k=ALIASES.get(s, str(h).strip())
        occurrences[k]+=1
        if occurrences[k]>1: k=f'{k} (bổ sung {occurrences[k]})'
        mapping[c]=k
    assumptions={}
    if platform=='Shopee': assumptions={'H':'Mã đơn hàng','I':'Trạng thái đơn hàng'}
    else:
        assumptions={'A':'Ngày hóa đơn'}
        if kind=='1 sản phẩm': assumptions['P']='Tên hàng 1'
    for c,k in assumptions.items():
        if c not in mapping:
            mapping[c]=k; fixes.append({'column':c,'old':raw.get(c),'mapped':k})
    if kind=='1 sản phẩm' and 'Tên hàng 1' not in mapping.values():
        qtycol=next((c for c,k in mapping.items() if k=='SL1'),None)
        if qtycol and len(qtycol)==1:
            productcol=chr(ord(qtycol)-1)
            if productcol not in mapping:
                mapping[productcol]='Tên hàng 1'
                fixes.append({'column':productcol,'old':raw.get(productcol),'mapped':'Tên hàng 1'})
    assert len(mapping)==len(set(mapping.values())), mapping
    return mapping,fixes
def typed(c,ss,key):
    v=cell_value(c,ss)
    if v is None or v=='': return None
    typ=c.attrib.get('t')
    if key in {'Số hóa đơn','Mã số thuế','Mã đơn hàng','Mã hóa đơn Kiotviet'} or key.startswith('SKU hàng') or key=='Mã SP / SKU ID Kiotviet':
        if typ not in ('s','inlineStr','str','e'):
            d=Decimal(v)
            return str(d.quantize(Decimal(1))) if d==d.to_integral_value() else v
        return v
    if typ in ('s','inlineStr','str','e'): return v
    if typ=='b': return v=='1'
    d=Decimal(v)
    return int(d) if d==d.to_integral_value() else float(d)

def main():
    columns=BASE+PRODUCTS+EXTRAS
    metadata=[]; issues=[]; fixes=[]; blank=0; errors=collections.Counter(); missingcache=0; rowcount=0
    digest=hashlib.sha256(); totals=collections.defaultdict(Decimal); kinds=collections.Counter(); months=collections.Counter()
    files=sorted((p for p in ROOT.rglob('*.xlsx') if not p.name.startswith('~$')),key=sortfile)
    with (OUT/'rows.jsonl').open('w',encoding='utf-8') as output:
        for idx,p in enumerate(files):
            with zipfile.ZipFile(p) as z:
                targets=[(n,t) for n,t in book_info(z) if is_target(n)]
                assert len(targets)==2, (p,targets)
                ss=strings(z)
                for name,target in targets:
                    platform='Tiktok' if 'Tiktok' in p.name else 'Shopee'
                    kind='nhiều sản phẩm' if 'nhieu' in plain(name) else '1 sản phẩm'
                    mapping=None; count=0; firstout=rowcount+2; noid=[]; unlabelled=collections.Counter(); nonempty=0
                    with z.open(target) as f:
                        for event,e in ET.iterparse(f,events=('end',)):
                            if e.tag != '{'+NS['m']+'}row': continue
                            rn=int(e.attrib['r']); cells=[c for c in e if c.tag.endswith('}c')]
                            if rn==2:
                                raw={col(c.attrib['r']):cell_value(c,ss) for c in cells}
                                mapping,hfix=map_headers(raw,platform,kind)
                                if hfix: fixes.append({'file':str(p.relative_to(ROOT)),'sheet':name,'fixes':hfix})
                                for k in mapping.values():
                                    if k not in columns: columns.append(k)
                                e.clear(); continue
                            if rn<3: e.clear(); continue
                            assert mapping is not None
                            data={}; sourcevalues=0; rowerrors=[]; rowmissing=0
                            for c in cells:
                                cc=col(c.attrib['r']); key=mapping.get(cc)
                                rawv=cell_value(c,ss)
                                if rawv is None or rawv=='':
                                    if c.find('m:f',NS) is not None: rowmissing+=1
                                    continue
                                if key is None:
                                    key=f'Dữ liệu không có tiêu đề ({platform}, {cc})'
                                    mapping[cc]=key
                                    if key not in columns: columns.append(key)
                                    unlabelled[cc]+=1
                                value=typed(c,ss,key)
                                if value is not None:
                                    data[key]=value; sourcevalues+=1
                                if c.attrib.get('t')=='e' or rawv in ERRORS: rowerrors.append((key,rawv))
                            if not data:
                                blank+=1; e.clear(); continue
                            nonempty+=1
                            ids=[data.get(k) for k in ['Mã đơn hàng','Số hóa đơn']]
                            notes=[]
                            if not any(v is not None and str(v).strip() and not str(v).startswith('#') for v in ids):
                                if all(str(v).startswith('#') or v in (0,False,'******','*******') for v in data.values()):
                                    issues.append({'file':p.name,'sheet':name,'row':rn,'excluded':'template_only','values':data})
                                    e.clear(); continue
                                noid.append({'row':rn,'values':data})
                                notes.append('Thiếu mã đơn hàng và số hóa đơn trong file nguồn')
                            assert sourcevalues==len(data),(p,name,rn)
                            for key,value in rowerrors: errors[value]+=1
                            if rowerrors: notes.append('Lỗi có sẵn trong nguồn: '+', '.join(sorted(set(v for k,v in rowerrors))))
                            missingcache+=rowmissing
                            data.update(dict(zip(PROV,[platform,kind,str(p.parent.relative_to(ROOT)),p.name,name,rn])))
                            if notes: data['Ghi chú dữ liệu']='; '.join(notes)
                            line=json.dumps(data,ensure_ascii=False,separators=(',',':'))
                            output.write(line+'\n'); digest.update((line+'\n').encode('utf-8'))
                            for k in ['Giá sản phẩm','Giảm giá của shop','Giá sau giảm (có VAT)','Giá trị trước thuế']:
                                v=data.get(k)
                                if isinstance(v,(int,float)) and not isinstance(v,bool): totals[k]+=Decimal(str(v))
                            rowcount+=1; count+=1; kinds[f'{platform}: {kind}']+=1; months[str(p.parent.name)]+=1
                            e.clear()
                    metadata.append({'file':str(p.relative_to(ROOT)),'sheet':name,'rows':count,'start':firstout,'end':rowcount+1,'mapping':mapping,'unlabelled':dict(unlabelled),'no_id_rows':noid})
            if (idx+1)%25==0: print(f'Prepared {idx+1}/{len(files)} files; {rowcount:,} rows',flush=True)
    result={'files':len(files),'sheets':len(metadata),'row_count':rowcount,'columns':columns+PROV,'sha256_jsonl':digest.hexdigest(),'totals':{k:str(v) for k,v in totals.items()},'by_kind':dict(kinds),'by_month':dict(months),'source_errors':dict(errors),'formula_empty_cache':missingcache,'header_fixes':fixes,'excluded_template_rows':issues,'sources':metadata,'blank_rows_ignored':blank}
    (OUT/'merge_manifest.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k not in ['sources','header_fixes','excluded_template_rows']},ensure_ascii=False,indent=2))
    print('header fixes:',len(fixes),'unlabelled sources:',sum(bool(x['unlabelled']) for x in metadata),'sources with no-id rows:',sum(bool(x['no_id_rows']) for x in metadata),'excluded template rows:',len(issues))

if __name__=='__main__': main()
