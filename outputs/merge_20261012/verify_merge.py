from inspect_sources import *
from decimal import Decimal
import math

def main():
    manifest=json.loads((OUT/'merge_manifest.json').read_text(encoding='utf-8'))
    file=OUT/'ĐỐI SOÁT ĐÃ GỘP.xlsx'
    columns=manifest['columns']
    differences=[]; rows=0; formulas=0; totals=collections.defaultdict(Decimal); provenance=collections.Counter(); errors=collections.Counter()
    with zipfile.ZipFile(file) as z, (OUT/'rows.jsonl').open(encoding='utf-8') as expected:
        sheets=book_info(z)
        assert len(sheets)==1,sheets
        ss=strings(z)
        with z.open(sheets[0][1]) as f:
            for ev,e in ET.iterparse(f,events=('end',)):
                if e.tag != '{'+NS['m']+'}row': continue
                values={}
                for c in e:
                    if not c.tag.endswith('}c'): continue
                    ci=0
                    for letter in re.sub(r'\d','',c.attrib['r']): ci=ci*26+ord(letter)-64
                    value=cell_value(c,ss)
                    if value is None or value=='': continue
                    if c.find('m:f',NS) is not None: formulas+=1
                    typ=c.attrib.get('t')
                    if typ=='e': errors[value]+=1
                    if typ not in ('s','inlineStr','str','e'):
                        value=(value=='1') if typ=='b' else Decimal(value)
                    values[ci-1]=value
                rn=int(e.attrib['r'])
                if rn==1:
                    assert [values.get(i) for i in range(len(columns))]==columns
                    e.clear();continue
                line=next(expected,None)
                assert line is not None, f'Extra output row {rn}'
                exp=json.loads(line)
                actual={columns[i]:v for i,v in values.items()}
                for key in set(actual)|set(exp):
                    a=actual.get(key); b=exp.get(key)
                    same=(a==b)
                    if isinstance(a,Decimal) and isinstance(b,(int,float)) and not isinstance(b,bool):
                        same=abs(a-Decimal(str(b)))<=Decimal('0.0000001')
                    if not same and len(differences)<30: differences.append({'row':rn,'column':key,'actual':str(a),'expected':b})
                for k in manifest['totals']:
                    if isinstance(actual.get(k),Decimal): totals[k]+=actual[k]
                provenance[(actual.get('Thư mục nguồn'),actual.get('Tệp nguồn'),actual.get('Sheet nguồn'))]+=1
                rows+=1; e.clear()
        assert next(expected,None) is None,'Missing output rows'
        table_files=[n for n in z.namelist() if re.fullmatch(r'xl/tables/table\d+\.xml',n)]
        assert len(table_files)==1,table_files
        table=ET.fromstring(z.read(table_files[0]))
        assert table.find('m:autoFilter',NS) is not None
        head=z.read(sheets[0][1])[:10000].decode('utf-8')
        assert 'state="frozen"' in head and 'ySplit="1"' in head
    assert rows==manifest['row_count'],(rows,manifest['row_count'])
    assert not differences,differences
    assert formulas==0,formulas
    assert not errors,dict(errors)
    for source in manifest['sources']:
        path=Path(source['file'])
        key=(str(path.parent),path.name,source['sheet'])
        assert provenance[key]==source['rows'],(key,provenance[key],source['rows'])
    totaldiff={k:str(totals[k]-Decimal(v)) for k,v in manifest['totals'].items()}
    assert all(abs(Decimal(v))<Decimal('0.01') for v in totaldiff.values()),totaldiff
    report={'status':'PASS','files':manifest['files'],'source_sheets':manifest['sheets'],'output_sheets':1,'data_rows':rows,'columns':len(columns),'cell_value_differences':len(differences),'formulas':formulas,'output_excel_errors':dict(errors),'money_total_differences':totaldiff,'source_errors_preserved_as_text':manifest['source_errors'],'file_size':file.stat().st_size}
    (OUT/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
