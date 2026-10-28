from pathlib import Path
import json, zipfile, os, re
from decimal import Decimal
from lxml import etree as ET

OUT=Path(__file__).parent
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def tag(n):return '{'+NS+'}'+n

def main():
    path=OUT/'ĐỐI SOÁT ĐÃ GỘP.xlsx'
    columns=json.loads((OUT/'merge_manifest.json').read_text(encoding='utf-8'))['columns']
    with zipfile.ZipFile(path) as z:
        styles=ET.fromstring(z.read('xl/styles.xml'))
        for font in styles.find(tag('fonts')):
            name=font.find(tag('name'))
            if name is None:name=ET.SubElement(font,tag('name'))
            name.set('val','Arial')
            size=font.find(tag('sz'))
            if size is None:size=ET.SubElement(font,tag('sz'))
            size.set('val','10')
        numfmts=styles.find(tag('numFmts'))
        if numfmts is None:
            numfmts=ET.Element(tag('numFmts'));styles.insert(0,numfmts)
        next_id=max([163]+[int(x.get('numFmtId')) for x in numfmts])+1
        fmt={}
        for code in ['dd/mm/yyyy','dd/mm/yyyy hh:mm','#,##0.0000','@']:
            ET.SubElement(numfmts,tag('numFmt'),numFmtId=str(next_id),formatCode=code)
            fmt[code]=next_id;next_id+=1
        numfmts.set('count',str(len(numfmts)))
        borders=styles.find(tag('borders'))
        border_id=len(borders)
        border=ET.SubElement(borders,tag('border'))
        for side in ['left','right','top','bottom']:
            edge=ET.SubElement(border,tag(side),style='thin')
            ET.SubElement(edge,tag('color'),rgb='FFD8E0E5')
        ET.SubElement(border,tag('diagonal'))
        borders.set('count',str(len(borders)))
        xfs=styles.find(tag('cellXfs'));style_ids={}
        for xf in xfs:xf.set('borderId',str(border_id));xf.set('applyBorder','1')
        for key,code,align,wrap in [('text',None,'left',True),('id','@','left',False),('number','#,##0.0000','right',False),('integer',None,'right',False),('date','dd/mm/yyyy','center',False),('datetime','dd/mm/yyyy hh:mm','right',False)]:
            style_ids[key]=len(xfs)
            xf=ET.SubElement(xfs,tag('xf'),numFmtId='3' if key=='integer' else str(fmt.get(code,0)),fontId='0',fillId='0',borderId=str(border_id),xfId='0',applyAlignment='1',applyNumberFormat='1',applyBorder='1')
            ET.SubElement(xf,tag('alignment'),horizontal=align,vertical='center',wrapText='1' if wrap else '0')
        xfs.set('count',str(len(xfs)))
        sheet=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        sf=sheet.find(tag('sheetFormatPr'))
        if sf is None:
            sf=ET.Element(tag('sheetFormatPr'));sheet.insert(2,sf)
        sf.set('defaultRowHeight','45');sf.set('customHeight','1')
        cols=sheet.find(tag('cols'))
        if cols is None:
            cols=ET.Element(tag('cols'));sheet.insert(list(sheet).index(sheet.find(tag('sheetData'))),cols)
        else:cols.clear()
        ids={'Số hóa đơn','Mã số thuế','Mã đơn hàng','Mã SP / SKU ID Kiotviet','Mã hóa đơn Kiotviet'}
        types=[]
        for i,name in enumerate(columns):
            kind='text';width=20
            if name.startswith('Ngày'):
                kind='date' if name=='Ngày hóa đơn' else 'datetime';width=17 if kind=='date' else 23
            elif name in ids or name.startswith('SKU hàng'):
                kind='id';width=31 if 'SKU' in name else 23
            elif re.match(r'^(Giá |Giảm giá|Đơn giá)',name):kind='number';width=23
            elif name.startswith('SL') or name=='Dòng nguồn':kind='number';width=12
            elif name.startswith('Tên hàng') or name=='Sản phẩm tặng kèm':width=55
            elif name=='Trạng thái đơn hàng':width=60
            elif name in {'Tệp nguồn','Ghi chú dữ liệu'}:width=85
            elif name in {'Sheet nguồn','Email','Địa chỉ'}:width=36
            types.append(kind)
            ET.SubElement(cols,tag('col'),min=str(i+1),max=str(i+1),width=str(width),customWidth='1',style=str(style_ids[kind]))
        count=0
        for row in sheet.find(tag('sheetData')):
            rn=int(row.get('r'));row.set('ht','56' if rn==1 else '45');row.set('customHeight','1')
            if rn==1:continue
            for c in row:
                if c.tag!=tag('c'):continue
                idx=0
                for ch in re.sub(r'\d','',c.get('r')):idx=idx*26+ord(ch)-64
                kind=types[idx-1]
                v=c.find(tag('v'))
                if kind=='number' and c.get('t') not in {'s','str','inlineStr','e','b'} and v is not None and v.text:
                    value=Decimal(v.text)
                    if value==value.to_integral_value():kind='integer'
                c.set('s',str(style_ids[kind]));count+=1
        changes={'xl/styles.xml':ET.tostring(styles,encoding='UTF-8',xml_declaration=True,standalone=True),'xl/worksheets/sheet1.xml':ET.tostring(sheet,encoding='UTF-8',xml_declaration=True,standalone=True)}
        tmp=path.with_suffix('.formatted.xlsx')
        with zipfile.ZipFile(tmp,'w') as dest:
            for item in z.infolist():dest.writestr(item,changes.get(item.filename,z.read(item.filename)))
    os.replace(tmp,path)
    print(json.dumps({'formatted_cells':count,'columns':len(columns),'file_size':path.stat().st_size}))

if __name__=='__main__':main()
