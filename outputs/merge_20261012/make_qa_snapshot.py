from format_merge import *

path=OUT/'ĐỐI SOÁT ĐÃ GỘP.xlsx'
with zipfile.ZipFile(path) as z:
    sheet=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    data=sheet.find(tag('sheetData'))
    for row in list(data):
        if int(row.get('r'))>7:data.remove(row)
    dim=sheet.find(tag('dimension'))
    if dim is not None:dim.set('ref','A1:AX7')
    changes={'xl/worksheets/sheet1.xml':ET.tostring(sheet,encoding='UTF-8',xml_declaration=True,standalone=True)}
    for name in z.namelist():
        if name.startswith('xl/tables/') and name.endswith('.xml'):
            table=ET.fromstring(z.read(name));table.set('ref','A1:AX7')
            af=table.find(tag('autoFilter'))
            if af is not None:af.set('ref','A1:AX7')
            changes[name]=ET.tostring(table,encoding='UTF-8',xml_declaration=True,standalone=True)
    with zipfile.ZipFile(OUT/'qa_snapshot.xlsx','w') as dest:
        for item in z.infolist():dest.writestr(item,changes.get(item.filename,z.read(item.filename)))
print('Prepared exact saved first seven rows for visual verification')
