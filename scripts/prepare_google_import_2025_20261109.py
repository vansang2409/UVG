"""Stage verified import workbooks for the user-specified existing Google spreadsheet."""
from pathlib import Path
from datetime import date,datetime
from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel
import hashlib,json,sys,math

ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'outputs/import-2025-20261109';WORK=OUT/'google-upload'
sys.stdout.reconfigure(encoding='utf-8')
TARGET='1jACpDPSlOCjq-25hF_g5-06eZKD3e2Fq88j472wjTiA'
SHEETS={'new':(0,'Import HD moi'),'adjustment':(700000101,'Dieu chinh ve 0'),'replacement':(700000102,'Thay the ve 0'),'return_adjustment':(700000103,'Dieu chinh BTL - cho so HD')}
def save(name,v):(WORK/name).write_text(json.dumps(v,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
def color(h):return {k:int(h[i:i+2],16)/255 for k,i in [('red',0),('green',2),('blue',4)]}
def main():
    WORK.mkdir(exist_ok=True)
    p=json.loads((OUT/'authoring.json').read_text(encoding='utf-8'));v=json.loads((OUT/'verification.json').read_text(encoding='utf-8'));assert v['passed']
    hashes={r['file']:r['sha256'] for r in v['exported_files']};manifest=[];init=[];styles=[]
    for o in p['outputs']:
        if o['kind']=='report':continue
        sid,title=SHEETS[o['kind']];path=OUT/'deliverables'/o['file'];sha=hashlib.sha256(path.read_bytes()).hexdigest();assert sha==hashes[o['file']]
        wb=load_workbook(path,read_only=True,data_only=False);ws=wb.active;cols=len(o['headers']);end=len(o['rows'])+9
        matrix=[]
        for r in ws.iter_rows(min_row=1,max_row=end,max_col=cols,values_only=True):
            row=[]
            for x in r:
                if isinstance(x,(datetime,date)):x=to_excel(x)
                if isinstance(x,float):assert math.isfinite(x)
                row.append(x)
            matrix.append(row)
        wb.close();assert len(matrix)==end and matrix[8]==o['headers']
        save(f'{sid}-matrix.json',matrix)
        chunks=[]
        for start in range(0,end,1000):
            filename=f'{sid}-{start}.json';save(filename,matrix[start:start+1000]);chunks.append({'file':filename,'start':start,'rows':min(1000,end-start)})
        # Google Sheets cannot merge A:H across a frozen A:B boundary.
        props={'sheetId':sid,'title':title,'gridProperties':{'rowCount':end,'columnCount':cols,'frozenRowCount':9,'frozenColumnCount':0,'hideGridlines':True}}
        init.append({'updateSheetProperties':{'properties':props,'fields':'title,gridProperties'}} if sid==0 else {'addSheet':{'properties':props}})
        rect={'sheetId':sid,'startRowIndex':0,'endRowIndex':end,'startColumnIndex':0,'endColumnIndex':cols}
        styles += [{'repeatCell':{'range':rect,'cell':{'userEnteredFormat':{'textFormat':{'fontFamily':'Arial','fontSize':10},'verticalAlignment':'MIDDLE'}},'fields':'userEnteredFormat(textFormat,verticalAlignment)'}},
                   {'updateDimensionProperties':{'range':{'sheetId':sid,'dimension':'ROWS','startIndex':0,'endIndex':end},'properties':{'pixelSize':80 if o['kind']!='new' else 38},'fields':'pixelSize'}},
                   {'updateDimensionProperties':{'range':{'sheetId':sid,'dimension':'ROWS','startIndex':0,'endIndex':7},'properties':{'pixelSize':40},'fields':'pixelSize'}},
                   {'repeatCell':{'range':{'sheetId':sid,'startRowIndex':0,'endRowIndex':7,'startColumnIndex':0,'endColumnIndex':8},'cell':{'userEnteredFormat':{'wrapStrategy':'WRAP'}},'fields':'userEnteredFormat.wrapStrategy'}}]
        for row in range(7):styles.append({'mergeCells':{'range':{'sheetId':sid,'startRowIndex':row,'endRowIndex':row+1,'startColumnIndex':0,'endColumnIndex':8},'mergeType':'MERGE_ALL'}})
        styles += [{'updateDimensionProperties':{'range':{'sheetId':sid,'dimension':'ROWS','startIndex':7,'endIndex':8},'properties':{'pixelSize':18},'fields':'pixelSize'}},
                   {'repeatCell':{'range':{'sheetId':sid,'startRowIndex':0,'endRowIndex':1,'startColumnIndex':0,'endColumnIndex':8},'cell':{'userEnteredFormat':{'textFormat':{'bold':True}}},'fields':'userEnteredFormat.textFormat.bold'}}]
        header={'sheetId':sid,'startRowIndex':8,'endRowIndex':9,'startColumnIndex':0,'endColumnIndex':cols}
        table={'sheetId':sid,'startRowIndex':8,'endRowIndex':end,'startColumnIndex':0,'endColumnIndex':cols}
        styles += [{'addBanding':{'bandedRange':{'range':table,'rowProperties':{'headerColor':color('24435C'),'firstBandColor':color('C1E5F1'),'secondBandColor':color('FFFFFF')}}}},
                   {'repeatCell':{'range':header,'cell':{'userEnteredFormat':{'backgroundColor':color('24435C'),'textFormat':{'bold':True,'foregroundColor':color('FFFFFF')},'wrapStrategy':'WRAP','horizontalAlignment':'CENTER'}},'fields':'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,horizontalAlignment)'}},
                   {'updateDimensionProperties':{'range':{'sheetId':sid,'dimension':'ROWS','startIndex':8,'endIndex':9},'properties':{'pixelSize':76},'fields':'pixelSize'}},
                   {'setBasicFilter':{'filter':{'range':table}}}]
        for i,h in enumerate(o['headers']):
            width=490 if any(t in h for t in ['Tên hàng','Ghi chú','Chờ vì','Lý do']) else 266 if any(t in h for t in ['Tên','Khách','Địa chỉ','Chuỗi','Các HĐ']) else 189 if any(t in h for t in ['Dòng nguồn','Chứng từ']) else 161
            styles.append({'updateDimensionProperties':{'range':{'sheetId':sid,'dimension':'COLUMNS','startIndex':i,'endIndex':i+1},'properties':{'pixelSize':width},'fields':'pixelSize'}})
            fmt={};fractional=any(isinstance(r[i],float) and not r[i].is_integer() for r in matrix[9:])
            if any(t in h for t in ['Mã','Số hóa đơn','Ký hiệu','Số chứng từ','Chứng từ AMIS','Chuỗi HĐ','HĐ gốc']):fmt['numberFormat']={'type':'TEXT','pattern':'@'}
            if any(t in h for t in ['Tiền','tiền','VAT','Thanh toán','thanh toán','Thành tiền','Đơn giá','Số dư']):fmt['numberFormat']={'type':'NUMBER','pattern':'#,##0.00####;[Red](#,##0.00####);0' if fractional else '#,##0;[Red](#,##0);0'}
            if any(t in h for t in ['Số lượng','SL trả']):fmt['numberFormat']={'type':'NUMBER','pattern':'#,##0.00####' if fractional else '#,##0'}
            if any(t in h for t in ['Ngày hóa đơn','Ngày HĐ']):fmt['numberFormat']={'type':'DATE','pattern':'dd/mm/yyyy'}
            if any(t in h for t in ['Tên','Ghi chú','Chờ vì','Lý do','Các HĐ','Chuỗi','Kiểm chi tiết','Địa chỉ','Dòng nguồn','Trạng thái']):fmt['wrapStrategy']='WRAP'
            if fmt:styles.append({'repeatCell':{'range':{'sheetId':sid,'startRowIndex':9,'endRowIndex':end,'startColumnIndex':i,'endColumnIndex':i+1},'cell':{'userEnteredFormat':fmt},'fields':','.join('userEnteredFormat.'+k for k in fmt)}})
        if o['kind']=='return_adjustment':
            styles.append({'repeatCell':{'range':{'sheetId':sid,'startRowIndex':9,'endRowIndex':end,'startColumnIndex':9,'endColumnIndex':13},'cell':{'userEnteredFormat':{'backgroundColor':color('FFF2CC')}},'fields':'userEnteredFormat.backgroundColor'}})
            styles.append({'addConditionalFormatRule':{'index':0,'rule':{'ranges':[{'sheetId':sid,'startRowIndex':9,'endRowIndex':end,'startColumnIndex':30,'endColumnIndex':31}],'booleanRule':{'condition':{'type':'CUSTOM_FORMULA','values':[{'userEnteredValue':'=$AE10="Chờ kiểm tra số lượng/tiền trả"'}]},'format':{'backgroundColor':color('FDE9E7'),'textFormat':{'bold':True,'foregroundColor':color('9C0006')}}}}}})
        manifest.append({'sheetId':sid,'title':title,'kind':o['kind'],'source_file':o['file'],'sha256':sha,'rows':end,'columns':cols,'invoices':o['invoices'],'chunks':chunks})
    assert len(manifest)==4
    giftpath=OUT/'gift-kt-check.json'
    if giftpath.exists():
        gift=json.loads(giftpath.read_text(encoding='utf-8'));sid=700000104;title='16 HD bieu tang - KT check';cols=len(gift['headers'])
        matrix=[['16 HĐ BIẾU TẶNG / KHÔNG THU TIỀN — KẾ TOÁN KIỂM TRA']+[None]*(cols-1),['Chưa đưa vào import. Kế toán ghi ý kiến ở cột K; nguồn meInvoice kiểm đến 01/11/2026.']+[None]*(cols-1),gift['headers']]+gift['rows']
        assert len(gift['rows'])==16
        save(f'{sid}-matrix.json',matrix);save(f'{sid}-0.json',matrix)
        manifest.append({'sheetId':sid,'title':title,'kind':'accountant_review','source_file':giftpath.name,'sha256':hashlib.sha256(giftpath.read_bytes()).hexdigest(),'rows':len(matrix),'columns':cols,'invoices':16,'chunks':[{'file':f'{sid}-0.json','start':0,'rows':len(matrix)}]})
    save('manifest.json',{'spreadsheetId':TARGET,'sheets':manifest});save('init.json',init);save('styles.json',styles)
    print(json.dumps({'sheets':[{k:s[k] for k in ['sheetId','title','rows','columns','invoices']} for s in manifest],'chunks':sum(len(s['chunks']) for s in manifest),'styles':len(styles)},ensure_ascii=False))
if __name__=='__main__':main()
