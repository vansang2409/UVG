import json
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

ROOT = Path(__file__).parent
SOURCE = json.loads((ROOT / 'source.json').read_text(encoding='utf-8'))
OUT = ROOT / 'TT-111725 - Hợp đồng sửa chữa.docx'

def v(row, col=0):
    return SOURCE['cells'][row - 1][col]['value']

def joined(first, last):
    return ' '.join(v(r).strip() for r in range(first, last + 1))

doc = Document()
section = doc.sections[0]
section.page_width = Mm(210)
section.page_height = Mm(297)
section.left_margin = Mm(21)
section.right_margin = Mm(21)
section.top_margin = Mm(9)
section.bottom_margin = Mm(16)
section.header_distance = Mm(0)
section.footer_distance = Mm(0)

normal = doc.styles['Normal']
normal.font.name = 'Times New Roman'
normal.font.size = Pt(10)
normal.font.color.rgb = __import__('docx').shared.RGBColor(0, 0, 0)
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.line_spacing = Pt(11.5)
normal.paragraph_format.widow_control = True

title = doc.styles['Title']
title.font.name = 'Times New Roman'
title.font.size = Pt(22.5)
title.font.bold = True
title.font.color.rgb = __import__('docx').shared.RGBColor(0, 0, 0)
title.paragraph_format.space_before = Pt(0)
title.paragraph_format.space_after = Pt(1)
title.paragraph_format.line_spacing = Pt(26)

def font(run, size=None, bold=None, italic=None):
    run.font.name = 'Times New Roman'
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic

def paragraph(text='', *, bold=False, italic=False, size=None, before=0, after=0, align=None, label=None, style=None):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    if align is not None:
        p.alignment = align
    if label and text.startswith(label):
        font(p.add_run(label), size, True, italic)
        font(p.add_run(text[len(label):]), size, False, italic)
    else:
        font(p.add_run(text), size, bold, italic)
    return p

def table_setup(table, widths, border=False):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for column, width in zip(table.columns, widths):
        column.width = Mm(width)
    props = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single' if border else 'nil')
        e.set(qn('w:sz'), '6')
        e.set(qn('w:color'), '000000')
        borders.append(e)
    props.insert_element_before(borders, 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription')
    margins = OxmlElement('w:tblCellMar')
    for edge, value in [('top', 30 if border else 0), ('left', 45 if border else 0), ('bottom', 30 if border else 0), ('right', 45 if border else 0)]:
        e = OxmlElement('w:' + edge)
        e.set(qn('w:w'), str(value))
        e.set(qn('w:type'), 'dxa')
        margins.append(e)
    props.insert_element_before(margins, 'w:tblLook', 'w:tblCaption', 'w:tblDescription')
    for row in table.rows:
        cannot_split = OxmlElement('w:cantSplit')
        row._tr.get_or_add_trPr().append(cannot_split)
        for cell, width in zip(row.cells, widths):
            cell.width = Mm(width)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def cell_text(cell, text, *, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=10):
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(11.5)
    for i, line in enumerate(str(text).split('\n')):
        if i:
            p.add_run().add_break()
        font(p.add_run(line), size, bold)

paragraph(v(1), style='Title', bold=True)
paragraph(v(2), bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.RIGHT)
paragraph(v(3), bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.RIGHT, after=1)

for row in range(4, 14):
    paragraph(v(row), bold=row in (4, 5, 9, 10))

paragraph(joined(15, 16), before=8, after=2)
ports = doc.add_table(rows=1, cols=2)
table_setup(ports, [90, 78])
cell_text(ports.cell(0, 0), v(17, 1), align=WD_ALIGN_PARAGRAPH.CENTER)
cell_text(ports.cell(0, 1), v(17, 4), align=WD_ALIGN_PARAGRAPH.CENTER)
paragraph(v(18), bold=True, after=7)

goods = doc.add_table(rows=4, cols=6)
table_setup(goods, [10, 71, 14, 14, 29, 30], border=True)
quantity_header = goods.cell(0, 2).merge(goods.cell(0, 3))
headers = [(goods.cell(0, 0), v(20, 0)), (goods.cell(0, 1), v(20, 1)), (quantity_header, v(20, 2)), (goods.cell(0, 4), v(20, 4)), (goods.cell(0, 5), v(20, 5))]
for cell, text in headers:
    cell_text(cell, text, bold=True)
delivery = goods.cell(1, 0).merge(goods.cell(1, 5))
cell_text(delivery, v(21), bold=True)
for c, text in enumerate(['1', v(22, 1), '1.00', v(22, 3), '4,000.00', '4,000.00']):
    align = WD_ALIGN_PARAGRAPH.LEFT if c in (1, 3) else WD_ALIGN_PARAGRAPH.RIGHT
    cell_text(goods.cell(2, c), text, align=align, size=9.5 if c == 1 else 10)
total = goods.cell(3, 0).merge(goods.cell(3, 4))
cell_text(total, v(23), bold=True)
cell_text(goods.cell(3, 5), '4,000.00', bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)

paragraph(v(24), bold=True, italic=True, before=2, after=2)
paragraph(v(25), label='2. Terms of payment:', after=2)
paragraph(v(26), label='3. Shipment:', after=2)
p = paragraph(v(27).strip(), label='Re - import time :', after=2)
p.paragraph_format.left_indent = Mm(3)
paragraph(v(28), label='4. Term of delivery:', after=2)
paragraph(joined(29, 32), label='5. Force Majeure:', after=3)
paragraph(joined(33, 34), label='6. Default:', after=3)
paragraph(joined(35, 40), label='7. Claim:', after=8)

signatures = doc.add_table(rows=2, cols=2)
table_setup(signatures, [90, 78])
cell_text(signatures.cell(0, 0), v(42).strip(), bold=True)
cell_text(signatures.cell(0, 1), v(42, 3), bold=True)
cell_text(signatures.cell(1, 0), v(43), bold=True, size=9)
cell_text(signatures.cell(1, 1), v(43, 3), bold=True, size=9)
doc.core_properties.title = v(1) + ' ' + v(2)
doc.core_properties.subject = 'Repair Contract'
doc.core_properties.author = ''
doc.core_properties.last_modified_by = ''
# The bundled Word default template contains an ornamental title rule.
# Remove paragraph borders to match the supplied contract's plain title.
for element in list(doc.styles.element.iter(qn('w:pBdr'))):
    element.getparent().remove(element)
for element in list(doc.element.iter(qn('w:pBdr'))):
    element.getparent().remove(element)
for element in doc.settings.element.iter(qn('w:zoom')):
    element.set(qn('w:percent'), '100')
doc.save(OUT)

# Check every source text cell against the editable Word XML, ignoring only
# whitespace changes caused by reflowing the original spreadsheet's lines.
from zipfile import ZipFile
from lxml import etree
with ZipFile(OUT) as archive:
    root = etree.fromstring(archive.read('word/document.xml'))
    extracted = ' '.join(root.xpath('//w:t/text()', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
normalize = lambda s: re.sub(r'\s+', ' ', s).strip()
corpus = normalize(extracted)
missing = []
for row in SOURCE['cells']:
    for cell in row:
        if isinstance(cell['value'], str) and cell['value'].strip():
            if normalize(cell['value']) not in corpus:
                missing.append(cell['value'])
assert not missing, missing
print('Saved:', OUT)
print('All source text cells preserved. Quantity: 1 SET. Unit price and total: USD 4,000.00.')
