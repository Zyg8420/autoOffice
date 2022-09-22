from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

doc = Document('file/春晓.docx')
for para in doc.paragraphs:
    for run in para.runs:
        # 字体加粗
        run.font.bold = True
        # 字体设置为斜体
        run.font.italic = True
        # 字体下划线
        run.font.underline = True
        # 设置划线
        # run.font.strike = True
        # 设置字体大小未24号字体
        run.font.size = Pt(24)
        # 设置字体颜色
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.name = '等线'
        r = run._element.rPr.rFonts
        r.set(qn('w:eastAsia'), '等线')

doc.save('file/春晓2.docx')
