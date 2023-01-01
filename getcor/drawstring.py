from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm ,cm
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas
import arabic_reshaper
from bidi.algorithm import get_display
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.lib.pagesizes import  A4

arabic_text= "انا انس "
rehaped_text = arabic_reshaper.reshape(arabic_text)
bidi_text = get_display(rehaped_text)
registerFont(TTFont('Arial','ARIAL.ttf'))

# 151.0   108.0
outfile = "test1.pdf"
template = PdfReader("document-page1.pdf", decompress=False).pages[0]
template_obj = pagexobj(template)
page = Canvas(outfile , pagesize=A4)
xobj_name = makerl(page, template_obj)
page.doForm(xobj_name)

# احداثيات اخذتها من البرنامج getcoor.py 
l = [137.0,98.5],[59.5,98.5],[136.0,79.5],[57.5,79.5],[127.0,62]
l = [126.0,95.5],[61.0,96.0],[133.0,77.5],[59.0,77.5],[127.0,58.0]
l = [130.0,98.0],[61.5,97.5],[131.0,79.0],[59.0,79.0],[130.0,60.0]
# page = Canvas("document-page1.pdf")
page.setFont('Arial', 12)
for i in l :
    page.drawString(i[0]*mm, i[1]*mm, bidi_text) 
page.showPage()
page.save()

# page.build(flowables, onFirstPage=onFirstPage)