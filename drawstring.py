from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm ,cm
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas
import arabic_reshaper
from bidi.algorithm import get_display

arabic_text= "انا انس "
rehaped_text = arabic_reshaper.reshape(arabic_text)
bidi_text = get_display(rehaped_text)
registerFont(TTFont('Arial','ARIAL.ttf'))



page = Canvas("test.pdf", pagesize=A4)
page.setFont('Arial', 12)
page.drawString(0*cm, 0*cm, bidi_text) 
page.showPage()
page.save()