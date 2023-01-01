from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfform
from reportlab.lib.colors import transparent ,magenta ,pink ,blue
from reportlab.pdfgen.canvas import Canvas
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.lib.pagesizes import landscape , A4
from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.pdfbase.ttfonts import TTFont

registerFont(TTFont('Arial','ARIAL.ttf'))

def create_simple_form(name1 , nums, start , stop):
    outfile = f"{name1}.pdf"
    template = PdfReader("evaluation.pdf", decompress=False).pages[0]
    template_obj = pagexobj(template)
    canvas = Canvas(outfile , pagesize=landscape(A4))
    xobj_name = makerl(canvas, template_obj)
    canvas.doForm(xobj_name)
    canvas.setFont("Arial", 2 )
    form = canvas.acroForm
    x = 60
    w = 20
    w1 = 90
    h = 10
    vspace = 412
    border = 5   
    border1 = 5   
    border2 = 15
    oddones =[i for i in range(4,14,2)]
    oddones1 =[i for i in range(14,41,2)]
    oddones2 = [i for i in range(10,22,2)]
    for i in range(start , stop+1):
        form.textfield(name=str(i), tooltip='num',
                    x=x-27, y=vspace, width=w , height=h,
                    borderColor= transparent,
                    fillColor=transparent , borderWidth=0,borderStyle='solid',forceBorder=False  )

        form.textfield(name='name'+str(i), tooltip='name',
                    x=x, y=vspace, width=w1 , height=h,
                        borderColor= transparent,
                    fillColor=transparent ,borderWidth=0,borderStyle='solid',forceBorder=False )
        vspace -= 14
        # if i % 2 == 0 :
        #     vspace -=15 + border1
        #     border1 +=1
        # else:
        #     vspace-=15

        # if i in oddones:
        #     vspace+=border
        #     border +=2

        # if i in oddones1:
        #     vspace+=border2
        #     # border2 +=5       
        # if i in oddones2:
        #     vspace -=5

        # nums=[i for i in range(32 , 41 ,2)]
        # if i in nums:
        #     vspace+=5
        nums2 =nums
        
        if i in nums2:
            vspace+=3
    canvas.save()

nums2=[4 ,10,15,19,18]
create_simple_form('evaluation2' ,nums2, 1 , 25)

nums2=[29 ,35,40,44,43]
create_simple_form('evaluation3',nums2, 26 , 50)
# for page in range(8,29,2):
#     create_simple_form(page)