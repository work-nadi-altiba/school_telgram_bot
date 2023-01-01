from PyPDF2  import PdfFileWriter ,PdfFileReader,  WaterMark
import os

tmp_name = "__tmp.pdf"

output_file = PdfFileWriter()

with open(inFile, 'rb') as f:
    # Read the pdf (create a pdf stream)
    pdf_original = PdfFileReader(f, strict=False)
    # put all buffer in a single file
    output_file.appendPagesFromReader(pdf_original)
    # create new PDF with water mark
    WaterMark._page(fixPage, tmp_name)
    # Open the created pdf
    with open(tmp_name, 'rb') as ftmp:
        # Read the temp pdf (create a pdf stream obj)
        temp_pdf = PdfFileReader(ftmp)
        for p in range(startPage, startPage+pages):
            original_page = output_file.getPage(p)
            temp_page = temp_pdf.getPage(0)
            original_page.mergePage(temp_page)

        # write result
        if output_file.getNumPages():
            # newpath = inFile[:-4] + "_numbered.pdf"
            with open(outFile, 'wb') as f:
                output_file.write(f)
    os.remove(tmp_name)