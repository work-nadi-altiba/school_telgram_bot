from PyPDF2 import PdfFileWriter, PdfFileReader , PdfFileMerger

# inputpdf = PdfFileReader(open("record2.pdf", "rb"))

# for i in range(inputpdf.numPages):
#     output = PdfFileWriter()
#     output.addPage(inputpdf.getPage(i))
#     with open("document-page%s.pdf" % str(i+1), "wb") as outputStream:
#         output.write(outputStream)

# import glob , re
# files = glob.glob("*.pdf")
# for i in files:
    # if i == 'document-page8.pdf':
    #     print('found')
    # if i == 'document-page10.pdf':
    #     print('found')
    # if i == 'document-page12.pdf':
    #     print('found')
    # if i == 'document-page14.pdf':
    #     print('found')
    # if i == 'document-page18.pdf':
    #     print('found')        
    # if i == 'document-page20.pdf':
    #     print('found')
    # if i == 'document-page22.pdf':
    #     print('found')
    # if i == 'document-page24.pdf':
    #     print('found') 
    # if i == 'document-page26.pdf':
    #     print('found')     
    # if i == 'document-page28.pdf':
    #     print('found')                                    
merger = PdfFileMerger()
# # for file in files 
# merger.append(PdfFileReader(open('document-page28.pdf', 'rb')))
# merger.append(PdfFileReader(open('document-page29.pdf', 'rb')))

# merger.write("merged.pdf")


# for i in range( 1 , 41):
#     replace=[x for x in range(8,29,2)]  
#     if i in replace:
#         # print(f'result{i}.pdf')   
#         merger.append(PdfFileReader(open(f'anas.pdf', 'rb')))
#         # print(f"merger.append(PdfFileReader(open(f\"result{i}.pdf\", 'rb')))")
#     elif i == 34:
#         merger.append(PdfFileReader(open(f'anas.pdf', 'rb')))
#     else:
#         # print(f"document-page{i}.pdf")
#         # print(f"merger.append(PdfFileReader(open(f\"document-page{i}.pdf\", 'rb')))")
#         merger.append(PdfFileReader(open(f"anas.pdf", 'rb')))
#     # print(f'document-page{i}.pdf')

    # print(f"merger.append(PdfFileReader(open(f\"document-page{i}.pdf\", 'rb')))")
merger.append(PdfFileReader(open(f"anas.pdf", 'rb')))
merger.append(PdfFileReader(open(f"anas.pdf", 'rb')))
merger.append(PdfFileReader(open(f"anas.pdf", 'rb')))

merger.write("merged.pdf")