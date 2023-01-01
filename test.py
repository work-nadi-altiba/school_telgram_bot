# import pandas as pd
# # # # from pandas import *
# xls_file = '6h.csv'
# df = pd.read_csv(xls_file)

# # # # .to_string(index=False) 
# # result = df[df['الصف و الشعبة']=='الصف السادس-د'].to_string(index=False) 
# sorted_df = df.sort_values(by=["اسم الطالب"], ascending=True)
# sorted_df.to_csv(xls_file, index=False)

# n = df['اسم الطالب'].tolist()

# print('names:', n)

# # # print(result)

# 0;1.;1|0;Text Box 1.;انس محمود سمير الجعافرة
# counter = 1
# for i in range(3):
#     # print(i)
#     for x in range(1,14):
#         print( i , f'{x}.' ,str(counter) + '|'+ str(i) ,f'Text Box {x}.' ,'anas' ,sep=';' , end='|')
#         counter+=1
    # print(i , end='|')


# ==========================================================

from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import fillpdf
from fillpdf import fillpdfs

def read_csv():
    df = pd.read_csv("7dal.csv")
    result = df[df['الصف و الشعبة']=='الصف السادس-د'].to_string(index=False) 
    names = df['اسم الطالب'].tolist()
    return names

# anas4.pdf page_data
page_data = {'name5': '', 'name3': '', 'name6': '', 'name4': '', 'name2': '', 'name1': '', 'name7': '', 'name8': '', 'name9': '', 'name10': '', 'name11': '', 
'name12': '', 'name13': '', 'name14': '', 'name15': '', 'name16': '', 'name17': '', '1': '', '2': '', '3': '', '4': '', '5': '', '6': '', '7': '', '8': '', '9': '', '10': '', '11': '', '12': '', '13': '', '14': '', '15': '', '16': '', '17': ''}

# zaid-3.pdf page_data
page_data = {'name1':'' ,1:'' , 'name2':'' ,2:'' , 'name3':'' ,3:'' , 'name4':'' ,4:'' , 'name5':'' ,5:'' , 'name6':'' ,6:'' , 'name7':'' ,7:'' , 'name8':'' ,8:'' , 'name9':'' ,9:'' , 'name10':'' ,10:'' , 'name11':'' ,11:'' , 'name12':'' ,12:'' , 'name13':'' ,13:'' , 'name14':'' ,14:'' , 'name15':'' ,15:'' , 'name16':'' ,16:'' , 'name17':'' ,17:'' , 'name18':'' ,18:'' , 'name19':'' ,19:'' , 'name20':'' ,20:'' , 'name21':'' ,21:'' , 'name22':'' ,22:'' , 'name23':'' ,23:'' , 'name24':'' ,24:'' , 'name25':'' ,25:'' , 'name26':'' ,26:'' , 'name27':'' ,27:'' , 'name28':'' ,28:'' , 'name29':'' ,29:'' , 'name30':'' ,30:'' , 'name31':'' ,31:'' , 'name32':'' ,32:'' , 'name33':'' ,33:'' , 'name34':'' ,34:''}  

names = read_csv()
# # counter = 0

# # print(names)

# # for name in names:
# #     counter+=1
# #     page_data[str(counter)] = str(counter)
# #     page_data[f'name{counter}'] = str(name)

# # print(page_data)
# # print(names[0])
# # for name in names : 
# out_num =0
# counter = 0
# try: 
#     for name in range(len(names)):
#         for i in range(1 , 18):
#             i2 =counter +1 
#             print(i ,i2,names[counter])
#             page_data[str(i)] = i2
#             page_data['name'+str(i)] = str(names[counter])
#             counter += 1
#             # print(page_data['Text Box 1'])
#         fillpdfs.write_fillable_pdf('zaid-4.pdf', f'out{out_num}.pdf', page_data, flatten=False)
# #         page_data = {'name5': '', 'name3': '', 'name6': '', 'name4': '', 'name2': '', 'name1': '', 'name7': '', 'name8': '', 'name9': '', 'name10': 
# # '', 'name11': '', 'name12': '', 'name13': '', 'name14': '', 'name15': '', 'name16': '', 'name17': '', '1': '', '2': '', '3': '', '4': '', '5': '', '6': '', '7': '', '8': '', '9': '', '10': '', '11': '', '12': '', '13': '', '14': '', '15': '', '16': '', '17': ''}
#         out_num += 1
#         print('anas' )
#         input('press anything please')
#             # print_form_fields(input_pdf_path, sort=False, page_number=None)
#             # fillpdfs.print_form_fields('anas3.pdf', page_number=1)
# except IndexError  as e: 
#     fillpdfs.write_fillable_pdf('zaid-4_zaid-4_merged-1.pdf' , f'out{out_num}.pdf' , page_data, flatten=True)
#     print(e)

# from PyPDF2 import PdfFileWriter, PdfFileReader , PdfFileMerger

                                  
# merger = PdfFileMerger()

# merger.append(PdfFileReader(open(f"out0.pdf", 'rb')))
# merger.append(PdfFileReader(open(f"out1.pdf", 'rb')))
# # merger.append(PdfFileReader(open(f"out1.pdf", 'rb')))

# merger.write("out_merged.pdf")

#         # input('press any')
# # print(fillpdfs.get_form_fields("zaid-3.pdf"))

# # page_data = {'Text Box 1': '', '1': '', '2': '', 'Text Box 2': '', '3': '', 'Text Box 3': '', '4': '', 'Text Box 4': '', '5': '', 'Text Box 5': '', '6': '', 'Text Box 6': '', 'Text Box 7': '', '7': '', 'Text Box 8': '', '8': '', 'Text Box 9': '', '9': '', 'Text Box 10': '', '10': '', '11': '', 'Text Box 11': '', 'Text Box 12': '', '12': '', 'Text Box 13': '', '13': '', 'Text Box 14': '', '14': ''}
# # page_data['1'] = 1
# # page_data['Text Box 1'] = 'anas mahmoud aljaafreh'
# # # print(page_data['Text Box 1'])
# # fillpdfs.write_fillable_pdf('anas3.pdf', 'out.pdf', page_data, flatten=False)
# # print_form_fields(input_pdf_path, sort=False, page_number=None)
# # fillpdfs.print_form_fields('anas3.pdf', page_number=1)

# # ===========================================

# # import json

# # def readFile():
# #     with open("pdf.json" , 'r' ,encoding="utf8") as f:
# #         cont = f.readlines()
# #     cont =''.join(cont)
# #     return cont

# # pdfData= json.loads(f'{readFile()}')
# # for i in range(len(pdfData['info']['FieldsInfo']['Fields'])):
# #     print(pdfData['info']['FieldsInfo']['Fields'][i]['PageIndex'],pdfData['info']['FieldsInfo']['Fields'][i]['FieldName'])
# #     input('press anything')

# # import pandas as pd
# # df = pd.read_csv("6d.csv")
# # result = df[df['الصف و الشعبة']=='الصف السادس-د'].to_string(index=False) 
# # n = df['اسم الطالب'].tolist()
# # print('names:', n)

out_num =0
counter = 0
for name in names:
    i2 =counter +1 
    print(i2,names[counter])
    page_data[str(i2)] = i2
    page_data['name'+str(i2)] = str(names[counter])
    counter += 1

fillpdfs.write_fillable_pdf('zaid-4_zaid-4_merged-1.pdf' , f'out{out_num}.pdf' , page_data, flatten=True)