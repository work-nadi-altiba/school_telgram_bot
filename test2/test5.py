# from pathlib import Path 
# from docxtpl import DocxTemplate
# import pandas as pd

# document_path = Path(__file__).parent / 'test.docx'
# doc = DocxTemplate(document_path)
# context= {  "1":"","name1":"", "2":"","name2":"", "3":"","name3":"", "4":"","name4":"", "5":"","name5":"", "6":"","name6":"", "7":"","name7":"", "8":"","name8":"", "9":"","name9":"", "10":"","name10":"", "11":"","name11":"", "12":"","name12":"", "13":"","name13":"", "14":"","name14":"", "15":"","name15":"", "16":"","name16":"", "17":"","name17":"", "18":"","name18":"", "19":"","name19":"", 
# "20":"","name20":"", "21":"","name21":"", "22":"","name22":"", "23":"","name23":"", "24":"","name24":"", "25":"","name25":"", "26":"","name26":"", "27":"","name27":"", "28":"","name28":"", "29":"","name29":"", "30":"","name30":"", "31":"","name31":"", "32":"","name32":"", "33":"","name33":"", "34":"","name34":"", "35":"","name35":"", "36":"","name36":"", "37":"","name37":"", "38":"","name38":"", "39":"","name39":"", "40":"","name40":"", "41":"","name41":"", "42":"","name42":"", "43":"","name43":"", "44":"","name44":"", "45":"","name45":"", "46":"","name46":"", "47":"","name47":"", "48":"","name48":"", "49":"","name49":"", "50":"","name50":"", "51":"","name51":"", "52":"","name52":"", "53":"","name53":"", "54":"","name54":"", "55":"","name55":"",'grade' : ''}




# df = pd.read_csv("MySchoolStudents.csv")

# def sort_csv():
#     df = pd.read_csv("MySchoolStudents.csv")
#     result = df['الصف و الشعبة'].unique()
#     # names = result['اسم الطالب'].tolist()
#     return sorted(result)

# def get_names(classname):
#     result = df[df['الصف و الشعبة'] == str(classname)]
#     names = result['اسم الطالب'].tolist()
#     return names

# grades = sort_csv()
# # names = get_names(grades[1])
# # print(names)


# def fill_print( grade ):
#     names = get_names(grade)
#     counter = 1
#     # print(names)
#     # input('press anything')
#     for name in names :
#         context[str(counter)] = str(counter)
#         context[f'name{counter}'] = str(name) 
#         counter+=1

#     context['grade'] = str(grade)
#     doc.render(context)
#     doc.save(Path(__file__).parent /f'{grade}.docx')
#     # input('press anything')


# for grade in grades:
#     fill_print(grade)
#     print(grade)
#     context= {  "1":"","name1":"", "2":"","name2":"", "3":"","name3":"", "4":"","name4":"", "5":"","name5":"", "6":"","name6":"", "7":"","name7":"", "8":"","name8":"", "9":"","name9":"", "10":"","name10":"", "11":"","name11":"", "12":"","name12":"", "13":"","name13":"", "14":"","name14":"", "15":"","name15":"", "16":"","name16":"", "17":"","name17":"", "18":"","name18":"", "19":"","name19":"", 
# "20":"","name20":"", "21":"","name21":"", "22":"","name22":"", "23":"","name23":"", "24":"","name24":"", "25":"","name25":"", "26":"","name26":"", "27":"","name27":"", "28":"","name28":"", "29":"","name29":"", "30":"","name30":"", "31":"","name31":"", "32":"","name32":"", "33":"","name33":"", "34":"","name34":"", "35":"","name35":"", "36":"","name36":"", "37":"","name37":"", "38":"","name38":"", "39":"","name39":"", "40":"","name40":"", "41":"","name41":"", "42":"","name42":"", "43":"","name43":"", "44":"","name44":"", "45":"","name45":"", "46":"","name46":"", "47":"","name47":"", "48":"","name48":"", "49":"","name49":"", "50":"","name50":"", "51":"","name51":"", "52":"","name52":"", "53":"","name53":"", "54":"","name54":"", "55":"","name55":"",'grade' : ''}

import winsound
duration = 1000  # milliseconds
freq = 440  # Hz
winsound.Beep(freq, duration)
winsound.Beep(freq, duration)
winsound.Beep(freq, duration)
winsound.Beep(freq, duration)