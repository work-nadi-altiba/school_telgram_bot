import requests
import json
from pygments import highlight
from pygments.lexers import JsonLexer
from pygments.formatters import TerminalFormatter
from openpyxl import Workbook , load_workbook
from docxtpl import DocxTemplate
from docx2pdf import convert
import subprocess
import os 
import glob
from odf.opendocument import load
from odf import text, table
from odf.table import Table ,TableCell ,TableRow
from odf.namespaces import DRAWNS, STYLENS, SVGNS, TEXTNS
from odf.draw import CustomShape
from odf.style import Style, TextProperties
from num2words import num2words
from odf.text import P, Span
import hijri_converter
import fitz
import webcolors
import ezodf
import shutil
import PyPDF2
import PyPDF4
import os
from openpyxl import load_workbook
from openpyxl.styles import Font
import random
import re
import itertools
import openpyxl

def create_tables(auth , grouped_list):
    # auth = get_auth(username , password)
    institution_area_data = inst_area(auth)
    institution_data = inst_name(auth)
    curr_year_code = get_curr_period(auth)['data'][0]['code']


    for group in grouped_list:
        
        template_file = openpyxl.load_workbook('tamplete.xlsx')
        marks_sheet = template_file.worksheets[2]

        for row_number, dataFrame in enumerate(group, start=4):
            islam_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'سلامية' in key] # التربية الاسلامية
            arabic_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'عربية' in key] # اللغة العربية
            english_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'جليزية' in key]   # اللغة الانجليزية 
            math_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'رياضيات' in key] # الرياضيات 
            social_subjects = [value for key ,value in dataFrame['subject_sums'].items() if 'اجتماعية و الوطنية' in key]   # التربية الاجتماعية و الوطنية 
            science_subjects = [value for key ,value in dataFrame['subject_sums'].items() if 'العلوم' in key]  # العلوم
            art_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'الفنية والموس' in key]    # التربية الفنية والموسيقية
            sport_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'رياضية' in key] # التربية الرياضية
            vocational_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'مهنية' in key] # التربية المهنية 
            computer_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'حاسوب' in key]   # الحاسوب
            financial_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'مالية' in key]  # الثقافة المالية
            franch_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'فرنسية' in key]    # اللغة الفرنسية 
            christian_subject = [value for key ,value in dataFrame['subject_sums'].items() if 'الدين المسيحي' in key]  # الدين المسيحي

            marks_sheet.cell(row=row_number, column=1).value = row_number-3
            marks_sheet.cell(row=row_number, column=2).value = dataFrame['student__full_name']
            marks_sheet.cell(row=row_number, column=3).value = dataFrame['student_nat']
            marks_sheet.cell(row=row_number, column=4).value = dataFrame['student_birth_place']
            marks_sheet.cell(row=row_number, column=5).value = dataFrame['student_birth_date'].split('/')[0]
            marks_sheet.cell(row=row_number, column=6).value = dataFrame['student_birth_date'].split('/')[1]
            marks_sheet.cell(row=row_number, column=7).value = dataFrame['student_birth_date'].split('/')[2]
            marks_sheet.cell(row=row_number, column=8).value = islam_subject[0][0] if islam_subject and len(islam_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=9).value = islam_subject[0][1] if islam_subject and len(islam_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=10).value = islam_subject[0][0]+islam_subject[0][1] if islam_subject and len(islam_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=11).value = arabic_subject[0][0] if arabic_subject and len(arabic_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=12).value = arabic_subject[0][1] if arabic_subject and len(arabic_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=13).value = arabic_subject[0][0]+arabic_subject[0][1] if arabic_subject and len(arabic_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=14).value = english_subject[0][0] if english_subject and len(english_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=15).value = english_subject[0][1] if english_subject and len(english_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=16).value = english_subject[0][0]+english_subject[0][1] if english_subject and len(english_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=17).value = math_subject[0][0] if math_subject and len(math_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=18).value = math_subject[0][1] if math_subject and len(math_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=19).value = math_subject[0][0]+math_subject[0][1] if math_subject and len(math_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=20).value = social_subjects[0][0] if social_subjects and len(social_subjects[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=21).value = social_subjects[0][1] if social_subjects and len(social_subjects[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=22).value = social_subjects[0][0]+social_subjects[0][1] if social_subjects and len(social_subjects[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=23).value = science_subjects[0][0] if science_subjects and len(science_subjects[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=24).value = science_subjects[0][1] if science_subjects and len(science_subjects[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=25).value = science_subjects[0][0]+science_subjects[0][1] if science_subjects and len(science_subjects[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=26).value = art_subject[0][0] if art_subject and len(art_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=27).value = art_subject[0][1] if art_subject and len(art_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=28).value = art_subject[0][0]+art_subject[0][1] if art_subject and len(art_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=29).value = sport_subject[0][0] if sport_subject and len(sport_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=30).value = sport_subject[0][1] if sport_subject and len(sport_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=31).value = sport_subject[0][0]+sport_subject[0][1] if sport_subject and len(sport_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=32).value = financial_subject[0][0] if financial_subject and len(financial_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=33).value = financial_subject[0][1] if financial_subject and len(financial_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=34).value = financial_subject[0][0]+financial_subject[0][1] if financial_subject and len(financial_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=35).value = vocational_subject[0][0] if vocational_subject and len(vocational_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=36).value = vocational_subject[0][1] if vocational_subject and len(vocational_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=37).value = vocational_subject[0][0]+vocational_subject[0][1] if vocational_subject and len(vocational_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=38).value = computer_subject[0][0] if computer_subject and len(computer_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=39).value = computer_subject[0][1] if computer_subject and len(computer_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=40).value = computer_subject[0][0]+computer_subject[0][1] if computer_subject and len(computer_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=41).value = franch_subject[0][0] if franch_subject and len(franch_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=42).value = franch_subject[0][1] if franch_subject and len(franch_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=43).value = franch_subject[0][0]+franch_subject[0][1] if franch_subject and len(franch_subject[0]) > 0 else ''
            # marks_sheet.cell(row=row_number, column=44).value = dataFrame[0][] if = and len(=[0]) > 0 else ''
            # marks_sheet.cell(row=row_number, column=45).value = dataFrame[0][] if = and len(=[0]) > 0 else ''
            # marks_sheet.cell(row=row_number, column=46).value = dataFrame[0][] if = and len(=[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=47).value = christian_subject[0][0] if christian_subject and len(christian_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=48).value = christian_subject[0][1] if christian_subject and len(christian_subject[0]) > 0 else ''
            marks_sheet.cell(row=row_number, column=49).value = christian_subject[0][0]+christian_subject[0][1] if christian_subject and len(christian_subject[0]) > 0 else ''

        if 'الثامن' in group[0]['student_grade_name']:
            marks_sheet['a1'] = 1800
            marks_sheet['a3'] =f'جدول العلامات المدرسيه للصف الثامن الأساسي للعام الدراسي ( {curr_year_code} )'
            # اسلامية
            # h/i/j
            marks_sheet['h3'],marks_sheet['i3'],marks_sheet['j3'] = [200]*3
            # عربية 
            # k/l/m
            marks_sheet['k3'],marks_sheet['l3'],marks_sheet['m3'] = [300]*3
            # انجليزية 
            # n/o/p
            marks_sheet['n3'],marks_sheet['o3'],marks_sheet['p3'] = [200]*3
            # رياضيات
            # q/r/s
            marks_sheet['q3'],marks_sheet['r3'],marks_sheet['s3'] = [200]*3
            # اجتماعيات 
            # t/u/v
            marks_sheet['t3'],marks_sheet['u3'],marks_sheet['v3'] = [200]*3
            # علوم
            # w/x/y
            marks_sheet['w3'],marks_sheet['x3'],marks_sheet['y3'] = [200]*3
        elif 'التاسع' in group[0]['student_grade_name']:
            marks_sheet['a1'] = 2000
            marks_sheet['a3'] =f'جدول العلامات المدرسيه للصف التاسع  الأساسي للعام الدراسي ( {curr_year_code} )'
            # اسلامية
            # h/i/j
            marks_sheet['h3'],marks_sheet['i3'],marks_sheet['j3'] = [200]*3
            # عربية 
            # k/l/m
            marks_sheet['k3'],marks_sheet['l3'],marks_sheet['m3'] = [300]*3
            # انجليزية 
            # n/o/p
            marks_sheet['n3'],marks_sheet['o3'],marks_sheet['p3'] = [200]*3
            # رياضيات
            # q/r/s
            marks_sheet['q3'],marks_sheet['r3'],marks_sheet['s3'] = [200]*3
            # اجتماعيات 
            # t/u/v
            marks_sheet['t3'],marks_sheet['u3'],marks_sheet['v3'] = [200]*3
            # علوم
            # w/x/y
            marks_sheet['w3'],marks_sheet['x3'],marks_sheet['y3'] = [400]*3
        elif 'العاشر' in group[0]['student_grade_name']:
            marks_sheet['a1'] = 2000
            marks_sheet['a3'] =f'جدول العلامات المدرسيه للصف العاشر الأساسي للعام الدراسي ( {curr_year_code} )'
            # اسلامية
            # h/i/j
            marks_sheet['h3'],marks_sheet['i3'],marks_sheet['j3'] = [200]*3
            # عربية 
            # k/l/m
            marks_sheet['k3'],marks_sheet['l3'],marks_sheet['m3'] = [300]*3
            # انجليزية 
            # n/o/p
            marks_sheet['n3'],marks_sheet['o3'],marks_sheet['p3'] = [200]*3
            # رياضيات
            # q/r/s
            marks_sheet['q3'],marks_sheet['r3'],marks_sheet['s3'] = [200]*3
            # اجتماعيات 
            # t/u/v
            marks_sheet['t3'],marks_sheet['u3'],marks_sheet['v3'] = [200]*3
            # علوم
            # w/x/y
            marks_sheet['w3'],marks_sheet['x3'],marks_sheet['y3'] = [400]*3        
        else:
            marks_sheet['a3'] = f'جدول العلامات الدراسية للصفوف من الأول الى السابع الأساسي ( {curr_year_code} )'
            if 'سابع' in group[0]['student_grade_name']:
                marks_sheet['a1'] = 1100
            elif 'سادس' in group[0]['student_grade_name']:
                marks_sheet['a1'] = 900
            else:
                marks_sheet['a1'] = 800
            
        marks_sheet['b3'] = institution_area_data['data'][0]['Areas']['name']
        marks_sheet['c3'] = ''
        marks_sheet['d3'] = institution_data['data'][0]['Institutions']['code_name']
        marks_sheet['e3'] = institution_area_data['data'][0]['AreaAdministratives']['name']
        marks_sheet['f3'] = group[0]['student_grade_name']
        marks_sheet['g3'] = group[0]['student_class_name_letter']
        
        template_file.save(' جدول '+group[0]['student_class_name_letter']+'.xlsx')
        
def create_certs(grouped_list):
    for group in grouped_list:
        
        template_file = load_workbook('a4_gray_cert.xlsx')
        sheet1 = template_file.worksheets[0]
        
        names_averages =  sort_dictionary_list_based_on(group)

        group = sort_dictionary_list_based_on(group ,simple=False)

        for row_number, dataFrame in enumerate(names_averages, start=5):
            sheet1.cell(row=row_number, column=2).value = dataFrame[1]
            sheet1.cell(row=row_number, column=4).value = dataFrame[0]
            
        counter = 1
        for group_item in group:
            
            sheet2 = template_file.copy_worksheet(template_file.worksheets[1])
            sheet2.title = str(counter)
            counter += 1
            sheet2.sheet_view.rightToLeft = True    
            sheet2.sheet_view.rightToLeft = True   

            img = openpyxl.drawing.image.Image('Pasted image.png')
            img.anchor = 'e2'
            sheet2.add_image(img)

            # group_item = grouped_list[0][0]
            # print(sheet2)
            sheet2['b7'] = group_item['student__full_name']
            # مكان و تاريخ الولادة
            sheet2['h7']= str(group_item['student_birth_place']) + ' ' + str(group_item['student_birth_date'])
            #الرقم الوطني
            sheet2['b9']= group_item['student_nat_id']
            #الجنسية
            sheet2['h9']= group_item['student_nat']
            #الصف و الشعبة 
            sheet2['b11']= group_item['student_class_name_letter']
            #المدرسة و رقمها الوطني
            sheet2['g11']= group_item['student_school_name']
            #المنطقة التعليمية 
            sheet2['b13']= group_item['student_edu_place']  
            #البلدة 
            sheet2['f13']= ''
            #اللواء
            sheet2['i13']= group_item['student_directory']
            # put the subjects cells inder here 
            i ='c,d,e,g,j,f'.split(',')
            r = range(18,32)

            value_item = 0

            # التربية الاسلامية
            islam_subject = [value for key ,value in group_item['subject_sums'].items() if 'سلامية' in key]
            if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                sheet2['C18'] = 200
                maxMark = 200
            else:
                sheet2['C18'] = 100
                maxMark = 100
            sheet2['D18'] = islam_subject[0][value_item] if islam_subject and len(islam_subject[0]) != 0 else ''
            # sheet2['E18'] = islam_subject[0][value_item]
            # sheet2['F18'] = islam_subject[0][value_item]
            sheet2['G18'] = convert_avarage_to_words(islam_subject[0][value_item]) if islam_subject else ''
            sheet2['J18'] = score_in_words(islam_subject[0][value_item],max_mark=maxMark) if islam_subject else ''

            # اللغة العربية
            arabic_subject = [value for key ,value in group_item['subject_sums'].items() if 'عربية' in key]
            if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                sheet2['C19'] = 300
                maxMark = 300
            else:
                sheet2['C19'] = 100
                maxMark = 100
            sheet2['D19'] = arabic_subject[0][value_item] if arabic_subject and len(arabic_subject[0]) != 0 else ''
            # sheet2['E19'] = arabic_subject[0][value_item]
            # sheet2['F19'] = arabic_subject[0][value_item]
            sheet2['G19'] = convert_avarage_to_words(arabic_subject[0][value_item]) if arabic_subject else ''
            sheet2['J19'] = score_in_words(arabic_subject[0][value_item],max_mark=maxMark) if arabic_subject else ''

            # اللغة الانجليزية 
            english_subject = [value for key ,value in group_item['subject_sums'].items() if 'جليزية' in key]
            if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                sheet2['C20'] = 200
                maxMark = 200
            else:
                sheet2['C20'] = 100
                maxMark = 100
            sheet2['D20'] = english_subject[0][value_item] if english_subject and len(english_subject[0]) != 0 else ''
            # sheet2['E20'] = english_subject[0][value_item]
            # sheet2['F20'] = english_subject[0][value_item]
            sheet2['G20'] = convert_avarage_to_words(english_subject[0][value_item]) if english_subject else ''
            sheet2['J20'] = score_in_words(english_subject[0][value_item],max_mark=maxMark) if english_subject else ''

            # الرياضيات 
            math_subject = [value for key ,value in group_item['subject_sums'].items() if 'رياضيات' in key]
            if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                sheet2['C21'] = 200
                maxMark = 200
            else:
                sheet2['C21'] = 100
                maxMark = 100
            sheet2['D21'] = math_subject[0][value_item] if math_subject and len(math_subject[0]) != 0 else ''
            # sheet2['E21'] = math_subject[0][value_item]
            # sheet2['F21'] = math_subject[0][value_item]
            sheet2['G21'] = convert_avarage_to_words(math_subject[0][value_item]) if math_subject else ''
            sheet2['J21'] = score_in_words(math_subject[0][value_item],max_mark=maxMark) if math_subject else ''

            # التربية الاجتماعية و الوطنية 
            social_subjects = [value for key ,value in group_item['subject_sums'].items() if 'اجتماعية و الوطنية' in key]
            if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                sheet2['C22'] = 200
                maxMark = 200
                sheet2['D22'] = int(social_subjects[0][value_item]*(2/3)) if social_subjects and len(social_subjects[0]) != 0 else ''
            elif 'سادس' in group_item['student_grade_name'] or 'سابع' in group_item['student_grade_name']:
                sheet2['D22'] = int(social_subjects[0][value_item]/3) if social_subjects and len(social_subjects[0]) != 0 else ''
                sheet2['C22'] = 100
                maxMark = 100                
            else:
                sheet2['C22'] = 100
                maxMark = 100
                
            # sheet2['D22'] = social_subjects[0][value_item] if social_subjects and len(social_subjects[0]) != 0 else ''
            # sheet2['E22'] = social_subjects[0][value_item]
            # sheet2['F22'] = social_subjects[0][value_item]
            sheet2['G22'] = convert_avarage_to_words(social_subjects[0][value_item]) if social_subjects else ''
            sheet2['J22'] = score_in_words(int(social_subjects[0][value_item]*(2/3)),max_mark=maxMark) if social_subjects else ''

            # العلوم
            science_subjects = [value for key ,value in group_item['subject_sums'].items() if 'العلوم' in key]
            if 'ثامن' in group_item['student_grade_name'] :
                sheet2['C23'] = 200
                maxMark = 200
            elif 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                sheet2['C23'] = 400
                maxMark = 400
            else:
                sheet2['C23'] = 100
                maxMark = 100
            sheet2['D23'] = science_subjects[0][value_item] if science_subjects and len(science_subjects[0]) != 0 else ''
            # sheet2['E23'] = science_subjects[0][value_item]
            # sheet2['F23'] = science_subjects[0][value_item]
            sheet2['G23'] = convert_avarage_to_words(science_subjects[0][value_item]) if science_subjects else ''
            sheet2['J23'] = score_in_words( science_subjects[0][value_item],max_mark=maxMark) if  science_subjects else ''

            # التربية الفنية والموسيقية
            art_subject = [value for key ,value in group_item['subject_sums'].items() if 'الفنية والموس' in key]
            sheet2['C24'] = 100 if art_subject and len(art_subject[0]) != 0 else ''
            sheet2['D24'] = art_subject[0][value_item] if art_subject and len(art_subject[0]) != 0 else ''
            # sheet2['E24'] = art_subject[0][value_item]
            # sheet2['F24'] = art_subject[0][value_item]
            sheet2['G24'] = convert_avarage_to_words(art_subject[0][value_item]) if art_subject else ''
            sheet2['J24'] = score_in_words(art_subject[0][value_item] ) if art_subject else ''

            # التربية الرياضية
            sport_subject = [value for key ,value in group_item['subject_sums'].items() if 'رياضية' in key]
            sheet2['C25'] = 100 if sport_subject and len(sport_subject[0]) != 0 else ''
            sheet2['D25'] = sport_subject[0][value_item] if sport_subject and len(sport_subject[0]) != 0 else ''
            # sheet2['E25'] = sport_subject[0][value_item]
            # sheet2['F25'] = sport_subject[0][value_item]
            sheet2['G25'] = convert_avarage_to_words(sport_subject[0][value_item]) if sport_subject else ''
            sheet2['J25'] = score_in_words(sport_subject[0][value_item] ) if sport_subject else ''

            # التربية المهنية 
            vocational_subject = [value for key ,value in group_item['subject_sums'].items() if 'مهنية' in key]
            sheet2['C26'] = 100 if vocational_subject and len(vocational_subject[0]) != 0 else ''
            sheet2['D26'] = vocational_subject[0][value_item] if vocational_subject and len(vocational_subject[0]) != 0 else ''
            # sheet2['E26'] = vocational_subject[0][value_item]
            # sheet2['F26'] = vocational_subject[0][value_item]
            sheet2['G26'] = convert_avarage_to_words(vocational_subject[0][value_item]) if vocational_subject else ''
            sheet2['J26'] = score_in_words(vocational_subject[0][value_item] ) if vocational_subject else ''

            # الحاسوب
            computer_subject = [value for key ,value in group_item['subject_sums'].items() if 'حاسوب' in key]
            sheet2['C27'] = 100 if computer_subject and len(computer_subject[0]) != 0 else ''
            sheet2['D27'] = computer_subject[0][value_item] if computer_subject and len(computer_subject[0]) != 0 else ''
            # sheet2['E27'] = computer_subject[0][value_item]
            # sheet2['F27'] = computer_subject[0][value_item]
            sheet2['G27'] = convert_avarage_to_words(computer_subject[0][value_item]) if computer_subject else ''
            sheet2['J27'] = score_in_words(computer_subject[0][value_item] ) if computer_subject else ''

            # الثقافة المالية
            financial_subject = [value for key ,value in group_item['subject_sums'].items() if 'مالية' in key]
            sheet2['C28'] = 100 if financial_subject and len(financial_subject[0]) != 0 else ''
            sheet2['D28'] = financial_subject[0][value_item] if financial_subject and len(financial_subject[0]) != 0 else ''
            # sheet2['E28'] = financial_subject[0][value_item]
            # sheet2['F28'] = financial_subject[0][value_item]
            sheet2['G28'] = convert_avarage_to_words(financial_subject[0][value_item]) if financial_subject else ''
            sheet2['J28'] = score_in_words(financial_subject[0][value_item] ) if financial_subject else ''

            # اللغة الفرنسية 
            franch_subject = [value for key ,value in group_item['subject_sums'].items() if 'فرنسية' in key]
            sheet2['C29'] = 100 if franch_subject and len(franch_subject[0]) != 0 else ''
            sheet2['D29'] = franch_subject[0][value_item] if franch_subject and len(franch_subject[0]) != 0 else ''
            # sheet2['E29'] = franch_subject[0][value_item]
            # sheet2['F29'] = franch_subject[0][value_item]
            sheet2['G29'] = convert_avarage_to_words(franch_subject[0][value_item]) if franch_subject else ''
            sheet2['J29'] = score_in_words(franch_subject[0][value_item] ) if franch_subject else ''

            # الدين المسيحي
            christian_subject = [value for key ,value in group_item['subject_sums'].items() if 'الدين المسيحي' in key]
            sheet2['C30'] = 100 if christian_subject and len(christian_subject[0]) != 0 else ''
            sheet2['D30'] = christian_subject[0][value_item] if christian_subject and len(christian_subject[0]) != 0 else ''
            # sheet2['E30'] = christian_subject[0][value_item]
            # sheet2['F30'] = christian_subject[0][value_item]
            sheet2['G30'] = convert_avarage_to_words(christian_subject[0][value_item]) if christian_subject else ''
            sheet2['J30'] = score_in_words(christian_subject[0][value_item] ) if christian_subject else ''

            #المعدل المئوي بالرقام 
            sheet2['c32']= group_item['t1+t2+year_avarage'][0]
            #بالحروف
            sheet2['e32']= convert_avarage_to_words(group_item['t1+t2+year_avarage'][0]) if group_item else ''
            #ترتيب الطالب على الصف 
            sheet2['j32']= counter-1

            #النتيجة 
            sheet2['b33']= 'مقصر' if any(item < 49 for item in [value[0] for key , value in group_item['subject_sums'].items()] ) else score_in_words(int(group_item['t1+t2+year_avarage'][0]))
            #عدد ايام غياب الطالب 
            sheet2['c35']= ''
            #عدد ايام الدوام الرسمي الكامل 
            sheet2['g35']= ''
            #اسم و توقيع مربي الصف 
            sheet2['j35']= ''
            #التاريخ
            sheet2['b36']= ''
            #اسم و توقيع مدير المدرسة
            sheet2['i36']= ''
        template_file.remove(template_file['sheet'])
        template_file.save(group[0]['student_class_name_letter']+'.xlsx')

def sort_dictionary_list_based_on(dictionary_list , key='t1+t2+year_avarage',item=0, reverse=True , simple=True):
    if simple :
        return [(i['t1+t2+year_avarage'][item], i['student__full_name'] )for i in sorted(dictionary_list, key=lambda x: x[key][item] , reverse=reverse)]
    else:
        return sorted(dictionary_list, key=lambda x: x[key][item] , reverse=reverse)

def convert_avarage_to_words(digit):
    number_fraction = str(digit).split('.')
    if '.' in str(digit):
        number_in_words = re.sub("ريال.*", "",num2words(int(number_fraction[0]),lang='ar', to='currency'))
        fraction = int(number_fraction[1])
        if fraction == 1:
            fraction_in_words = 'عشر'
        elif fraction == 2:
            fraction_in_words = 'عشرين'
        else:
            fraction_in_words = str(num2words(fraction,lang='ar', to='year')) + ' اعشار'
        lst = ["مائة", "مئتان", "ثلاثمائة", "أربعمائة", "خمسمائة", "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة"]
        for item in lst:
            if item in number_in_words:
                word = number_in_words.replace(item, item + ' ').split()
                if len(word) > 1:
                    number_in_words = number_in_words.replace(item, item + ' و')
                else:
                    number_in_words = number_in_words.replace(item, item )
        return (number_in_words + ' و '+ fraction_in_words )
    else:
        number_in_words = re.sub("ريال.*", "",num2words(int(number_fraction[0]),lang='ar', to='currency'))
        lst = ["مائة", "مئتان", "ثلاثمائة", "أربعمائة", "خمسمائة", "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة"]
        for item in lst:
            if item in number_in_words:
                word = number_in_words.replace(item, item + ' ').split()
                if len(word) > 1:
                    number_in_words = number_in_words.replace(item, item + ' و ')
                else:
                    number_in_words = number_in_words.replace(item, item )
        number_in_words = number_in_words.replace('و  و', 'و')
        return re.sub(r" و $", "", number_in_words).replace('و  و', 'و')

def score_in_words(digit, max_mark=100):
    excellent_threshold = 0.9
    very_good_threshold = 0.8
    good_threshold = 0.7
    average_threshold = 0.6
    pass_threshold = 0.5
    
    if digit >= excellent_threshold * max_mark:
        return 'ممتاز'
    elif digit >= very_good_threshold * max_mark:
        return 'جيد جدا'
    elif digit >= good_threshold * max_mark:
        return 'جيد'
    elif digit >= average_threshold * max_mark:
        return 'متوسط'
    elif digit >= pass_threshold * max_mark:
        return 'مقبول'
    else:
        return 'مقصر'

def add_averages_to_group_list(grouped_list , skip_art_sport=True):
    
    for group in grouped_list:
        for item in group:
            term_1_avarage ,term_2_avarage , year_avarage = [0]*3        
            if 'سادس' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]/3,1)
                        term_2_avarage +=round(value[1]/3,1)
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 900)* 100,1) , round((term_2_avarage / 900)* 100,1) , round((year_avarage / 900)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'سابع' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]/3,1)
                        term_1_avarage +=round(value[1]/3,1)
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 1100)* 100,1) , round((term_2_avarage / 1100)* 100,1) , round((year_avarage / 1100)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'ثامن' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]*2/3,1)
                        term_1_avarage +=round(value[1]*2/3,1)
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 1800)* 100,1) , round((term_2_avarage / 1800)* 100,1) , round((year_avarage / 1800)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'تاسع' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]*2/3,1)
                        term_1_avarage +=round(value[1]*2/3,1)
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 2000)* 100,1) , round((term_2_avarage / 2000)* 100,1) , round((year_avarage / 2000)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'عاشر' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]*2/3,1)
                        term_1_avarage +=round(value[1]*2/3,1)
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 2000)* 100,1) , round((term_2_avarage / 2000)* 100,1) , round((year_avarage / 2000)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            else:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]*2/3,1)
                        term_1_avarage +=round(value[1]*2/3,1)
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 800)* 100,1) , round((term_2_avarage / 800)* 100,1) , round((year_avarage / 800)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

def add_subject_sum_dictionary (grouped_list):
    subject_sums = {}
    for group in grouped_list:
        for items in group:
            science_sum = 0
            social_sum = 0
            for i in items['subjects_assessments_info'][0]:
                if "علوم الأرض" in i['subject_name'] or 'الكيمياء' in i['subject_name'] or 'الحياتية' in i['subject_name'] or 'الفيزياء' in i['subject_name'] or 'العلوم' in i['subject_name']:
                    # compute sum for science subjects
                    science_sum +=  sum(int(i['term1'][key]) for key in i['term1'] if re.compile(r'^assessment\d+$').match(key) and '_max_mark' not in key and i['term1'][key])
                elif 'التربية الوطنية و المدنية' in i['subject_name'] or 'الجغرافيا' in i['subject_name'] or 'التاريخ' in i['subject_name']:
                    # compute sum for social subjects
                    social_sum +=  sum(int(i['term1'][key]) for key in i['term1'] if re.compile(r'^assessment\d+$').match(key) and '_max_mark' not in key and i['term1'][key])
                else:
                    # compute sum for other subjects
                    subject_sum = sum(int(i['term1'][key]) for key in i['term1'] if re.compile(r'^assessment\d+$').match(key) and '_max_mark' not in key and i['term1'][key])
                    # update dictionary with other subject sum
                    subject_sums[i['subject_name']] = subject_sum
            if science_sum != 0:
                # update dictionary with science subject sum
                subject_sums['العلوم'] = science_sum if science_sum != 0 else science_sum
            if social_sum != 0 :
                # update dictionary with social subject sum
                subject_sums['التربية الاجتماعية و الوطنية'] = social_sum 
            print (items['student__full_name'],items['student_class_name_letter'],subject_sums)
            items['subject_sums'] = subject_sums
            subject_sums={}

def playsound():
    # Execute the shell command to play a sine wave sound with frequency 440Hz for 2 seconds
    subprocess.run(['play', '-n', 'synth', '2', 'sin', '440'])
    
def group_students(dic_list4 , i = None):
    # sort the list based on the 'class_name' key
    sorted_list = sorted(dic_list4, key=lambda x: x['student_class_name_letter'])

    # group the sorted list by the 'class_name' key
    grouped_list = []
    for key, group in itertools.groupby(sorted_list, key=lambda x: x['student_class_name_letter']):
        group_list = list(group)
        if all(x.get('student_class_name_letter') for x in group_list):
            grouped_list.append(group_list)
    if i :
        for i in grouped_list:
            print(len(i),i[0]['student_class_name_letter'])
        return 0
    else : 
        return grouped_list
    
def get_students_info_subjectsMarks(username,password):
    '''
    دالة لاستخراج معلومات و علامات الطلاب لاستخدامها لاحقا في انشاء الجداول و العلامات
    '''
    auth=get_auth(username,password)
    dic_list=[]
    target_student_marks=[]
    school_name = inst_name(auth=auth)['data'][0]['Institutions']['name']
    edu_directory = inst_area(auth=auth)['data'][0]['Areas']['name']
    curr_year = get_curr_period(auth)['data'][0]['id']
    for i in get_school_students_ids(auth=auth)['data']:
        dic_list.append({'student_id':i['student_id'],'student__full_name':i['user']['name'],'student_nat':i['user']['nationality_id'],'student_birth_place':i['user']['birthplace_area_id'] if i['user']['birthplace_area_id'] is not None and i['user']['birthplace_area_id'] != 'None' else '' ,'student_birth_date' : i['user']['date_of_birth'] ,'student_nat_id':i['user']['identity_number'],'student_grade_id':i['education_grade_id'], 'student_grade_name' : i['education_grade_id'] ,'student_class_name_letter':'','student_edu_place' : edu_directory ,'student_directory':edu_directory,'student_school_name':school_name,'subjects_assessments_info':[] })
            
    sub_dic = {'subject_name':'','subject_number':'','term1':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''} ,'term2':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''}}
    subjects_assessments_info=[]
    # target_student_subjects = list(set(d['education_subject_id'] for d in target_student_marks))

    for i in range(0, len(dic_list), 8):
        start = i
        end = i+7 if i+7 < len(dic_list) else i+(len(dic_list)-i)-1
        student_ids = [i for i in [i['student_id'] for i in dic_list[start:end]]]
        joined_string = ','.join(str(i) for i in [f'student_id:{i}' for i in student_ids])
        marks = make_request(auth=auth,url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&academic_period_id={curr_year}&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects&_limit=0&_orWhere='+joined_string)['data']
        for student_id in student_ids:
            sub_dic = {'subject_name':'','subject_number':'','term1':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''} ,'term2':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''}}
            for mark in marks:
                if student_id in mark.values():
                    target_student_marks.append(mark)
            target_student_subjects = list(set(d['education_subject_id'] for d in target_student_marks))
            for subject in target_student_subjects:
                dictionaries = [assessments for assessments in target_student_marks if subject == assessments['education_subject_id']]
                sub_dic['subject_name'] = dictionaries[0]['education_subject']['name']
                sub_dic['subject_number']= dictionaries[0]['education_subject_id']
                sub_dic['term1']['assessment1'] = [assessments['marks'] for assessments in dictionaries if 'S1A1' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S1A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['assessment2'] = [assessments['marks'] for assessments in dictionaries if 'S1A2' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S1A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['assessment3'] = [assessments['marks'] for assessments in dictionaries if 'S1A3' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S1A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['assessment4'] = [assessments['marks'] for assessments in dictionaries if 'S1A4' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S1A4' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment1'] = [assessments['marks'] for assessments in dictionaries if 'S2A1' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S2A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment2'] = [assessments['marks'] for assessments in dictionaries if 'S2A2' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S2A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment3'] = [assessments['marks'] for assessments in dictionaries if 'S2A3' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S2A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment4'] = [assessments['marks'] for assessments in dictionaries if 'S2A4' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if 'S2A3' in assessments['assessment_period']['code']] else ''
                
                sub_dic['term1']['max_mark_assessment1'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A1' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['max_mark_assessment2'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A2' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['max_mark_assessment3'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A3' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['max_mark_assessment4'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A4' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S1A4' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment1'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A1' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment2'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A2' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment3'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A3' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment4'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A4' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if 'S2A3' in assessments['assessment_period']['code']] else ''
                subjects_assessments_info.append(sub_dic)   
                sub_dic = {'subject_name':'','subject_number':'','term1':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''} ,'term2':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''}}
                # [dic for dic in dic_list if dic['student_id']==3439303][0]['subjects_assessments_info']
            target_index = next((i for i, dic in enumerate(dic_list) if dic['student_id'] == student_id != 0 ), None)
            if target_index is not None and len(target_student_subjects) != 0:
                dic_list[target_index]['subjects_assessments_info'].append(subjects_assessments_info)
                dic_list[target_index]['student_class_name_letter'] = dictionaries[0]['institution_classes_id']             
                # print(dic_list[target_index])
                subjects_assessments_info=[]
                target_student_marks = []

    class_name_letter = list(set([i['student_class_name_letter'] for i in dic_list if i['student_class_name_letter'] != '' ]))
    joined_string = ','.join(str(i) for i in [f'institution_class_id:{i}' for i in class_name_letter])
    classes_data = make_request(auth=auth,url='https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere='+joined_string)['data']            
    class_list = []
    for i in classes_data:
        class_list.append({'class_id': i['institution_class_id'] , 'class_name': i['institution_class']['name'] })
        class_dict = {i['class_id']: i['class_name'] for i in class_list if i['class_id'] != ''}
    for i in dic_list:
        class_id = i['student_class_name_letter']
        if class_id != '':
            i['student_class_name_letter'] = class_dict.get(class_id, class_id)
    grade_id = list(set([i['student_grade_id'] for i in dic_list if i['student_grade_id'] != '' ]))
    grade_data = get_grade_info(auth)
    grade_list = []
    for i in grade_data:
        grade_list.append({'grade_id': i['education_grade_id'] , 'grade_name': re.sub('.*للصف','الصف', i['name']) })
        grade_dict = {i['grade_id']: i['grade_name'] for i in grade_list if i['grade_name'] != ''}
    for i in dic_list:
        grade_id = i['student_grade_id']
        if grade_id != '':
            i['student_grade_name'] = grade_dict.get(grade_id, grade_id)
            
    nat_id = list(set([i['student_grade_id'] for i in dic_list if i['student_grade_id'] != '' ]))
    birth_place_id = list(set([i['student_grade_id'] for i in dic_list if i['student_grade_id'] != '' ]))
    birth_place_data = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Area-AreaAdministratives?_limit=0&_contain=AreaAdministrativeLevels')['data']
    nationality_data = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/User-NationalityNames')['data']
    nationality_data_list = []
    birth_place_list = []
    for i in nationality_data:
        nationality_data_list.append({'nat_id': i['id'] , 'nat_name': i['name'] })
        nattionality_dict = {i['nat_id']: i['nat_name'] for i in nationality_data_list if i['nat_name'] != ''}
    for i in dic_list:
        nat_id = i['student_nat']
        if nat_id != '':
            i['student_nat'] = nattionality_dict.get(nat_id, nat_id)

    for i in birth_place_data:
        birth_place_list.append({'birth_place_id': i['id'] , 'birth_place_name': i['name'] })
        birth_place_dict = {i['birth_place_id']: i['birth_place_name'] for i in birth_place_list if i['birth_place_name'] != ''}
    for i in dic_list:
        birth_place_id = i['student_birth_place']
        if birth_place_id != '':
            i['student_birth_place'] = birth_place_dict.get(birth_place_id, birth_place_id)

    return dic_list

def get_grade_info(auth):
    my_list = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0')['data']
    return my_list

def get_school_students_ids(auth):
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    curr_year = get_curr_period(auth)['data'][0]['id']
    return make_request(auth=auth,url=f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.Students?_limit=0&_finder=Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_number,Users.external_reference,Users.status&institution_id={inst_id}&academic_period_id={curr_year}&_contain=Users')

def fill_official_marks_a3_two_face_doc2_offline_version(username, password ,students_data_lists, ods_file ):
    '''
    doc is the copy that you want to send 
    '''
    context = {'46': 'A6:A30', '4': 'A39:A63', '3': 'L6:L30', '45': 'L39:L63', '44': 'A71:A95', '6': 'A103:A127', '5': 'L71:L95', '43': 'L103:L127', '42': 'A135:A159', '8': 'A167:A191', '7': 'L135:L159', '41': 'L167:L191', '40': 'A199:A223', '10': 'A231:A255', '9': 'L199:L223', '39': 'L231:L255', '38': 'A263:A287', '12': 'A295:A319', '11': 'L263:L287', '37': 'L295:L319', '36': 'A327:A351', '14': 'A359:A383', '13': 'L327:L351', '35': 'L359:L383', '34': 'A391:A415', '16': 'A423:A447', '15': 'L391:L415', '33': 'L423:L447', '32': 'A455:A479', '18': 'A487:A511', '17': 'L455:L479', '31': 'L487:L511', '30': 'A519:A543', '20': 'A551:A575', '19': 'L519:L543', '29': 'L551:L575', '28': 'A583:A607', '22': 'A615:A639', '21': 'L583:L607', '27': 'L615:L639', '26': 'A647:A671', '24': 'A679:A703', '23': 'L647:L671', '25': 'L679:L703'}
    
    page = 4
    name_counter = 1
    name_counter = 1
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    user_id = user_info(auth , username)['data'][0]['id']
    
    user = user_info(auth , username)
    school_name = inst_name(auth)['data'][0]['Institutions']['name']
    baldah = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={inst_id}&_contain=InstitutionLands.CustomFieldValues')['data'][0]['address'].split('-')[0]
    grades= make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Education.EducationGrades?_limit=0')
    modeeriah = inst_area(auth)['data'][0]['Areas']['name']
    school_year = get_curr_period(auth)['data']
    hejri1 = str(hijri_converter.convert.Gregorian(school_year[0]['start_year'], 1, 1).to_hijri().year)
    hejri2 =  str(hijri_converter.convert.Gregorian(school_year[0]['end_year'], 1, 1).to_hijri().year)
    melady1 = str(school_year[0]['start_year'])
    melady2 = str(school_year[0]['end_year'])
    teacher = user['data'][0]['name'].split(' ')[0]+' '+user['data'][0]['name'].split(' ')[-1]
    
    classes=[]
    mawad=[]
    modified_classes=[]
    
    # Open the ODS file and load the sheet you want to fill
    doc = ezodf.opendoc(ods_file) 
       
    sheet_name = 'sheet'
    sheet = doc.sheets[sheet_name]


    for students_data_list in students_data_lists:
        
#         ['الصف السابع', 'أ', 'اللغة الانجليزية', '786118']
        
        class_data = students_data_list['class_name'].split('-')
        mawad.append(class_data[2])
        classes.append('-'.join([class_data[0],class_data[1]]))
        class_name = class_data[0].replace('الصف ' , '')
        class_char = class_data[1]
        sub_name = class_data[2]   

        
        sheet[f"D{int(context[str(page)].split(':')[0][1:])-5 }"].set_value(f' الصف: {class_name}')
        sheet[f"I{int(context[str(page)].split(':')[0][1:])-5 }"].set_value(f'الشعبة (   {class_char}    )')    
        sheet[f"O{int(context[str(page+1)].split(':')[0][1:])-5}"].set_value(sub_name)
              
        for counter,student_info in enumerate(students_data_list['sdtudent_data'], start=1):
            if counter >= 26:
                page += 2
                counter = 1
                
                sheet[f"D{int(context[str(page)].split(':')[0][1:])-5}"].set_value(f' الصف: {class_name}')
                sheet[f"I{int(context[str(page)].split(':')[0][1:])-5}"].set_value(f'الشعبة (   {class_char}    )')  
                sheet[f"O{int(context[str(page+1)].split(':')[0][1:])-5}"].set_value(sub_name)
                #    المادة الدراسية     
                
                # {'id': 3824166, 'name': 'نورالدين محمود راضي الدغيمات', 'term1': {'assessment1': 9, 'assessment2': 10, 'assessment3': 11, 'assessment4': 20}}
                
                for student_info in students_data_list['sdtudent_data'][25:] :
                    row_idx = counter + int(context[str(page)].split(':')[0][1:]) - 1  # compute the row index based on the counter
                    sheet[f"A{row_idx}"].set_value(name_counter)
                    sheet[f"B{row_idx}"].set_value(student_info['name'])
                    if 'term1' in student_info and 'assessment1' in student_info['term1'] and 'assessment2' in student_info['term1'] and 'assessment3' in student_info['term1'] and 'assessment4' in student_info['term1']:
                        sheet[f"D{row_idx}"].set_value(student_info['term1']['assessment1']) 
                        sheet[f"E{row_idx}"].set_value(student_info['term1']['assessment2']) 
                        sheet[f"F{row_idx}"].set_value(student_info['term1']['assessment3'])
                        sheet[f"G{row_idx}"].set_value(student_info['term1']['assessment4'])
                    if 'term2' in student_info:
                        row_idx2 = counter + int(context[str(page+1)].split(':')[0][1:]) - 1  # compute the row index based on the counter 
                        sheet[f"L{row_idx2}"].set_value(student_info['term2']['assessment1']) 
                        sheet[f"M{row_idx2}"].set_value(student_info['term2']['assessment2']) 
                        sheet[f"N{row_idx2}"].set_value(student_info['term2']['assessment3'])
                        sheet[f"O{row_idx2}"].set_value(student_info['term2']['assessment4'])                       
                    counter += 1
                    name_counter += 1              
                break                    
            row_idx = counter + int(context[str(page)].split(':')[0][1:]) - 1  # compute the row index based on the counter
            sheet[f"A{row_idx}"].set_value(name_counter)
            sheet[f"B{row_idx}"].set_value(student_info['name']) 
            if 'term1' in student_info and 'assessment1' in student_info['term1'] and 'assessment2' in student_info['term1'] and 'assessment3' in student_info['term1'] and 'assessment4' in student_info['term1']:
                sheet[f"D{row_idx}"].set_value(student_info['term1']['assessment1']) 
                sheet[f"E{row_idx}"].set_value(student_info['term1']['assessment2']) 
                sheet[f"F{row_idx}"].set_value(student_info['term1']['assessment3'])
                sheet[f"G{row_idx}"].set_value(student_info['term1']['assessment4'])
            if 'term2' in student_info:
                row_idx2 = counter + int(context[str(page+1)].split(':')[0][1:]) - 1  # compute the row index based on the counter 
                sheet[f"L{row_idx2}"].set_value(student_info['term2']['assessment1']) 
                sheet[f"M{row_idx2}"].set_value(student_info['term2']['assessment2']) 
                sheet[f"N{row_idx2}"].set_value(student_info['term2']['assessment3'])
                sheet[f"O{row_idx2}"].set_value(student_info['term2']['assessment4'])                
            name_counter += 1 
        name_counter = 1
        page += 2

    
    for i in classes: 
        modified_classes.append(mawad_representations(i))
        
    modified_classes = ' ، '.join(modified_classes)
    mawad = sorted(set(mawad))
    mawad = ' ، '.join(mawad)

    custom_shapes = {
        'modeeriah': f'لواء {modeeriah}',
        'hejri1': hejri1,
        'hejri2': hejri2,
        'melady1': melady1,
        'melady2': melady2,
        'baldah': baldah,
        'school': school_name,
        'classes': modified_classes,
        'mawad': mawad,
        'teacher' : teacher,
        'modeeriah_20_2': f'لواء {modeeriah}',
        'hejri_20_1': hejri1,
        'hejri_20_2': hejri2,
        'melady_20_1': melady1,
        'melady_20_2': melady2,
        'baldah_20_2': baldah,
        'school_20_2': school_name,
        'classes_20_2': modified_classes,
        'mawad_20_2': mawad,
        'teacher_20_2': teacher ,
        'modeeriah_20_1': f'لواء {modeeriah}',
        'hejri1': hejri1,
        'hejri2': hejri2,
        'melady1': melady1,
        'melady2': melady2,
        'baldah_20_1': baldah,
        'school_20_1': school_name,
        'classes_20_1': modified_classes,
        'mawad_20_1': mawad,
        'teacher_20_1': teacher
    }
    # FIXME: make the customshapes crop _20_ to the rest of the key in the custom_shapes
    # Iterate through the cells of the sheet and fill in the values you want
    doc.save()
            
    return custom_shapes 

def Read_E_Side_Note_Marks(file_path=None , file_content=None):
    if file_content is None:
        # Load the workbook
        wb = load_workbook(file_path)
    else:
        wb = load_workbook(filename=file_content)
        
    sheets = wb.sheetnames
    sheet = wb[wb.sheetnames[0]]

    read_file_output_lists = []

    for sheet in sheets :
        rows = []
        data = []
        # Loop over the rows in each sheet
        for row in wb[sheet].iter_rows(min_row=3, values_only=True):
            row = [cell if cell is not None else '' for cell in row]
            # Append the row data to the list
            rows.append(list(row))    

        for row in rows:
            dic = {
                'id': row[1], 
                'name':  row[2],
                'term1': {'assessment1':  row[3], 'assessment2':row[4], 'assessment3': row[5], 'assessment4': row[6]},
                'term2': {'assessment1': row[8], 'assessment2': row[9], 'assessment3': row[10], 'assessment4': row[11]}
                    }
            data.append(dic)
        temp_dic = {'class_name':sheet ,"sdtudent_data": data}
        read_file_output_lists.append(temp_dic)

    return read_file_output_lists

def enter_marks_arbitrary_controlled_version(username , password , required_data_list ,range1 ,range2):
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    
    for item in required_data_list : 
        for Student_id in item['students_ids']:
            enter_mark(auth 
                ,marks= random.randint(range1, range2)
                ,assessment_grading_option_id= 8
                ,assessment_id= item['assessment_id']
                ,education_subject_id= item['education_subject_id']
                ,education_grade_id= item['education_grade_id']
                ,institution_id= inst_id
                ,academic_period_id= period_id
                ,institution_classes_id= item['institution_classes_id']
                ,student_status_id= 1
                ,student_id= Student_id
                ,assessment_period_id= item['assessment_id'])
                        
def assessments_commands_text(lst):
    S1 = [i for i in lst if i.get('SEname') !='الفصل الثاني']
    S2 = [i for i in lst if i.get('SEname') =='الفصل الثاني']    
    text = ""

    if S1:
        text = 'الفصل الاول\n'
        for item in S1:
            text += '/' + item['code'] + ' الصف ال' + num2words(int(re.match('G\d{1,}', item['code']).group()[1:]), lang='ar', to='ordinal') + ' ' + item['AssesName'] + ' ' + ' علامة النجاح ' + str(item['pass_mark']) + ' و العلامة القصوى ' + str(item['max_mark']) + '\n'

    if S2:
        text += 'الفصل الثاني\n'
        for item in S2:
            text += '/' + item['code'] + ' الصف ال' + num2words(int(re.match('G\d{1,}', item['code']).group()[1:]), lang='ar', to='ordinal') + ' ' + item['AssesName'] + ' ' + ' علامة النجاح ' + str(item['pass_mark']) + ' و العلامة القصوى ' + str(item['max_mark']) + '\n'

    if not S1 and not S2:
        # change this to send message for user that there is no assessement to fill now
        print("Both S1 and S2 lists are empty.")
    else:
        return text
    
def get_editable_assessments( auth , username):
    required_data_list = get_required_data_to_enter_marks(auth=auth ,username=username)
    ass_data = [[y['assessment_id'],y['education_subject_id']] for y in required_data_list ]
    ass_data = [item for sublist in [get_all_assessments_periods_data2(auth, i[0],i[1]) for i in ass_data] for item in sublist if item.get('editable')==True]
    # unique_lst = [dict(t) for t in {tuple(sorted(d.items())) for d in lst}]
    unique_dict_list = [dict(t) for t in {tuple(sorted(d.items())) for d in ass_data}]
    sorted_dict = sorted(unique_dict_list , key=lambda x: x['code'])
    return sorted_dict

def assessments_periods_min_max_mark(auth , assessment_id , education_subject_id ):
    '''
         استعلام عن القيمة القصوى و الدنيا لكل التقويمات  
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود بمعلومات عن تقيمات الصفوف في السنة الدراسية  
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentItemsGradingTypes.json?_contain=EducationSubjects,AssessmentGradingTypes.GradingOptions&assessment_id={assessment_id}&education_subject_id={education_subject_id}&_limit=0"
    return make_request(url,auth)

def get_all_assessments_periods_data2(auth , assessment_id ,education_subject_id):
    '''
         استعلام عن تعريفات التقويمات في السنة الدراسية و امكانية تحرير التقويم و  العلامة القصوى و الدنيا
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود تعريفات التقويمات في السنة الدراسية و امكانية تحرير التقويم و  العلامة القصوى و الدنيا  
    '''
    terms = get_AcademicTerms(auth=auth , assessment_id=assessment_id)['data']
    season_assessments = []
    dic =  {'SEname': '', 'AssesName': '' ,'AssesId': '' , 'pass_mark': '' , 'max_mark' : '' , 'editable' : '' , 'code':'' , 'gradeId':''}
    min_max=[]
    for i in assessments_periods_min_max_mark(auth , assessment_id, education_subject_id)['data']:
        min_max.append({'id': i['assessment_period_id'] , 'pass_mark':i['assessment_grading_type']['pass_mark'] , 'max_mark' : i['assessment_grading_type']['max'] } )                    
    for term in terms:
        for asses in get_assessments_periods(auth, term['name'], assessment_id=assessment_id)['data']:
            dic = {'SEname': asses["academic_term"], 'AssesName': asses["name"], 'AssesId': asses["id"] , 'pass_mark': [dictionary['pass_mark'] for dictionary in min_max if dictionary.get('id') == asses["id"]][0] , 'max_mark' : [dictionary['max_mark'] for dictionary in min_max if dictionary.get('id') == asses["id"]][0] , 'editable':asses['editable'], 'code': asses['code'], 'gradeId':asses['assessment_id']}
            season_assessments.append(dic)
    return season_assessments

def enter_marks_arbitrary(username , password , assessment_period_id ,range1 ,range2):
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    
    required_data_list = get_required_data_to_enter_marks(auth , username)
    for item in required_data_list : 
        for Student_id in item['students_ids']:
            enter_mark(auth 
                ,marks= random.randint(range1, range2)
                ,assessment_grading_option_id= 8
                ,assessment_id= item['assessment_id']
                ,education_subject_id= item['education_subject_id']
                ,education_grade_id= item['education_grade_id']
                ,institution_id= inst_id
                ,academic_period_id= period_id
                ,institution_classes_id= item['institution_classes_id']
                ,student_status_id= 1
                ,student_id= Student_id
                ,assessment_period_id= assessment_period_id)

def get_class_students_ids(auth,academic_period_id,institution_subject_id,institution_class_id,institution_id):
    '''
    استدعاء معلومات عن الطلاب في الصف
    عوامل الدالة هي الرابط و التوكن و تعريفي الفترة الاكاديمية و تعريفي مادة المؤسسة و تعريفي صف المؤسسة و تعريفي المؤسسة
    تعود بمعلومات تفصيلية عن كل طالب في الصف بما في ذلك اسمه الرباعي و التعريفي و مكان سكنه
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id&_limit=0&academic_period_id={academic_period_id}&institution_subject_id={institution_subject_id}&institution_class_id={institution_class_id}&institution_id={institution_id}&_contain=Users"
    student_ids = [student['student_id'] for student in make_request(url,auth)['data']]
    return student_ids

def get_required_data_to_enter_marks(auth ,username):
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    user_id = user_info(auth , username)['data'][0]['id']
    years = get_curr_period(auth)
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id)['data']]
    required_data_to_enter_marks = []
    
    for class_id in classes_id_1 : 
        class_info = get_teacher_classes2( auth , class_id)['data']
        dic = {'assessment_id':'','education_subject_id':'' ,'education_grade_id':'','institution_classes_id':'','students_ids':[] }
        dic['assessment_id'] = get_assessment_id_from_grade_id(auth , class_info[0]['institution_subject']['education_grade_id'])
        dic['education_subject_id'] = class_info[0]['institution_subject']['education_subject_id']
        dic['education_grade_id'] = class_info[0]['institution_subject']['education_grade_id']
        dic['institution_classes_id'] = class_info[0]['institution_class_id']
        dic['students_ids'] = get_class_students_ids(auth,period_id,class_info[0]['institution_subject_id'],class_info[0]['institution_class_id'],inst_id)

        required_data_to_enter_marks.append(dic)
    
    return required_data_to_enter_marks

def get_grade_info(auth):
    
    my_list = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0')['data']
    return my_list
   
def get_grade_name_from_grade_id(auth , grade_id):
    
    my_list = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0')['data']

    return [d['name'] for d in my_list if d.get('education_grade_id') == grade_id][0].replace('الفترات التقويمية ل','ا')

def get_assessment_id_from_grade_id(auth , grade_id):
    
    my_list = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0')['data']

    return [d['id'] for d in my_list if d.get('education_grade_id') == grade_id][0]

def create_e_side_marks_doc(username , password ,template='./templet_files/e_side_marks.xlsx' ,outdir='./send_folder' ):
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    userInfo = user_info(auth , username)['data'][0]
    user_id , user_name = userInfo['id'] , userInfo['first_name']+' '+ userInfo['last_name']+'-' + str(username)
    years = get_curr_period(auth)
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id)['data']]
    classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i])['data'] for i in range(len(classes_id_1))]
    classes_id_3 = []  

    # load the existing workbook
    existing_wb = load_workbook(template)

    # Select the worksheet
    existing_ws = existing_wb.active

    for class_info in classes_id_2:
        classes_id_3.append([{"institution_class_id": class_info[0]['institution_class_id'] ,"sub_name": class_info[0]['institution_subject']['name'],"class_name": class_info[0]['institution_class']['name']}])

    for v in range(len(classes_id_1)):
        # id
        print (classes_id_3[v][0]['institution_class_id'])
        # subject name 
        print (classes_id_3[v][0]['sub_name'])
        # class name
        print (classes_id_3[v][0]['class_name'])

        # copy the worksheet
        new_ws = existing_wb.copy_worksheet(existing_ws)

        # rename the new worksheet
        new_ws.title = classes_id_3[v][0]['class_name']+'-'+classes_id_3[v][0]['sub_name']+'-'+str(classes_id_3[v][0]['institution_class_id'])
        new_ws.sheet_view.rightToLeft = True    
        existing_ws.sheet_view.rightToLeft = True   


        students = get_class_students(auth
                                    ,period_id
                                    ,classes_id_1[v]
                                    ,classes_id_3[v][0]['institution_class_id']
                                    ,inst_id)
        students_names = sorted([i['user']['name'] for i in students['data']])
        print(students_names)
        students_id_and_names = []
        for IdAndName in students['data']:
            students_id_and_names.append({'student_name': IdAndName['user']['name'] , 'student_id':IdAndName['student_id']})

        assessments_json = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id=4&institution_classes_id='+ str(classes_id_3[v][0]['institution_class_id'])+ f'&institution_id={inst_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects')

        marks_and_name = []
        dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} }
        for i in students_id_and_names:   
            for v in assessments_json['data']:
                if v['student_id'] == i['student_id'] :  
                    dic['id'] = i['student_id'] 
                    dic['name'] = i['student_name'] 
                    if v['assessment_period']['name'] == 'التقويم الأول' and v['assessment_period']['academic_term'] == 'الفصل الأول':
                        dic['term1']['assessment1'] = v["marks"]     
                    elif v['assessment_period']['name'] == 'التقويم الثاني' and v['assessment_period']['academic_term'] == 'الفصل الأول':
                        dic['term1']['assessment2']  = v["marks"]             
                    elif v['assessment_period']['name'] == 'التقويم الثالث' and v['assessment_period']['academic_term'] == 'الفصل الأول':
                        dic['term1']['assessment3']  = v["marks"]           
                    elif v['assessment_period']['name'] == 'التقويم الرابع' and v['assessment_period']['academic_term'] == 'الفصل الأول':
                        dic['term1']['assessment4']  = v["marks"]
                    elif v['assessment_period']['name'] == 'التقويم الأول' and v['assessment_period']['academic_term'] == 'الفصل الثاني':
                        dic['term2']['assessment1']  = v["marks"]     
                    elif v['assessment_period']['name'] == 'التقويم الثاني' and v['assessment_period']['academic_term'] == 'الفصل الثاني':
                        dic['term2']['assessment2']  = v["marks"]             
                    elif v['assessment_period']['name'] == 'التقويم الثالث' and v['assessment_period']['academic_term'] == 'الفصل الثاني':
                        dic['term2']['assessment3']  = v["marks"]           
                    elif v['assessment_period']['name'] == 'التقويم الرابع' and v['assessment_period']['academic_term'] == 'الفصل الثاني':
                        dic['term2']['assessment4']  = v["marks"]
            marks_and_name.append(dic)
            dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} }
        # Set the font for the data rows
        data_font = Font(name='Arial', size=16, bold=False)

        marks_and_name = [d for d in marks_and_name if d['name'] != '']
        marks_and_name = sorted(marks_and_name, key=lambda x: x['name'])

        # Write data to the worksheet and calculate the sum of some columns in each row
        for row_number, dataFrame in enumerate(marks_and_name, start=3):
            new_ws.cell(row=row_number, column=1).value = row_number-2
            new_ws.cell(row=row_number, column=2).value = dataFrame['id']
            new_ws.cell(row=row_number, column=3).value = dataFrame['name']
            new_ws.cell(row=row_number, column=4).value = dataFrame['term1']['assessment1']
            new_ws.cell(row=row_number, column=5).value = dataFrame['term1']['assessment2']
            new_ws.cell(row=row_number, column=6).value = dataFrame['term1']['assessment3']
            new_ws.cell(row=row_number, column=7).value = dataFrame['term1']['assessment4']
            new_ws.cell(row=row_number, column=8).value = f'=SUM(D{row_number}:G{row_number})'
            new_ws.cell(row=row_number, column=9).value = dataFrame['term2']['assessment1']
            new_ws.cell(row=row_number, column=10).value = dataFrame['term2']['assessment2']
            new_ws.cell(row=row_number, column=11).value = dataFrame['term2']['assessment3']
            new_ws.cell(row=row_number, column=12).value = dataFrame['term2']['assessment4']
            new_ws.cell(row=row_number, column=13).value = f'=SUM(I{row_number}:L{row_number})'

            # Set the font for the data rows
            for cell in new_ws[row_number]:
                cell.font = data_font
    existing_wb.remove(existing_wb['Sheet1'])

    # save the modified workbook
    existing_wb.save(f'{outdir}/{user_name}.xlsx')

def split_A3_pages(input_file, outdir):
    # Open the A3 PDF file in read-binary mode
    with open(input_file, 'rb') as pdf_file:
        # Create a PDF reader object
        pdf_reader = PyPDF4.PdfFileReader(pdf_file)

        # Create a new PDF writer object
        pdf_writer = PyPDF4.PdfFileWriter()

        # Iterate through each page of the PDF file
        for page_num in range(pdf_reader.getNumPages()):
            # Get the current page of the PDF file
            page = pdf_reader.getPage(page_num)

            # Create a new page with A4 size by splitting the A3 page into two A4 pages
            x1, y1, x2, y2 = page.mediaBox.lowerLeft + page.mediaBox.upperRight
            x_mid = (x1 + x2) / 2
            a4_size = (100, -25)
            page1 = PyPDF4.pdf.PageObject.createBlankPage(None, a4_size[0], a4_size[1])
            page1.mergeTranslatedPage(page, -x1, -y1)
            page1.mediaBox.upperRight = (x2 - x_mid + 50, y2 - a4_size[1])
            page2 = PyPDF4.pdf.PageObject.createBlankPage(None, a4_size[0], a4_size[1])
            page2.mergeTranslatedPage(page, -x_mid, -y1)
            page2.mediaBox.upperRight = (x2 - x_mid , y2 - a4_size[1])

            # Add the two A4 pages to the PDF writer object
            pdf_writer.addPage(page1)
            pdf_writer.addPage(page2)

        # Save the new A4 pages to a new PDF file
        with open(f'{outdir}/output.pdf', 'wb') as output_file:
            pdf_writer.write(output_file)

def reorder_official_marks_to_A4(input_file, out_file):
    # Load the PDF document
    with open(input_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfFileReader(file)

        # List of page locations in the new order
        new_order_list = ["1=1","52=2","51=3","2=4","3=5","50=6","49=7","3=8","4=9","46=10","45=11","4=12","5=13","44=14","43=15","6=16","7=17","42=18","41=19","8=20","9=21","40=22","39=23","10=24","11=25","38=26","37=27","12=28","13=29","36=30","35=31","14=32","15=33","34=34","33=35","16=36","17=37","32=38","31=39","18=40","19=41","30=42","29=43","20=44","21=45","28=46","27=47","22=48","23=49","26=50","25=51","24=52"]

        # Create a dictionary from the new order list
        new_order_dict = {}
        for item in new_order_list:
            location, page_number = item.split('=')
            new_order_dict[int(page_number)] = int(location)

        # Sort the dictionary by values
        sorted_dict = dict(sorted(new_order_dict.items(), key=lambda x: x[1]))

    #     print(len(sorted_dict))
        # Create a new PDF document object
        pdf_writer = PyPDF2.PdfFileWriter()

        # Iterate over the sorted dictionary's keys and add the corresponding page to the new PDF document
        for page_number in sorted_dict.keys():
            pdf_writer.addPage(pdf_reader.getPage(page_number - 1))

    #     Save the new PDF document with the reordered pages
        with open(out_file, 'wb') as file:
            pdf_writer.write(file)
            
def delete_files_except(filenames, dir_path):
    """
    Deletes every ODS or PDF file in the specified directory except for the files with the specified names.
    """
    for file in os.listdir(dir_path):
        if file not in filenames and (file.endswith(".ods") or file.endswith(".pdf") or file.endswith(".bak") ):
            os.remove(os.path.join(dir_path, file))

def fill_official_marks_doc_wrapper_offline(usnername , password ,lst, ods_name='send', outdir='./send_folder' ,ods_num=1):
    ods_file = f'{ods_name}{ods_num}.ods'
    copy_ods_file('./templet_files/official_marks_doc_a3_two_face.ods' , f'{outdir}/{ods_file}')
    custom_shapes = fill_official_marks_a3_two_face_doc2_offline_version(username= usnername, password= password ,students_data_lists=lst, ods_file=f'{outdir}/{ods_file}')
    
    fill_custom_shape(doc= f'{outdir}/{ods_file}' ,sheet_name= 'الغلاف الداخلي' , custom_shape_values= custom_shapes , outfile=f'{outdir}/modified.ods')
    fill_custom_shape(doc=f'{outdir}/modified.ods', sheet_name='الغلاف الازرق', custom_shape_values=custom_shapes, outfile=f'{outdir}/final_'+ods_file)
    os.system(f'soffice --headless --convert-to pdf:writer_pdf_Export --outdir {outdir} {outdir}/final_{ods_file} ')
    add_margins(f"{outdir}/final_{ods_name}{ods_num}.pdf", f"{outdir}/output_file.pdf",top_rec=30, bottom_rec=50, left_rec=68, right_rec=120)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/{custom_shapes['teacher']}.pdf",page=1 , top_rec=60, bottom_rec=80, left_rec=70, right_rec=120)
    split_A3_pages(f"{outdir}/output_file.pdf" , outdir)
    reorder_official_marks_to_A4(f"{outdir}/output.pdf" , f"{outdir}/reordered.pdf")

    add_margins(f"{outdir}/reordered.pdf", f"{outdir}/output_file.pdf",top_rec=60, bottom_rec=50, left_rec=68, right_rec=20)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/output_file1.pdf",page=1 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120)
    add_margins(f"{outdir}/output_file1.pdf", f"{outdir}/output_file2.pdf",page=50 , top_rec=100, bottom_rec=80, left_rec=70, right_rec=60)    
    add_margins(f"{outdir}/output_file2.pdf", f"{outdir}/{custom_shapes['teacher']}_A4.pdf",page=51 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120)  
    delete_files_except([f"{custom_shapes['teacher']}.pdf",f"{custom_shapes['teacher']}_A4.pdf"], outdir)
    
def fill_official_marks_doc_wrapper(usnername , password , ods_name='send', outdir='./send_folder' ,ods_num=1):
    ods_file = f'{ods_name}{ods_num}.ods'
    copy_ods_file('./templet_files/official_marks_doc_a3_two_face.ods' , f'{outdir}/{ods_file}')
    
    custom_shapes = fill_official_marks_a3_two_face_doc2(username= usnername, password= password , ods_file=f'{outdir}/{ods_file}')
    fill_custom_shape(doc= f'{outdir}/{ods_file}' ,sheet_name= 'الغلاف الداخلي' , custom_shape_values= custom_shapes , outfile=f'{outdir}/modified.ods')
    fill_custom_shape(doc=f'{outdir}/modified.ods', sheet_name='الغلاف الازرق', custom_shape_values=custom_shapes, outfile=f'{outdir}/final_'+ods_file)
    os.system(f'soffice --headless --convert-to pdf:writer_pdf_Export --outdir {outdir} {outdir}/final_{ods_file} ')
    add_margins(f"{outdir}/final_{ods_name}{ods_num}.pdf", f"{outdir}/output_file.pdf",top_rec=30, bottom_rec=50, left_rec=68, right_rec=120)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/{custom_shapes['teacher']}.pdf",page=1 , top_rec=60, bottom_rec=80, left_rec=70, right_rec=120)
    split_A3_pages(f"{outdir}/output_file.pdf" , outdir)
    reorder_official_marks_to_A4(f"{outdir}/output.pdf" , f"{outdir}/reordered.pdf")

    add_margins(f"{outdir}/reordered.pdf", f"{outdir}/output_file.pdf",top_rec=60, bottom_rec=50, left_rec=68, right_rec=20)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/output_file1.pdf",page=1 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120)
    add_margins(f"{outdir}/output_file1.pdf", f"{outdir}/output_file2.pdf",page=50 , top_rec=100, bottom_rec=80, left_rec=70, right_rec=60)    
    add_margins(f"{outdir}/output_file2.pdf", f"{outdir}/{custom_shapes['teacher']}_A4.pdf",page=51 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120)  
    delete_files_except([f"{custom_shapes['teacher']}.pdf",f"{custom_shapes['teacher']}_A4.pdf"], outdir)

def delete_file(file_path):
    """Delete a file"""
    os.remove(file_path)

def copy_ods_file(source_file_path, destination_folder):
    """Copy an ODS file to a destination folder"""
    shutil.copy(source_file_path, destination_folder)
    
def fill_official_marks_a3_two_face_doc2(username, password , ods_file ):
    '''
    doc is the copy that you want to send 
    '''
    context = {'46': 'A6:A30', '4': 'A39:A63', '3': 'L6:L30', '45': 'L39:L63', '44': 'A71:A95', '6': 'A103:A127', '5': 'L71:L95', '43': 'L103:L127', '42': 'A135:A159', '8': 'A167:A191', '7': 'L135:L159', '41': 'L167:L191', '40': 'A199:A223', '10': 'A231:A255', '9': 'L199:L223', '39': 'L231:L255', '38': 'A263:A287', '12': 'A295:A319', '11': 'L263:L287', '37': 'L295:L319', '36': 'A327:A351', '14': 'A359:A383', '13': 'L327:L351', '35': 'L359:L383', '34': 'A391:A415', '16': 'A423:A447', '15': 'L391:L415', '33': 'L423:L447', '32': 'A455:A479', '18': 'A487:A511', '17': 'L455:L479', '31': 'L487:L511', '30': 'A519:A543', '20': 'A551:A575', '19': 'L519:L543', '29': 'L551:L575', '28': 'A583:A607', '22': 'A615:A639', '21': 'L583:L607', '27': 'L615:L639', '26': 'A647:A671', '24': 'A679:A703', '23': 'L647:L671', '25': 'L679:L703'}
    
    page = 4
    name_counter = 1
    name_counter = 1
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    user_id = user_info(auth , username)['data'][0]['id']
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id)['data']]
    classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i])['data'] for i in range(len(classes_id_1))]
    classes_id_3 = []
    
    user = user_info(auth , username)
    school_name = inst_name(auth)['data'][0]['Institutions']['name']
    baldah = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={inst_id}&_contain=InstitutionLands.CustomFieldValues')['data'][0]['address'].split('-')[0]
    grades= make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Education.EducationGrades?_limit=0')
    modeeriah = inst_area(auth)['data'][0]['Areas']['name']
    school_year = get_curr_period(auth)['data']
    hejri1 = str(hijri_converter.convert.Gregorian(school_year[0]['start_year'], 1, 1).to_hijri().year)
    hejri2 =  str(hijri_converter.convert.Gregorian(school_year[0]['end_year'], 1, 1).to_hijri().year)
    melady1 = str(school_year[0]['start_year'])
    melady2 = str(school_year[0]['end_year'])
    teacher = user['data'][0]['name'].split(' ')[0]+' '+user['data'][0]['name'].split(' ')[-1]
    
    classes=[]
    mawad=[]
    modified_classes=[]
    
    # Open the ODS file and load the sheet you want to fill
    doc = ezodf.opendoc(ods_file) 
       
    sheet_name = 'sheet'
    sheet = doc.sheets[sheet_name]

    for class_info in classes_id_2:
        classes_id_3.append([{"institution_class_id": class_info[0]['institution_class_id'] ,"sub_name": class_info[0]['institution_subject']['name'],"class_name": class_info[0]['institution_class']['name'] , 'sub_id' : class_info[0]['institution_subject']['education_subject_id']}])
    institution_subjects_id = [i[0]["institution_class_id"] for i in classes_id_3]
    for v in range(len(classes_id_1)):
        # id
        print (classes_id_3[v][0]['institution_class_id'])
        # subject name 
        print (classes_id_3[v][0]['sub_name'])
        # class name
        print (classes_id_3[v][0]['class_name'])
        mawad.append(classes_id_3[v][0]['sub_name'])
        classes.append(classes_id_3[v][0]['class_name'])
        class_name = classes_id_3[v][0]['class_name'].split('-')[0].replace('الصف ' , '')
        class_char = classes_id_3[v][0]['class_name'].split('-')[1]
        sub_name = classes_id_3[v][0]['sub_name']    
        students = get_class_students(auth
                                    ,period_id
                                    ,classes_id_1[v]
                                    ,classes_id_3[v][0]['institution_class_id']
                                    ,inst_id)
        # students_and_marks
        all1 = get_students_marks(auth
                                                ,period_id
                                                ,classes_id_3[v][0]['sub_id']
                                                ,classes_id_3[v][0]['institution_class_id']
                                                ,inst_id)   
        students_names = []
        for IdAndName in students['data']:
            students_names.append({'student_name': IdAndName['user']['name'] , 'student_id':IdAndName['student_id']})
   
        marks_and_name = []
        mark_data =  {'Sid':'' ,'Sname': '','S1':{ 'ass1': '' ,'ass2': '' , 'ass3': '' , 'ass4': ''} ,'S2':{ 'ass1': '' ,'ass2': '' , 'ass3': '' , 'ass4': ''} }
        term_mapping = {
            "الفصل الأول": "term1",
            "الفصل الثاني": "term2"
            # add more mappings here
        }

        assessment_mapping = {
            "التقويم الأول": "assessment1",
            "التقويم الثاني": "assessment2",
            "التقويم الثالث": "assessment3",
            "التقويم الرابع": "assessment4",
            # add more mappings here
        }

        students_marks = []
        students_info= students_names
        name_and_marks = []
        all_marks= all1

        for student_data in students_info:
            student_marks = {
                'id': int(student_data['student_id']), 
                'name': student_data['student_name'],
                'term1': {'assessment1': '', 'assessment2': '', 'assessment3': '', 'assessment4': ''},
                'term2': {'assessment1': '', 'assessment2': '', 'assessment3': '', 'assessment4': ''}
            }
            for mark_data in all_marks['data']:
                if mark_data['student_id'] == student_data['student_id']:
                    term_key = term_mapping.get(mark_data['assessment_period']['academic_term'])
                    if term_key is not None:
                        assessment_key = assessment_mapping.get(mark_data['assessment_period']['name'])
                        if assessment_key is not None:
                            student_marks[term_key][assessment_key] = mark_data['marks']
            students_marks.append(student_marks)

        sorted_students_names_and_marks = sorted(students_marks, key=lambda x: x['name'])
        
        sheet[f"D{int(context[str(page)].split(':')[0][1:])-5 }"].set_value(f' الصف: {class_name}')
        sheet[f"I{int(context[str(page)].split(':')[0][1:])-5 }"].set_value(f'الشعبة (   {class_char}    )')    
        sheet[f"O{int(context[str(page+1)].split(':')[0][1:])-5}"].set_value(sub_name)
              
        for counter,student_info in enumerate(sorted_students_names_and_marks, start=1):
            if counter >= 26:
                page += 2
                counter = 1
                
                sheet[f"D{int(context[str(page)].split(':')[0][1:])-5}"].set_value(f' الصف: {class_name}')
                sheet[f"I{int(context[str(page)].split(':')[0][1:])-5}"].set_value(f'الشعبة (   {class_char}    )')  
                sheet[f"O{int(context[str(page+1)].split(':')[0][1:])-5}"].set_value(sub_name)
                #    المادة الدراسية     
                
                # {'id': 3824166, 'name': 'نورالدين محمود راضي الدغيمات', 'term1': {'assessment1': 9, 'assessment2': 10, 'assessment3': 11, 'assessment4': 20}}
                
                for student_info in sorted_students_names_and_marks[25:] :
                    row_idx = counter + int(context[str(page)].split(':')[0][1:]) - 1  # compute the row index based on the counter
                    sheet[f"A{row_idx}"].set_value(name_counter)
                    sheet[f"B{row_idx}"].set_value(student_info['name'])
                    if 'term1' in student_info and 'assessment1' in student_info['term1'] and 'assessment2' in student_info['term1'] and 'assessment3' in student_info['term1'] and 'assessment4' in student_info['term1']:
                        sheet[f"D{row_idx}"].set_value(student_info['term1']['assessment1']) 
                        sheet[f"E{row_idx}"].set_value(student_info['term1']['assessment2']) 
                        sheet[f"F{row_idx}"].set_value(student_info['term1']['assessment3'])
                        sheet[f"G{row_idx}"].set_value(student_info['term1']['assessment4'])
                    if 'term2' in student_info:
                        row_idx2 = counter + int(context[str(page+1)].split(':')[0][1:]) - 1  # compute the row index based on the counter 
                        sheet[f"L{row_idx2}"].set_value(student_info['term2']['assessment1']) 
                        sheet[f"M{row_idx2}"].set_value(student_info['term2']['assessment2']) 
                        sheet[f"N{row_idx2}"].set_value(student_info['term2']['assessment3'])
                        sheet[f"O{row_idx2}"].set_value(student_info['term2']['assessment4'])                       
                    counter += 1
                    name_counter += 1              
                break                    
            row_idx = counter + int(context[str(page)].split(':')[0][1:]) - 1  # compute the row index based on the counter
            sheet[f"A{row_idx}"].set_value(name_counter)
            sheet[f"B{row_idx}"].set_value(student_info['name']) 
            if 'term1' in student_info and 'assessment1' in student_info['term1'] and 'assessment2' in student_info['term1'] and 'assessment3' in student_info['term1'] and 'assessment4' in student_info['term1']:
                sheet[f"D{row_idx}"].set_value(student_info['term1']['assessment1']) 
                sheet[f"E{row_idx}"].set_value(student_info['term1']['assessment2']) 
                sheet[f"F{row_idx}"].set_value(student_info['term1']['assessment3'])
                sheet[f"G{row_idx}"].set_value(student_info['term1']['assessment4'])
            if 'term2' in student_info:
                row_idx2 = counter + int(context[str(page+1)].split(':')[0][1:]) - 1  # compute the row index based on the counter 
                sheet[f"L{row_idx2}"].set_value(student_info['term2']['assessment1']) 
                sheet[f"M{row_idx2}"].set_value(student_info['term2']['assessment2']) 
                sheet[f"N{row_idx2}"].set_value(student_info['term2']['assessment3'])
                sheet[f"O{row_idx2}"].set_value(student_info['term2']['assessment4'])                
            name_counter += 1 
        name_counter = 1
        page += 2

    
    for i in classes: 
        modified_classes.append(mawad_representations(i))
        
    modified_classes = ' ، '.join(modified_classes)
    mawad = sorted(set(mawad))
    mawad = ' ، '.join(mawad)

    custom_shapes = {
        'modeeriah': f'لواء {modeeriah}',
        'hejri1': hejri1,
        'hejri2': hejri2,
        'melady1': melady1,
        'melady2': melady2,
        'baldah': baldah,
        'school': school_name,
        'classes': modified_classes,
        'mawad': mawad,
        'teacher' : teacher,
        'modeeriah_20_2': f'لواء {modeeriah}',
        'hejri_20_1': hejri1,
        'hejri_20_2': hejri2,
        'melady_20_1': melady1,
        'melady_20_2': melady2,
        'baldah_20_2': baldah,
        'school_20_2': school_name,
        'classes_20_2': modified_classes,
        'mawad_20_2': mawad,
        'teacher_20_2': teacher ,
        'modeeriah_20_1': f'لواء {modeeriah}',
        'hejri1': hejri1,
        'hejri2': hejri2,
        'melady1': melady1,
        'melady2': melady2,
        'baldah_20_1': baldah,
        'school_20_1': school_name,
        'classes_20_1': modified_classes,
        'mawad_20_1': mawad,
        'teacher_20_1': teacher
    }
    # FIXME: make the customshapes crop _20_ to the rest of the key in the custom_shapes
    # Iterate through the cells of the sheet and fill in the values you want
    doc.save()
            
    return custom_shapes 

def mawad_representations(string):
    y = {'روضة - 1': 'ر1', 'روضة - 2': 'ر2', 'الصف الأول': '1', 'الصف الثاني': '2', 'الصف الثالث': '3', 'الصف السابع': '7', 'الصف الثامن': '8', 'الصف التاسع': '9', 'الصف الرابع': '4', 'الصف الخامس': '5', 'الصف السادس': '6', 'الصف العاشر': '10', 'الصف الحادي عشر العلمي': '11', 'الصف الثاني عشر العلمي': '12 علمي', 'الصف الحادي عشر الأدبي': '11 ادبي', 'الصف الثاني عشر الأدبي': '12 ادبي', 'الصف الحادي عشر الشرعي': '11 شرغي', 'الصف الثاني عشر الشرعي': '12 شرعي', 'الصف الحادي عشر الصحي': '11 صحي', 'الصف الثاني عشر الصحي': '12 صحي', 'الصف الحادي عشر - إدارة معلوماتية': '11 ادارة', 'الصف الثاني عشر - إدارة معلوماتية': '12 ادارة', 'الصف الحادي عشر - اقتصاد منزلي': '11 اقتصاد', 'الصف الثاني عشر - اقتصاد منزلي': '12 اقتصاد', 'الصف الحادي عشر- فندقي': '11 فندقي', 'الصف الثاني عشر - فندقي': '12 فندقي', 'الصف الحادي عشر - صناعي': '11 صناعي', 'الصف الثاني عشر - صناعي': '12 صناعي', 'الصف الحادي عشر - زراعي': '11 زراعي', 'الصف الثاني عشر - زراعي': '12 زراعي'}

    search_str ,class_num = string.split('-')[0] ,string.split('-')[1]

    for key, value in y.items():
        search_key = search_str
        if search_key in key:
            replacement = value
            search_str = search_str.replace(search_key, replacement)

    return f'{search_str} - {class_num}'

def get_students_marks(auth,period_id,sub_id,instit_class_id,instit_id):
    '''
    دالة لاستدعاء علامات الطلاب و اسمائهم 
    و عواملها التوكن رقم السنة التعريفي ورقم المادة التعريفي و رقم المؤسسة و  رقم الصف التعريفي
    و تعود باسماء الطالب و علاماتهم
    '''
    url = f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id={sub_id}&institution_classes_id={instit_class_id}&institution_id={instit_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects'
    return make_request(url,auth)

def get_assessments_periods(auth ,term, assessment_id):
    '''
         استعلام عن تعريفات التقويمات في الفصل الدراسي 
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود بمعلومات عن تقيمات الصفوف في السنة الدراسية  
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentPeriods.json?_finder=academicTerm[academic_term:{term}]&assessment_id={assessment_id}&_limit=0"
    return make_request(url,auth)

def get_all_assessments_periods(auth , assessment_id):
    '''
         استعلام عن تعريفات التقويمات في السنة الدراسية 
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود بمعلومات عن كل تقيمات الصفوف في السنة الدراسية  
    '''
    terms = get_AcademicTerms(auth=auth , assessment_id=assessment_id)['data']
    season_assessments = []
    dic =  {'SEname': '', 'AssesName': '' ,'AssesId': '' }
    for term in terms:
        for asses in get_assessments_periods(auth, term['name'], assessment_id=assessment_id)['data']:
            dic = {'SEname': asses["academic_term"], 'AssesName': asses["name"], 'AssesId': asses["id"]}
            season_assessments.append(dic)
    return season_assessments
    
def get_assessments_id( auth ,education_grade_id ):
    '''
         استعلام عن تعريفي الصف الدراسي 
          عوامل الدالة تعريفي المرحلة الدراسية و التوكن
        تعود بمعلومات عن تقيمات الصفوف في السنة الدراسية  
    '''
    assessments = get_assessments(auth)
    for assessment in assessments['data'] : 
        if assessment['education_grade_id'] == education_grade_id :
            return assessment['id']

def get_AcademicTerms(auth,assessment_id):
    '''
    دالة لاستدعاء اسم الفصل 
    و عواملها التوكن و رقم تقيم الصف 
    و تعود باسماء الفصول على شكل جيسن
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentPeriods.json?_finder=uniqueAssessmentTerms&assessment_id={assessment_id}&_limit=0"
    return make_request(url,auth)        

def draw_rect_top(page, page_width, fill_color , width=50):
    """
    رسم مستطيل على الجزء العلوي من الصفحة
    """
    rect_top = fitz.Rect(0, 0, page_width, width)
    page.draw_rect(rect_top, color=fill_color, fill=fill_color)

def draw_rect_bottom(page, page_width, page_height, fill_color, width=50):
    """
    رسم مستطيل على الجزء السفلي من الصفحة
    """
    rect_bottom = fitz.Rect(0, page_height - width, page_width, page_height)
    page.draw_rect(rect_bottom, color=fill_color, fill=fill_color)

def draw_rect_left(page, page_height, fill_color, width=50):
    """
    رسم مستطيل على الجزء الأيسر من الصفحة
    """
    rect_left = fitz.Rect(0, 0, width, page_height)
    page.draw_rect(rect_left, color=fill_color, fill=fill_color)

def draw_rect_right(page, page_width, page_height, fill_color, width=50):
    """
    رسم مستطيل على الجزء الأيمن من الصفحة
    """
    rect_right = fitz.Rect(page_width - width, 0, page_width, page_height)
    page.draw_rect(rect_right, color=fill_color, fill=fill_color)

def add_margins(input_pdf1, output_pdf ,color_name="#8cd6e6",top_rec=50,bottom_rec=50,left_rec=50,right_rec=50 ,page=0):
    """
    إضافة هوامش باللون 8cd6e6 إلى جميع الجوانب من الصفحة
    """
    
    '''
    example of how to add colored margin for the first and scond page
    add_margins("existing_file.pdf", "output_file.pdf",top_rec=27, bottom_rec=20, left_rec=90, right_rec=120)
    add_margins("output_file.pdf", "output_file2.pdf",page=1 , top_rec=60, bottom_rec=25, left_rec=90, right_rec=120)
    '''
    # Open the PDF file
    input_pdf = fitz.open(input_pdf1)
    
    # Get the first page
    page = input_pdf[page]
    # Get the page dimensions
    page_width = page.rect.width
    page_height = page.rect.height

    # Convert the color from hex to RGB
    color_name = color_name
    color = webcolors.hex_to_rgb(color_name)
    color = tuple(c / 255 for c in color)  # Convert to RGB values between 0 and 1

    # Set the color
    fill_color = color  # Color code in RGB format

    # Draw rectangles on all four sides of the page
    draw_rect_top(page, page_width, fill_color , top_rec)
    draw_rect_bottom(page, page_width, page_height, fill_color, bottom_rec)
    draw_rect_left(page, page_height, fill_color , left_rec)
    draw_rect_right(page, page_width, page_height, fill_color, right_rec)

    # Save the modified PDF file
    input_pdf.save(output_pdf)

def mawad(string):
    y = {'روضة - 1': 'ر1', 'روضة - 2': 'ر2', 'الصف الأول': '1', 'الصف الثاني': '2', 'الصف الثالث': '3', 'الصف السابع': '7', 'الصف الثامن': '8', 'الصف التاسع': '9', 'الصف الرابع': '4', 'الصف الخامس': '5', 'الصف السادس': '6', 'الصف العاشر': '10', 'الصف الحادي عشر العلمي': '11', 'الصف الثاني عشر العلمي': '12 علمي', 'الصف الحادي عشر الأدبي': '11 ادبي', 'الصف الثاني عشر الأدبي': '12 ادبي', 'الصف الحادي عشر الشرعي': '11 شرغي', 'الصف الثاني عشر الشرعي': '12 شرعي', 'الصف الحادي عشر الصحي': '11 صحي', 'الصف الثاني عشر الصحي': '12 صحي', 'الصف الحادي عشر - إدارة معلوماتية': '11 ادارة', 'الصف الثاني عشر - إدارة معلوماتية': '12 ادارة', 'الصف الحادي عشر - اقتصاد منزلي': '11 اقتصاد', 'الصف الثاني عشر - اقتصاد منزلي': '12 اقتصاد', 'الصف الحادي عشر- فندقي': '11 فندقي', 'الصف الثاني عشر - فندقي': '12 فندقي', 'الصف الحادي عشر - صناعي': '11 صناعي', 'الصف الثاني عشر - صناعي': '12 صناعي', 'الصف الحادي عشر - زراعي': '11 زراعي', 'الصف الثاني عشر - زراعي': '12 زراعي'}

    search_str ,class_num = string.split('-')[0] ,string.split('-')[1]

    for key, value in y.items():
        search_key = search_str
        if search_key in key:
            replacement = value
            search_str = search_str.replace(search_key, replacement)

    return f'{search_str}-{class_num}'

def get_basic_info (username , password):
    auth = get_auth(username ,password )
    user = user_info(auth , username)
    inst_data = inst_name(auth)['data'][0]['Institutions']
    school_name = inst_data['name']
    inst_id= inst_name(auth)['data'][0]['Institutions']['id']
    baldah = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={inst_id}&_contain=InstitutionLands.CustomFieldValues')['data'][0]['address'].split('-')[0]
    grades= make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Education.EducationGrades?_limit=0')
    modeeriah = inst_area(auth)['data'][0]['Areas']['name']
    school_year = get_curr_period(auth)['data']
    melady = str(school_year[0]['end_year'])+' '+str(school_year[0]['start_year'])
    hejri =  str(hijri_converter.convert.Gregorian(school_year[0]['end_year'], 1, 1).to_hijri().year)+' '+str(hijri_converter.convert.Gregorian(school_year[0]['start_year'], 1, 1).to_hijri().year)
    teacher = user['data'][0]['name'].split(' ')[0]+' '+user['data'][0]['name'].split(' ')[-1]

def fill_custom_shape(doc, sheet_name, custom_shape_values, outfile):
    '''
    custom_shapes = {
        'modeeriah': f'لواء {modeeriah}',
        'hejri': hejri,
        'melady': melady,
        'baldah': baldah,
        'school': school_name,
        'classes': "7أ ، 7ب",
        'mawad': "اللغة الانجليزية",
        'teacher' : teacher
    }

    fill_custom_shape('official_marks_doc_a3_two_face.ods', 'الغلاف الداخلي', custom_shapes, 'tttttt.ods')
    '''
    print(doc)
    # Load the document
    doc = load(str(doc))
    try:
        # Iterate over the sheets in the document
        for sheet in doc.spreadsheet.childNodes[1:-1]:
            # Check if the sheet is the one we want (replace 'Sheet2' with the name of your sheet)
            if sheet.getAttribute('name') == sheet_name:
                # Get the custom shapes on the sheet
                custom_shapes = sheet.getElementsByType(CustomShape)
                for custom_shape in custom_shapes:     
                    # Check if the custom shape name is in the dictionary of custom shape values
                    if custom_shape.getAttribute('name')  in custom_shape_values:
                        # get the style name
                        p_style = custom_shape.childNodes[0].attributes.get(('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'style-name'), 'default_style')
                        span_style = custom_shape.childNodes[0].childNodes[0].attributes.get(('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'style-name'), 'default_style')
                        # clear the text
                        clear_text_custom_shape(custom_shape) 
                        value = custom_shape_values[custom_shape.getAttribute('name')]
                        # Create the text:p element inside the custom shape
                        text_p = P(stylename=str(p_style))
                        text_p.addElement(Span(text=str(value), stylename=str(span_style)))
                        # Add a new text paragraph to the custom shape with the new value
                        custom_shape.addElement(text_p)  
    except:
        pass
    # Save the modified document
    doc.save(outfile)
    
def clear_text_custom_shape(shape):
    # Remove all child nodes from the shape element
    while len(shape.childNodes) > 0:
        shape.removeChild(shape.childNodes[0])

def get_sheet_custom_shapes(document , sheet_name):
    # Load the document
    doc = load(str(document))
    # Loop through all sheets in the document
    for sheet in doc.spreadsheet.childNodes[1:-1]:      
        # Check if the sheet is the one we want (replace 'Sheet2' with the name of your sheet)
        if sheet.getAttribute('name') == str(sheet_name) :        
            # Get the text boxes on the sheet
            custom_shapes = sheet.getElementsByType(CustomShape)
            return [custom_shape.getAttribute("name") for custom_shape in custom_shapes]

def get_ods_sheets (doc='official_marks_doc_a3_two_face.ods'):
    # Load the ODF document
    doc = load(doc)
    # Get the sheets in the document
    sheets = doc.getElementsByType(Table)
    return [sheet.getAttribute("name") for sheet in sheets]

def page_counter_official_marks_doc_a3_two_face ():
    dual_page_dic = {}
    counter = 46
    start_cell = 6
    
    for i in range(3,47,2):
        print ( i , counter)
        print ( counter-1 , i+1 )
        if counter == 46 :
            dual_page_dic.update({ f'{counter}' : f'A{start_cell}:A{start_cell+24}'})
            dual_page_dic.update({ f'{i+1}' : f'A{start_cell+33}:A{start_cell+33+24}'})
            
            dual_page_dic.update({ f'{i}' : f'L{start_cell}:L{start_cell+24}'})
            dual_page_dic.update({ f'{counter-1}' : f'L{start_cell+33}:L{start_cell+33+24}'})
            
            start_cell +=33*2-1      
            counter -= 2
        elif i == 23:
            dual_page_dic.update({ f'{counter}' : f'A{start_cell}:A{start_cell+24}'})
            dual_page_dic.update({ f'{i+1}' : f'A{start_cell+32}:A{start_cell+32+24}'})
            
            dual_page_dic.update({ f'{i}' : f'L{start_cell}:L{start_cell+24}'})
            dual_page_dic.update({ f'{counter-1}' : f'L{start_cell+32}:L{start_cell+32+24}'}) 
                            
            break
        else:
            print(start_cell)
            dual_page_dic.update({ f'{counter}' : f'A{start_cell}:A{start_cell+24}'})
            dual_page_dic.update({ f'{i+1}' : f'A{start_cell+32}:A{start_cell+32+24}'})
            
            dual_page_dic.update({ f'{i}' : f'L{start_cell}:L{start_cell+24}'})
            dual_page_dic.update({ f'{counter-1}' : f'L{start_cell+32}:L{start_cell+32+24}'})            
            start_cell += 32*2
            counter -= 2
        # print(dual_page_dic)
        # input('press anything to continue')        
        
    print(dual_page_dic)

def generate_pdf(doc_path, path , rename_number):

    subprocess.call(['soffice',
                 # '--headless',
                 '--convert-to',
                 'pdf',
                 '--outdir',
                 path,
                 doc_path])
    subprocess.call(['mv',
                    f'{path}/generated.pdf' ,
                    f'{path}/send{rename_number}.pdf'])
    
def word2pdf(wordFile ,pdfFile):
    convert(wordFile , pdfFile)

def fill_doc(template , context , output):
    doc = DocxTemplate(template)
    context = context
    doc.render(context)
    doc.save(output)
    
def word_variables(template):
    doc = DocxTemplate(template)
    return doc.get_undeclared_template_variables()

def my_jq(data):
    json_str = json.dumps(data, indent=4, sort_keys=True, ensure_ascii=False).encode('utf8')
    return highlight(json_str.decode('utf8'), JsonLexer(), TerminalFormatter())

def make_request(url, auth):
    headers = {"Authorization": auth, "ControllerAction": "Results"}
    controller_actions = ["Results", "SubjectStudents", "Dashboard", "Staff",'StudentAttendances','SgTree']
    
    for controller_action in controller_actions:
        headers["ControllerAction"] = controller_action
        response = requests.request("GET", url, headers=headers)
        if "403 Forbidden" not in response.text :
            return response.json()
        
    return ['Some Thing Wrong']

def get_auth(username , password):
    ' دالة تسجيل الدخول للحصول على الرمز الخاص بالتوكن و يستخدم في header Authorization'
    url = "https://emis.moe.gov.jo/openemis-core/oauth/login"
    payload = {
        "username": username,
        "password": password
    }
    response = requests.request("POST", url, data=payload )

    if response.json()['data']['message'] == 'Invalid login creadential':
        return False
    else: 
        return response.json()['data']['token']    
    
def inst_name(auth):
    '''
    استدعاء اسم المدرسة و الرقم الوطني و الرقم التعريفي 
        عوامل الدالة الرابط و التوكن
        تعود بالرقم التعريفي و الرقم الوطني و اسم المدرسة 
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Staff?_limit=1&_contain=Institutions&_fields=Institutions.code,Institutions.id,Institutions.name"
    return make_request(url,auth)

def inst_area(auth , inst_id = None ):
    '''
    استدعاء لواء المدرسة و المنطقة
    عوامل الدالة الرابط و التوكن
    تعود باسم البلدية و اسم المنطقة و اللواء 
    '''
    if inst_id is None:
        inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Institutions.json?id={inst_id}&_contain=AreaAdministratives,Areas&_fields=AreaAdministratives.name,Areas.name"
    return make_request(url,auth)

def user_info(auth,username):
    '''
        استدعاء معلومات عن المعلم او المستخدم 
        عوامل الدالة الرابط و التوكن و رقم المستخدم
        تعود برقم المستخدم الوطني و اسمه الرباعي  
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/User-Users?username={username}&is_staff=1&_fields=id,username,openemis_no,first_name,middle_name,third_name,last_name,preferred_name,email,date_of_birth,nationality_id,identity_type_id,identity_number,status&_limit=1"
    return make_request(url,auth)

def get_teacher_classes1(auth,ins_id,staff_id,academic_period):
    '''
        استدعاء معلومات صفوف المعلم 
        عوامل الدالة الرابط و التوكن و التعريفي للمدرسة و تعريفي الفترة و staffid 
        تعود الدالة بتعريفي اي صف مع المعلم و كود الصف
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStaff?institution_id={ins_id}&staff_id={staff_id}&academic_period_id={academic_period}&_contain=InstitutionSubjects&_limit=0&_fields=InstitutionSubjects.id,InstitutionSubjects.education_subject_id,InstitutionSubjects.name"
    return make_request(url,auth)

def get_teacher_classes2(auth,inst_sub_id):
    '''
    استدعاء معلومات تفصيلية عن الصفوف 
    عوامل الدالة الرابط و التوكن و رقم المستخدم
    تعود باسم الصف و تعريفي الصف و عدد الطلاب في الصف و اسم المادة التي يدرسها المعلم في الصف
    '''
    # url = "https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere=institution_subject_id:10513896,institution_subject_id:10513912,institution_subject_id:10513928,institution_subject_id:10513944"
    url = f"https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere=institution_subject_id:{inst_sub_id}"
    
    return make_request(url,auth)

def get_class_students(auth,academic_period_id,institution_subject_id,institution_class_id,institution_id):
    '''
    استدعاء معلومات عن الطلاب في الصف
    عوامل الدالة هي الرابط و التوكن و تعريفي الفترة الاكاديمية و تعريفي مادة المؤسسة و تعريفي صف المؤسسة و تعريفي المؤسسة
    تعود بمعلومات تفصيلية عن كل طالب في الصف بما في ذلك اسمه الرباعي و التعريفي و مكان سكنه
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id,student_status_id,Users.id,Users.username,Users.openemis_no,Users.first_name,Users.middle_name,Users.third_name,Users.last_name,Users.address,Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_type_id,Users.identity_number,Users.external_reference,Users.status,Users.is_guardian&_limit=0&academic_period_id={academic_period_id}&institution_subject_id={institution_subject_id}&institution_class_id={institution_class_id}&institution_id={institution_id}&_contain=Users"
    return make_request(url,auth)

def enter_mark(auth
               ,marks
               ,assessment_grading_option_id
               ,assessment_id
               ,education_subject_id
               ,education_grade_id
               ,institution_id
               ,academic_period_id
               ,institution_classes_id
               ,student_status_id
               ,student_id
               ,assessment_period_id):
    '''
    دالة لادخال علامة الطالب 
    عوامل الدالة العلامة و رقم المؤسسة التعريفي و رقم الطالب و الرقم التعريفي للفترة الاكاديمة و رقم المادة التعريفي
    enter_mark(get_auth() 
                ,marks= 6
                ,assessment_grading_option_id= 8
                ,assessment_id= 188
                ,education_subject_id= 4
                ,education_grade_id= 275
                ,institution_id= 2600
                ,academic_period_id= 13
                ,institution_classes_id= 786120
                ,student_status_id= 1
                ,student_id= 3768676
                ,assessment_period_id= 624)
    و تعود الدالة بكود الاجابة 200 و اذا لم يعود به تصدر الدالة خطا
    '''
    url = 'https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentItemResults.json'
    headers = {"Authorization": auth , "ControllerAction" : "Results" }
    json_data = {
        'marks':marks,
        'assessment_grading_option_id':assessment_grading_option_id,
        'assessment_id':assessment_id,
        'education_subject_id':education_subject_id,
        'education_grade_id':education_grade_id,
        'institution_id':institution_id,
        'academic_period_id':academic_period_id,
        'institution_classes_id':institution_classes_id,
        'student_status_id':student_status_id,
        'student_id':student_id,
        'assessment_period_id':assessment_period_id,
        'action_type': 'default',
    }

    response = requests.post(url,headers=headers,json=json_data,)
    print(response.status_code)
    if response.status_code != 200:
        raise(Exception("couldn't enter the mark for some reason")) 

def get_curr_period(auth):
    '''
    دالة  تستدعي معلومات السنة الحالية من الخادم
    التوكن 
    و تعود على المستخدم بمعلومات السنة الدراسية الحالية 
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/AcademicPeriod-AcademicPeriods?current=1&_fields=id,code,start_date,end_date,start_year,end_year,school_days"
    return make_request(url,auth)

def get_assessments(auth,academic_term,assessment_id):
    '''
    دالة تستدعي معلومات عن الامتحانات في الفصل
    و عواملها اسم الفصل و تعريفي اختبار المرحلة 
    تعود بمعلومات عن الامتحانات المتوفرة على المنظومة في الفصل
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentPeriods.json?_finder=academicTerm[academic_term:{academic_term}]&assessment_id={assessment_id}&_limit=0"
    return make_request(url,auth)

def get_sub_info(auth,class_id,assessment_id,academic_period_id,institution_id):
    '''
    استدعاء معلومات عن مواد الصف
    و عواملها هي تعريفي الصف و تعريفي مرحلة الاختبار و الفترة الاكاديمية و تعريفي المؤسسة
    تعود بمعلومات عن مواد الصف و اهمها تعريفي المادة و كود المادة
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentItems.json?_finder=subjectNewTab[class_id:{class_id};assessment_id:{assessment_id};academic_period_id:{academic_period_id};institution_id:{institution_id}]&_limit=0"
    return make_request(url,auth)

def side_marks_document(username , password):
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    user_id = user_info(auth , username)['data'][0]['id']
    years = get_curr_period(auth)
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id)['data']]
    classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i])['data'] for i in range(len(classes_id_1))]
    classes_id_3 = []  
    for class_info in classes_id_2:
        classes_id_3.append([{"institution_class_id": class_info[0]['institution_class_id'] ,"sub_name": class_info[0]['institution_subject']['name'],"class_name": class_info[0]['institution_class']['name']}])
        
    for v in range(len(classes_id_1)):
        # id
        print (classes_id_3[v][0]['institution_class_id'])
        # subject name 
        print (classes_id_3[v][0]['sub_name'])
        # class name
        print (classes_id_3[v][0]['class_name'])
        students = get_class_students(auth
                                    ,period_id
                                    ,classes_id_1[v]
                                    ,classes_id_3[v][0]['institution_class_id']
                                    ,inst_id)
        students_names = sorted([i['user']['name'] for i in students['data']])
        print(students_names)
        
        context={}
        counter = 0
        for name in students_names :
            context[f'name{counter}'] = str(name) 
            counter+=1            
        context[f'sub'] = str(classes_id_3[v][0]['class_name']) 
        context[f'class_name'] = str(classes_id_3[v][0]['sub_name']) 
        context[f'school'] = str(inst_name(auth)['data'][0]['Institutions']['name']) 
        context['y1'] = str(years['data'][0]['start_year'])
        context['y2'] = str(years['data'][0]['end_year'])

        fill_doc('./templet_files/side_marks_note.docx' , context , f'./send_folder/send{v}.docx' )
        context.clear()
        generate_pdf(f'./send_folder/send{v}.docx' , './send_folder' ,v)
        # input("press enter to continue")
        # return students_names

def insert_students_names_in_excel_marks(template , students_id_and_names , outfile):
    workbook = load_workbook(filename=template)
    sheet = workbook.active
    counter = 2
    for i in students_id_and_names:
        sheet[f'B{counter}'] = i['student_name']
        sheet[f'A{counter}'] = i['student_id']
        counter+=1
    workbook.save( filename = outfile )

def delete_empty_rows(file , outfile):
    workbook = load_workbook(filename=file)
    sheet = workbook.active
    
    # Find the last row of data in the sheet
    max_row = sheet.max_row
    # breakpoint()
    # Loop through each row in reverse order and check if it is empty
    for row in range(max_row, 1, -1):
        if all([cell.value in (None, '') for cell in sheet[row]][:6]):
            # Remove the row if it is empty
            sheet.delete_rows(row, 1)
            
    # Compute sum for each row
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=7).value = f"=SUM(C{row}:F{row})"
    # Set header for sum column
    sheet.cell(row=1, column=7).value = "المجموع"
    # Auto-fit column width for sum column
    sheet.column_dimensions['G'].auto_size = True
    
    # Save the updated workbook
    workbook.save(outfile)

def read_excel_marks(file):
        workbook = load_workbook(filename=file)
        sheet = workbook.active
        counter = 2
        for value in sheet.values:
            if value[0] ==None:
                break
            elif not value[2] == None :
                value = list(value)
#                   التقويم الرابع و  الثالث و  الثاني و   الاول  
#                 value[2]+ value[3]+ value[4]+value[5]
                value[6]= value[2]+ value[3]+ value[4]+value[5]
                print(value)                
            else : 
                print(value)
                
def insert_students_names_and_marks(assessments_json, students_id_and_names , template , outfile):
    workbook = load_workbook(filename=template)
    sheet = workbook.active
    marks_and_name = []
    dic =  {'Sid':'' ,'Sname': '', 'ass1': '' ,'ass2': '' , 'ass3': '' , 'ass4': '' }
    for i in students_id_and_names:   
#         print(i['student_id'])
        for v in assessments_json['data']:
            if v['student_id'] == i['student_id'] :  
                dic['Sid'] = i['student_id'] 
                dic['Sname'] = i['student_name'] 
                if v['assessment_period']['name'] == 'التقويم الأول':
                    dic['ass1'] = v["marks"]     
                elif v['assessment_period']['name'] == 'التقويم الثاني':
                    dic['ass2'] = v["marks"]             
                elif v['assessment_period']['name'] == 'التقويم الثالث':
                    dic['ass3'] = v["marks"]           
                elif v['assessment_period']['name'] == 'التقويم الرابع':
                    dic['ass4']= v["marks"]
        marks_and_name.append(dic)
        dic =  { 'Sid':'' ,'Sname': '', 'ass1': '' ,'ass2': '' , 'ass3': '' , 'ass4': '' }

        headers = ['id', 'اسم الطالب', 'التقويم الاول', 'التقويم الثاني', 'التقويم الثالث', 'التقويم الرابع']
        for col_num, header in enumerate(headers, start=1):
            sheet.cell(row=1, column=col_num, value=header)
        # Iterate over the data and insert into rows
        for row_num, row_data in enumerate(marks_and_name, start=2):
            for col_num, cell_data in enumerate(row_data.values(), start=1):
                sheet.cell(row=row_num, column=col_num, value=cell_data)
                
    workbook.save( filename = outfile )
    delete_empty_rows(outfile , outfile)

def create_excel_sheets_marks(username, password ):
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    user_id = user_info(auth , username)['data'][0]['id']
    years = get_curr_period(auth)
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id)['data']]
    classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i])['data'] for i in range(len(classes_id_1))]
    classes_id_3 = []  
    for class_info in classes_id_2:
        classes_id_3.append([{"institution_class_id": class_info[0]['institution_class_id'] ,"sub_name": class_info[0]['institution_subject']['name'],"class_name": class_info[0]['institution_class']['name']}])

    for v in range(len(classes_id_1)):
        # id
        print (classes_id_3[v][0]['institution_class_id'])
        # subject name 
        print (classes_id_3[v][0]['sub_name'])
        # class name
        print (classes_id_3[v][0]['class_name'])
        students = get_class_students(auth
                                    ,period_id
                                    ,classes_id_1[v]
                                    ,classes_id_3[v][0]['institution_class_id']
                                    ,inst_id)
        # students_names = sorted([i['user']['name'] for i in students['data']])
        # print(students_names)
        students_id_and_names = []
        for IdAndName in students['data']:
            students_id_and_names.append({'student_name': IdAndName['user']['name'] , 'student_id':IdAndName['student_id']})

        assessments_json = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id=4&institution_classes_id='+ str(classes_id_3[v][0]['institution_class_id'])+ f'&institution_id={inst_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects')
        insert_students_names_and_marks(assessments_json , students_id_and_names , './templet_files/excel_marks.xlsx' , './send_folder/' + classes_id_3[v][0]['class_name'] + '.xlsx')

def count_files():
    files = glob.glob('./send_folder/*')
    return files

def delete_send_folder():
    files = glob.glob('./send_folder/*')
    for f in files:
        os.remove(f)

def get_students_marks(auth,period_id,sub_id,instit_class_id,instit_id):
    '''
    دالة لاستدعاء علامات الطلاب و اسمائهم 
    و عواملها التوكن رقم السنة التعريفي ورقم المادة التعريفي و رقم المؤسسة و  رقم الصف التعريفي
    و تعود باسماء الطالب و علاماتهم
    '''
    url = f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id={sub_id}&institution_classes_id={instit_class_id}&institution_id={instit_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects'
    return make_request(url,auth)

def main():
    print('starting script')
    
    # ods_file = 'send1.ods'
    # copy_ods_file('./templet_files/official_marks_doc_a3_two_face.ods' , f'./send_folder/{ods_file}')
    # outdir = '.'
    # path = '/opt/programming/school_programms1/telegram_bot/send_folder/انس الجعافره-9971055725.xlsx'
    # lst = Read_E_Side_Note_Marks(path)
    auth=get_auth(9971055725,9971055725)
    # print(make_request(auth=auth,url='://emis.moe.gov.jo/openemis-core​/restful​/v2​/Area-AreaAdministratives.json'))
    # output = make_request(auth=auth,url='://emis.moe.gov.jo/openemis-core/restful/Institution.StudentAbsencesPeriodDetails?institution_id=2600&academic_period_id=13&_limit=0&_fields=student_id,institution_id,academic_period_id,institution_class_id,education_grade_id,date,period,comment,absence_type_id')
    # print(output)

    # fill_official_marks_a3_two_face_doc2_offline_version(9971055725,9971055725,lst)
    # fill_official_marks_doc_wrapper_offline(9971055725,9971055725,lst)
    # create_e_side_marks_doc(9971055725,9971055725)
    
if __name__ == "__main__":
    main()