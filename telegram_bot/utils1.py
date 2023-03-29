# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import geckodriver_autoinstaller
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

def get_all_assessments_periods_data(auth , assessment_id):
    '''
         استعلام عن تعريفات التقويمات في السنة الدراسية و امكانية تحرير التقويم و  العلامة القصوى و الدنيا
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود تعريفات التقويمات في السنة الدراسية و امكانية تحرير التقويم و  العلامة القصوى و الدنيا  
    '''
    terms = get_AcademicTerms(auth=auth , assessment_id=assessment_id)['data']
    season_assessments = []
    dic =  {'SEname': '', 'AssesName': '' ,'AssesId': '' , 'pass_mark': '' , 'max_mark' : '' , 'editable' : ''}
    min_max=[]
    for i in assessments_periods_min_max_mark(get_auth(9991014194,9991014194) , 187, 3)['data']:
        min_max.append({'id': i['assessment_period_id'] , 'pass_mark':i['assessment_grading_type']['pass_mark'] , 'max_mark' : i['assessment_grading_type']['max'] } )                    
    for term in terms:
        for asses in get_assessments_periods(auth, term['name'], assessment_id=assessment_id)['data']:
            dic = {'SEname': asses["academic_term"], 'AssesName': asses["name"], 'AssesId': asses["id"] , 'pass_mark': [dictionary for dictionary in min_max if dictionary.get('id') == 623][0]['pass_mark'] , 'max_mark' : [dictionary for dictionary in min_max if dictionary.get('id') == 623][0]['max_mark'] , 'editable':asses['editable']}
            season_assessments.append(dic)
    return season_assessments

def assessments_periods_min_max_mark(auth , assessment_id , education_subject_id ):
    '''
         استعلام عن القيمة القصوى و الدنيا لكل التقويمات  
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود بمعلومات عن تقيمات الصفوف في السنة الدراسية  
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentItemsGradingTypes.json?_contain=EducationSubjects,AssessmentGradingTypes.GradingOptions&assessment_id={assessment_id}&education_subject_id={education_subject_id}&_limit=0"
    return make_request(url,auth)

def get_all_assessments_periods2(auth , assessment_id):
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
        dic['assessment_id'] = get_grade_id_from_assessment_id(auth , class_info[0]['institution_subject']['education_grade_id'])
        dic['education_subject_id'] = class_info[0]['institution_subject']['education_subject_id']
        dic['education_grade_id'] = class_info[0]['institution_subject']['education_grade_id']
        dic['institution_classes_id'] = class_info[0]['institution_class_id']
        dic['students_ids'] = get_class_students_ids(auth,period_id,class_info[0]['institution_subject_id'],class_info[0]['institution_class_id'],inst_id)

        required_data_to_enter_marks.append(dic)
    
    return required_data_to_enter_marks

def get_grade_id_from_assessment_id(auth , grade_id):
    
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
        new_ws.title = classes_id_3[v][0]['class_name']+'-'+str(classes_id_3[v][0]['institution_class_id'])
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
    user = user_info(auth , 9971055725)
    school_name = inst_name(auth)['data'][0]['Institutions']['name']
    baldah = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id=2600&_contain=InstitutionLands.CustomFieldValues')['data'][0]['address'].split('-')[0]
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

def make_request(url,auth):
    headers = {"Authorization": auth}
    response = requests.request("GET", url, headers=headers)
    if "403 Forbidden" in response.text :
        headers["ControllerAction"] = "Results"        
        response = requests.request("GET", url, headers=headers)        
    elif "403 Forbidden" in  response.text :
        headers["ControllerAction"] = "SubjectStudents"        
        response = requests.request("GET", url, headers=headers)                   
    else :
        headers["ControllerAction"] = "Staff"            
        response = requests.request("GET", url, headers=headers)             
    return response.json()

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

def inst_area(auth):
    '''
    استدعاء لواء المدرسة و المنطقة
    عوامل الدالة الرابط و التوكن
    تعود باسم البلدية و اسم المنطقة و اللواء 
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Institutions.json?id=2600&_contain=AreaAdministratives,Areas&_fields=AreaAdministratives.name,Areas.name"
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

def get_AcademicTerms(auth,assessment_id):
    '''
    دالة لاستدعاء اسم الفصل 
    و عواملها التوكن و رقم تقيم الصف 
    و تعود باسماء الفصول على شكل جيسن
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentPeriods.json?_finder=uniqueAssessmentTerms&assessment_id={assessment_id}&_limit=0"
    return make_request(url,auth)

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
    fill_official_marks_doc_wrapper(9971055725,9971055725 )
    # create_excel_sheets_marks(9971055725,9971055725 )
    # create_e_side_marks_doc(9971055725,9971055725)
    # enter_marks_arbitrary(9981054126,123456,665,10,17)
    
if __name__ == "__main__":
    main()