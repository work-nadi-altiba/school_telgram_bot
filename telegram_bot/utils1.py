# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import geckodriver_autoinstaller
import requests
import json
from pygments import highlight
from pygments.lexers import JsonLexer
from pygments.formatters import TerminalFormatter
from openpyxl import Workbook , load_workbook
# from docxtpl import DocxTemplate
# from docx2pdf import convert
import subprocess
import os 
import glob

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

def get_AcademicPeriods(auth,assessment_id):
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
    for i in data:
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

def main():
    print('starting script')
    # print(count_files())
    # side_marks_document(9971055725,9971055725)
    # create_excel_sheets_marks(9971055725,9971055725)
    # generate_pdf('./telegram_bot/generated.docx' , './telegram_bot' ,2)
    # print(user_info(auth=get_auth(9971055725,9971055725) , username=9971055725))
    # delete_empty_rows('/opt/programming/school_programms1/telegram_bot/send_folder/الصف السابع-أ.xlsx')

if __name__ == "__main__":
    main()