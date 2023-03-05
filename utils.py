# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import geckodriver_autoinstaller
import requests
import json
from pygments import highlight
from pygments.lexers import JsonLexer
from pygments.formatters import TerminalFormatter
# from docxtpl import DocxTemplate
# from docx2pdf import convert
import subprocess

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
        print ('Invalid login creadential')
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

        fill_doc('./test_files/side_marks_note.docx' , context , f'./test_files/send{v}.docx' )
        context.clear()
        generate_pdf('./test_files/generated.docx' , './test_files' ,v)
        input("press enter to continue")
        # return students_names

def main():
    side_marks_document(9971055725,9971055725)
    # generate_pdf('./telegram_bot/generated.docx' , './telegram_bot' ,2)

if __name__ == "__main__":
    main()