from selenium import webdriver
from selenium.webdriver.common.by import By
import geckodriver_autoinstaller
import requests

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

def inst_area(auth):
    '''
    استدعاء لواء المدرسة و المنطقة
    عوامل الدالة الرابط و التوكن
    تعود باسم البلدية و اسم المنطقة و اللواء 
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Institutions.json?id=2600&_contain=AreaAdministratives,Areas&_fields=AreaAdministratives.name,Areas.name"
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

def user_info(auth,username):
    '''
        استدعاء معلومات عن المعلم او المستخدم 
        عوامل الدالة الرابط و التوكن و رقم المستخدم
        تعود برقم المستخدم الوطني و اسمه الرباعي  
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/User-Users?username=9971055725&is_staff=1&_fields=id,username,openemis_no,first_name,middle_name,third_name,last_name,preferred_name,email,date_of_birth,nationality_id,identity_type_id,identity_number,status&_limit=1"
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

def get_teacher_classes1(auth,ins_id,staff_id,academic_period):
    '''
        استدعاء معلومات صفوف المعلم 
        عوامل الدالة الرابط و التوكن و التعريفي للمدرسة و تعريفي الفترة و staffid 
        تعود الدالة بتعريفي اي صف مع المعلم و كود الصف
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStaff?institution_id=2600&staff_id=3971236&academic_period_id=13&_contain=InstitutionSubjects&_limit=0&_fields=InstitutionSubjects.id,InstitutionSubjects.education_subject_id"
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

def get_teacher_classes2(auth,inst_sub_id):
    '''
    استدعاء معلومات تفصيلية عن الصفوف 
    عوامل الدالة الرابط و التوكن و رقم المستخدم
    تعود باسم الصف و تعريفي الصف و عدد الطلاب في الصف و اسم المادة التي يدرسها المعلم في الصف
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere=institution_subject_id:10513896,institution_subject_id:10513912,institution_subject_id:10513928,institution_subject_id:10513944"
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

def get_class_students(auth,academic_period_id,institution_subject_id,institution_class_id,institution_id):
    '''
    استدعاء معلومات عن الطلاب في الصف
    عوامل الدالة هي الرابط و التوكن و تعريفي الفترة الاكاديمية و تعريفي مادة المؤسسة و تعريفي صف المؤسسة و تعريفي المؤسسة
    تعود بمعلومات تفصيلية عن كل طالب في الصف بما في ذلك اسمه الرباعي و التعريفي و مكان سكنه
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id,student_status_id,Users.id,Users.username,Users.openemis_no,Users.first_name,Users.middle_name,Users.third_name,Users.last_name,Users.address,Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_type_id,Users.identity_number,Users.external_reference,Users.status,Users.is_guardian&_limit=0&academic_period_id=13&institution_subject_id=10513896&institution_class_id=786118&institution_id=2600&_contain=Users"
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


def main():
    pass

if __name__ == "__main__":
    main()