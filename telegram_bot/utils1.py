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
import tempfile
import zipfile
from PyPDF4 import PdfFileMerger
import datetime
from dateutil.relativedelta import relativedelta
import calendar 
import locale
from itertools import product
import pdb
import wfuzz
from tqdm import tqdm
from pprint import pprint
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import traceback
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

# Global variables should under her please
secondery_students = []

# New code should be under here please
def mark_all_students_as_present(auth ,term_days_dates ,r_data = None , proxies = None):

    session = requests.Session()
    if not r_data:
        r_data = get_required_data_to_enter_absent(auth,session=session)
    institution_id = r_data['institution_id']
    institution_class_id = r_data['institution_class_id']
    education_grade_id = r_data['education_grade_id']
    academic_period_id = r_data['academic_period_id']
    
    url = f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-StudentAttendances.json?_finder=classStudentsWithAbsenceSave[institution_id:{institution_id};institution_class_id:{institution_class_id};education_grade_id:{education_grade_id};academic_period_id:{academic_period_id};attendance_period_id:1;day_id:FUZZ;week_id:undefined;week_start_day:undefined;week_end_day:undefined;subject_id:0]&_limit=0'

    headers = [("User-Agent" , "python-requests/2.28.1"),("Accept-Encoding" , "gzip, deflate"),("Accept" , "*/*"),("Connection" , "close"),("Authorization" , f"{auth}"),("ControllerAction" , "StudentAttendances"),("Content-Type" , "application/json")]


    unsuccessful_requests = wfuzz_function(url , term_days_dates,headers,None,method='GET',proxies = proxies)

    while len(unsuccessful_requests) != 0:
        unsuccessful_requests = wfuzz_function(url , unsuccessful_requests,headers,None,method='GET',proxies = proxies)

    print("All requests were successful!")

def mark_students_absent_in_dates(auth ,students_id_with_names, absent_days_list ,institution_id , institution_class_id , education_grade_id , academic_period_id , year1 , year2 , helper=False ,proxies = None):
    '''
    usage:
        required_data = get_required_data_to_enter_absent(auth)
        # for i in f:
        #     print(f"{i} = required_data['{i}']",end=', ')
        mark_students_absent_in_dates(auth , id_with_names, absent_days_list, institution_id = required_data['institution_id'], institution_class_id = required_data['institution_class_id'], education_grade_id = required_data['education_grade_id'], academic_period_id = required_data['academic_period_id'], year1 = required_data['year1'], year2 = required_data['year2'] )
    '''
    students_dictionary = {}
    for idx, (name, student_id) in enumerate(students_id_with_names, start=1) : 
        students_dictionary[idx]= [student_id,name ]

    absent_days_list = [day.split('/') for day in absent_days_list]

    # students_dictionary
    absent_data = []

    for date in absent_days_list:
        student_name = str(students_dictionary[int(date[0])][0])
        student_id = str(students_dictionary[int(date[0])][1])
        year = year1 if int(date[-1]) in [9,10,11,12] else year2
        
        for day in date[1:-1]:
            date_str = f"{year}-{date[-1].zfill(2)}-{day.zfill(2)}"
            if helper:
                print ( student_name  , date_str)
            item = json.dumps({"student_id": student_id,
                                "institution_id": institution_id,
                                "academic_period_id": academic_period_id,
                                "institution_class_id": institution_class_id,
                                "absence_type_id": "2",
                                "student_absence_reason_id": None,
                                "comment": None,
                                "period": 1,
                                "date": date_str,
                                "subject_id": 0,
                                "education_grade_id": education_grade_id}).replace('}','')
            absent_data.append(item)

    # 'student_id': 7388854,
    # 'date': '2022-09-19',

    # {"student_id": 7388854, "institution_id": 2600, "academic_period_id": 13, "institution_class_id": 786118, "absence_type_id": "2", "student_absence_reason_id": null, "comment": null, "period": 1, "date": "2022-09-19", "subject_id": 0, "education_grade_id": 275,
    if helper:
        pprint(absent_data[0])


    headers = [("User-Agent" , "python-requests/2.28.1"),("Accept-Encoding" , "gzip, deflate"),("Accept" , "*/*"),("Connection" , "close"),("Authorization" , f"{auth}"),("ControllerAction" , "StudentAttendances"),("Content-Type" , "application/json")]

    body_postdata = json.dumps({ "action_type": "default"}).replace('{',' FUZZ ,')


    url = 'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-StudentAbsencesPeriodDetails.json?_limit=0'

    unsuccessful_requests = wfuzz_function(url , absent_data , headers ,body_postdata ,method='POST', proxies = proxies)

    while len(unsuccessful_requests) != 0:
        unsuccessful_requests = wfuzz_function(url , unsuccessful_requests , headers ,body_postdata ,method='POST',proxies = proxies)

    print("All requests were successful!")

def fill_students_absent_in_dates_wrapper(auth ,students_absent_multiline_string=None , random_values = False ,start_date_str = None , end_date_str = None , skip_dates_list = None , required_data=None, proxies = None):
    if not required_data:
        required_data = get_required_data_to_enter_absent(auth)
    id_with_names = get_names_for_absent_purposes(auth)
    if students_absent_multiline_string :
        absent_days_list = extract_absent_dates_from_text(students_absent_multiline_string)
    elif random_values :
        days_list = get_period_days_dates(start_date_str, end_date_str, skip_dates_list, skip_weekend=True)
        absent_days_list = create_random_absent_list(days_list , id_with_names)
    else:
        print('give me absent_multiline_string value or choose random_values')
    mark_students_absent_in_dates(auth , id_with_names, absent_days_list, institution_id = required_data['institution_id'], institution_class_id = required_data['institution_class_id'], education_grade_id = required_data['education_grade_id'], academic_period_id = required_data['academic_period_id'], year1 = required_data['year1'], year2 = required_data['year2'] ,proxies = proxies)

def fill_all_students_present_wrapper(auth , start_date_str , end_date_str , skip_dates_list ,required_data=None ,proxies = None):
    days_list = get_period_days_dates(start_date_str, end_date_str, skip_dates_list, skip_weekend=True)
    mark_all_students_as_present(auth , days_list , required_data ,proxies = proxies)

def erase_students_absent_dates(auth ,required_data=None ,helper=False,proxies = None):
    if not required_data:
        required_data = get_required_data_to_enter_absent(auth)
    absent_data_list = get_class_absent_days_with_id(auth ,required_data=required_data)
    
    institution_id=required_data['institution_id']
    institution_class_id=required_data['institution_class_id']
    education_grade_id=required_data['education_grade_id']
    academic_period_id=required_data['academic_period_id']
    
    absent_data = []    
    
    for date_item in absent_data_list:
        student_id = str(date_item[0])
        date_str = date_item[1]

        item = json.dumps({"student_id": student_id,
                            "institution_id": institution_id,
                            "academic_period_id": academic_period_id,
                            "institution_class_id": institution_class_id,
                            "absence_type_id": "0",
                            "student_absence_reason_id": None,
                            "comment": None,
                            "period": 1,
                            "date": date_str,
                            "subject_id": 0,
                            "education_grade_id": education_grade_id}).replace('}','')
        absent_data.append(item)

    # 'student_id': 7388854,
    # 'date': '2022-09-19',

    # {"student_id": 7388854, "institution_id": 2600, "academic_period_id": 13, "institution_class_id": 786118, "absence_type_id": "2", "student_absence_reason_id": null, "comment": null, "period": 1, "date": "2022-09-19", "subject_id": 0, "education_grade_id": 275,
    if helper:
        pprint(absent_data[0])


    headers = [("User-Agent" , "python-requests/2.28.1"),("Accept-Encoding" , "gzip, deflate"),("Accept" , "*/*"),("Connection" , "close"),("Authorization" , f"{auth}"),("ControllerAction" , "StudentAttendances"),("Content-Type" , "application/json")]

    body_postdata = json.dumps({ "action_type": "default"}).replace('{',' FUZZ ,')


    url = 'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-StudentAbsencesPeriodDetails.json?_limit=0'

    unsuccessful_requests = wfuzz_function(url , absent_data , headers ,body_postdata ,method='POST',proxies = proxies)

    while len(unsuccessful_requests) != 0:
        unsuccessful_requests = wfuzz_function(url , unsuccessful_requests , headers ,body_postdata ,method='POST',proxies = proxies)

    print("All requests were successful!")

def get_class_absent_days_with_id(auth ,simple_list=True , required_data = None):
    if not required_data:
        required_data = get_required_data_to_enter_absent(auth)
    
    institution_id=required_data['institution_id']
    institution_class_id=required_data['institution_class_id']
    education_grade_id=required_data['education_grade_id']
    academic_period_id=required_data['academic_period_id']
    
    url = f'https://emis.moe.gov.jo/openemis-core/restful/Institution.StudentAbsencesPeriodDetails?institution_id={institution_id}&institution_class_id={institution_class_id}&education_grade_id={education_grade_id}&academic_period_id={academic_period_id}&_limit=0&_fields=student_id,institution_id,academic_period_id,institution_class_id,education_grade_id,date,period,comment,absence_type_id'
    absent_data = make_request(auth=auth , url=url)
    if simple_list:
        absent_data = [[i['student_id'] , i['date']] for i in absent_data['data']]
    return absent_data

def create_random_absent_list(dates_list ,id_with_names):
    number_of_the_students = len(id_with_names)
    random_absent = []
    for date in dates_list:
        month , day = date.split('-')[1] , date.split('-')[2] 
        number_of_absency = random.randint(1, 8)
        random_students = []
        for _ in range(number_of_absency):
            random_student_index = random.randint(1, number_of_the_students)
            random_students.append(f'{random_student_index}/{day}/{month}')
        # print(random_students)
        random_absent.extend(random_students)
    return random_absent

def extract_absent_dates_from_text(text , helper=False):
    '''
    نص الغياب يجب ان يتبع القواعد الاتية :
    ان غياب كل طالب في سطر حسب رقمه على الاسماء التي تستخرجها دالة get_names_for_absent_purposes
    الغياب يكون فيه اليوم ثم الشهر 
    اذا كان الغياب بين ايام متواصلة استعمل علامة الناقص بين اليوم الاول و اليوم الاخير
    اذا كانت الايام متفرقة استعمل بين كل يوم و يوم علامة الفاصلة 
    كالمثال التالي:
        8/11 5,7/10 17-19/10 23,25/10
    المخرجات ستكون:
    1    /    24    /     9
    الشهر      اليوم      رقم
    الطالب
    المتسلسل        
    '''
    absent_days_list = []
    absent_string_list = text.split('\n')
    for idx ,item in enumerate(absent_string_list ,start=1):
        if helper:
            print(idx)
        dates = item.split(' ')
        for date in dates :
            if '-' in date:
                start , end = date.split('/')[0].split('-')
                month = date.split('/')[1]
                e = [f"{idx}/{i}/{month}" for i in range(int(start), int(end)+1)] 
                absent_days_list.extend(e)
                if helper:
                    print(' '.join(e),end=' ')              
            elif ',' in date :
                days ,month= date.split('/')
                splitted_days = days.split(',')
                d = [f"{idx}/{i}/{month}" for i in splitted_days]
                absent_days_list.extend(d)
                if helper:
                    print(' '.join(d) ,end=' ')
            else :
                absent_days_list.append(f'{idx}/{date}')
                if helper:
                    print(f'{idx}/{date}' ,end=' ')
        if helper:            
            print('\n'+'-'*50)
    return absent_days_list

def get_period_days_dates(start_date_str, end_date_str, skip_dates_list=[], skip_weekend=True):
    '''
    اعطاء الايام الموجودة في فترة من الزمن و استثناء ايام العطل الرسمية بتوجيه المستخدم
    # Example usage:
    start_date_str = "2023-08-20"
    end_date_str = "2023-11-12"
    skip_dates_list = ["2023-09-27"]  # Specify dates to skip in "Y-m-d" format
    
    result_dates = get_period_days_dates(start_date_str, end_date_str, skip_dates_list, skip_weekend=True)

    # len(result_dates)
    print('\n'.join(result_dates))
    # print(result_dates)
    '''
    start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d")
    skip_dates_list = [datetime.datetime.strptime(date_str, "%Y-%m-%d").date() for date_str in skip_dates_list]
    result_dates = []
    current_date = start_date
    delta = datetime.timedelta(days=1)

    while current_date <= end_date:
        if current_date.date() not in skip_dates_list:
            if not (skip_weekend and current_date.weekday() in [4, 5]):
                result_dates.append(current_date.strftime("%Y-%m-%d"))
        
        current_date += delta

    return result_dates

def contains_else_number_or_slash(text):
    # Define a regular expression pattern to match characters other than "/", digits, and whitespace
    pattern = re.compile(r'[^/\d\s]')
    
    # Search for the pattern in the text
    match = pattern.search(text)
    
    # Return True if a match is found (i.e., special characters are present), False otherwise
    return match is not None

def intended_for_pytest_for_the_absent_text(absent_days_list):
    for i in absent_days_list:
        if contains_else_number_or_slash(i):
            print(i)
    # get the monthes of the proccessed text 
    # لاحضار الاشهر التي تحتاج الى تعديل او المختلفة 
    # set([i.split('/')[2] for i in l])

def get_names_for_absent_purposes(auth , session=None):
    d = get_required_data_to_enter_absent(auth=auth)
    institution_id = d['institution_id']
    institution_class_id = d['institution_class_id']
    academic_period_id = d['academic_period_id']
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id,student_status_id,Users.id,Users.username,Users.openemis_no,Users.first_name,Users.middle_name,Users.third_name,Users.last_name,Users.address,Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_type_id,Users.identity_number,Users.external_reference,Users.status,Users.is_guardian&_limit=0&academic_period_id={academic_period_id}&institution_class_id={institution_class_id}&institution_id={institution_id}&_contain=Users"
    students_with_ids = make_request(url=url,auth=auth,session=session)
    u_names_with_ids = set([(i['student_id'] ,i['user']['name']) for i in students_with_ids['data']])

    sorted_list = sorted(u_names_with_ids, key=lambda x: x[1])
    
    return sorted_list

def get_required_data_to_enter_absent(auth , session=None):
    url = 'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-InstitutionClasses'
    classes = make_request(url=url , auth=auth , session=session)
    curr_per = get_curr_period(auth=auth, session=session)['data'][0]
    academic_period_id =curr_per['id']
    home_room_class = [i for i in classes['data'] if i['academic_period_id'] ==academic_period_id][0]

    institution_class_id = home_room_class['id']
    institution_id = home_room_class['institution_id']

    url = f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-InstitutionClasses.json?_finder=gradesByInstitutionAndAcademicPeriodAndInstitutionClass[institution_id:{institution_id};academic_period_id:{academic_period_id};institution_class_id:{institution_class_id}]&_limit=0'

    education_grade_id = make_request(url=url , auth=auth, session=session)['data'][0]['id']

    return {
            'institution_id':institution_id,
            'institution_class_id':institution_class_id,
            'education_grade_id':education_grade_id,
            'academic_period_id':academic_period_id,
            'year1': curr_per['start_year'],
            'year2': curr_per['end_year']
            }

def bulk_e_side_note_marks(passwords):
    for p in passwords.split('\n'):
        # print(p,'-------<>')
        try : 
            username = p.split('/')[0]
            password = p.split('/')[1]
        except:
            username ,password =[p]*2
        # print(username , password)
        # FIXME: صلح مشكلة السيشين في الريكيوست
        # session = requests.Session()
        try:
            create_e_side_marks_doc(username , password )
        except Exception as e:
            
            print("\033[91m There is error in \n{}/{}\033[00m" .format(username , password))
            # print(username , password)
            traceback.print_exc()
        # if not get_auth(username , password): 
        #     print(username , password)

def read_all_xlsx_in_folder(directory_path='./send_folder'):
    dic_list = []
    for item in os.listdir(directory_path):
        item_path = os.path.join(directory_path, item)  
        dic_list.append(Read_E_Side_Note_Marks_xlsx(file_path=item_path))
    return dic_list

def convert_to_marks_offline_from_send_folder(directory_path='./send_folder',do_not_delete_send_folder=True , template='./templet_files/official_marks_doc_a3_two_face_white_cover.ods' , color ="#8cd6e6"):
    dic_list = read_all_xlsx_in_folder(directory_path)
    for file_content in dic_list:
        fill_official_marks_doc_wrapper_offline(file_content , do_not_delete_send_folder=do_not_delete_send_folder , templet_file=template ,color=color)

def fill_student_absent_doc_wrapper(username, password ,template='./templet_files/new_empty_absence_notebook_doc_white_cover.ods' , outdir='./send_folder/' ,teacher_full_name=False):
    student_details = get_student_statistic_info(username,password,teacher_full_name=teacher_full_name)
    fill_student_absent_doc_name_days_cover(student_details , template , outdir )

def vacancies_dictionary2Html():
    from jinja2 import Template
    # from mydicts import dict_list1 ,dict_list2

    table_data = dict_list1 + dict_list2


    # Define the HTML table code as a string
    table_template = '''
    <style>
    table {
    border-collapse: collapse;
    width: 100%;
    margin-left: auto;
    margin-right: 0;
    font-size: 20px;
    font-family: 'Times New Roman', Times, serif;
    }

    th, td {
    border: 1px solid black;
    padding: 8px;
    }

    </style>

    <table dir="rtl">
    <thead>
    <tr>
    <th style="text-align: right;">اسم المدرسة</th>
    <th style="text-align: right;">انصبة المدرسة</th>
    <th style="text-align: right;">المعلمين</th>
    <th style="text-align: right;">الصفوف</th>
    </tr>
    </thead>
    <tbody>
    {% for item in data %}
    <tr>
    <td style="text-align: right;">{{ item['school_name'] }}</td>
    <td style="text-align: right;">{{ item['school_load'] | replace("\n", "<br>") }}</td>
    <td style="text-align: right;">{{ item['teachers'] | replace("\n", "<br>") }}</td>
    <td style="text-align: right;">{{ item['classes'] | replace("\n", "<br>") }}</td>
    </tr>
    {% endfor %}
    </tbody>
    </table>
    '''

    # format the data into the table template
    table_html = Template(table_template).render(data=table_data)

    html = f'''
    <html lang="ar">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>الشواغر</title>
    </head>
    <body>
    {table_html}
    </body>
    </html>
    '''



    # Specify the file path and name
    file_path = "تشكيلات.html"

    # Open the file in write mode
    with open(file_path, "w") as file:
        # Write the content to the file
        file.write(html)

    # Confirmation message
    print(f"Content saved to {file_path}.")

def tor_code():
    '''
    دالة لمتصفح تور كتبتها لكي اتمكن من معالجة مشكلة السيرفر الذي يحتاج مني ان يكون عنوان جهازي امريكي
    '''
    import stem.process
    from stem.util import term
    from stem import Signal
    from stem.control import Controller
    import subprocess
    import requests 

    SOCKS_PORT = 9050
    # CTL_SOCKS_PORT = 9051

    def get_tor_nodes_with_prefix(prefix , CTL_SOCKS_PORT = 9051):

        with Controller.from_port(port=CTL_SOCKS_PORT) as controller:
            controller.authenticate(password="MyStr0n9P#D")  # Authenticate with the Tor control port

            # Retrieve the list of Tor nodes
            relay_list = controller.get_network_statuses()

            nodes_with_prefix = []

            for relay in relay_list:
                if relay.flags and 'Exit' in relay.flags:  # Filter exit relays
                    if relay.address and relay.address.startswith(prefix):  # Filter IP address prefix
                        nodes_with_prefix.append(relay)

            return nodes_with_prefix

    # Example usage
    # nodes = get_tor_nodes_with_prefix('23.')
    # for node in nodes:
    #     print("Fingerprint:", node.fingerprint)
    #     print("IP Address:", node.address)

    with subprocess.Popen(['tor'], stdout=subprocess.PIPE, universal_newlines=True) as tor:
        bootstrapped = False
        for line in tor.stdout:
            if not bootstrapped and 'Bootstrapped 100%' not in line:
                continue
            elif not bootstrapped:
                bootstrapped = True
            print(line.strip())
            nodes = get_tor_nodes_with_prefix('23.')
            tor.terminate()
            break
    print(nodes)
    # Start an instance of Tor configured to only exit through Russia. This prints
    # Tor's bootstrap information as it starts. Note that this likely will not
    # work if you have another Tor instance running.

    def print_bootstrap_lines(line):
        if "Bootstrapped " in line:
            print(term.format(line, term.Color.BLUE))


        print(term.format("Starting Tor:\n", term.Attr.BOLD))

    for node in nodes :
        tor_process = stem.process.launch_tor_with_config(
        config = {
            'SocksPort': str(SOCKS_PORT),
            'ExitNodes': node.fingerprint ,
        },
        init_msg_handler = print_bootstrap_lines,
        )

        print(term.format("\nChecking our endpoint:\n", term.Attr.BOLD))
        
        print(term.format(node.address,term.Color.GREEN))

        try:
            auth = get_auth(9971055725,9971055725 , proxies={"http": "socks5://127.0.0.1:9050", "https": "socks5://127.0.0.1:9050"})
            print(term.format( auth, term.Color.BLUE))
        except:
            print(term.format('error occured' , term.Color.RED))

        tor_process.kill()  # stops tor

def get_year_days_dates(start_date=None , end_date=None , skip_start_date=None , skip_end_date=None):

    present_days = []
    start_date = datetime.date(2022, 8, 30) if not start_date else datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.date(2023, 6, 30) if not end_date else datetime.strptime(end_date , "%Y-%m-%d")
    skip_start_date = datetime.date(2023, 1, 1) if not skip_start_date else datetime.strptime(skip_start_date , "%Y-%m-%d")
    skip_end_date = datetime.date(2023, 2, 5) if not skip_end_date else datetime.strptime(skip_end_date , "%Y-%m-%d")

    current_date = start_date
    while current_date <= end_date:
        if current_date < skip_start_date or current_date > skip_end_date:
            if current_date.weekday() not in [4, 5]:  # Exclude Friday (4) and Saturday (5)
                present_days.append(current_date.strftime("%Y-%m-%d"))
        current_date += datetime.timedelta(days=1)

    return present_days

def group_students(dic_list , i = None):
    # sort the list based on the 'class_name' key
    sorted_list = sorted(dic_list, key=lambda x: x['student_class_name_letter'])

    # group the sorted list by the 'student_class_name_letter' key
    grouped_list = []
    for key, group in itertools.groupby(sorted_list, key=lambda x: x['student_class_name_letter']):
        group_list = list(group)
        if all(x.get('student_class_name_letter') for x in group_list ):
            grouped_list.append(group_list)
    if i :
        for i in grouped_list:
            print(len(i),i[0]['student_class_name_letter'])
        return 0
    else : 
        return grouped_list

def wfuzz_function(url, fuzz_list,headers,body_postdata,method='POST',proxies = None):
    """دالة استخدمها لارسال طلب بوست بشكل سريع

    Args:
        fuzz_list (list): قائمة في بيانات الطلاب المراد ادخالها
        headers (tuple-list): راسيات الطلب او الركويست
        body_postdata (str): جسم البوست داتا
        method (str, optional): طريقة الطلب. Defaults to 'POST'.

    Returns:
        any : تعود بقائمة الطلبات غير الناجحة
    """    
    unsuccessful_requests=[]
    with tqdm(total=len(fuzz_list), bar_format='{postfix[0]} {n_fmt}/{total_fmt}',
            postfix=["uploaded mark", {"value": 0}]) as t:
            s = wfuzz.get_payloads([fuzz_list])
            for idx , r in enumerate(s.fuzz(
                            url=url ,
                            # hc=[404] , 
                            # payloads=[("list",fuzz_list)] ,
                            headers=headers ,
                            postdata = body_postdata ,
                            proxies= proxies ,
                            method= method
                            ),start =1):
                    
                t.postfix[1]["value"] = idx
                t.update()    
            #     print(r)
            #     print(r.content)
            #     print(r.history.code) # كود الركويست
                if r.history.code != 200 :
                    unsuccessful_requests.append(r.description)
    return unsuccessful_requests

def upload_marks_optimized(username , password , classess_data , empty = False):
    '''
    file_name = 'علي المحاميد-9901024120(6).ods'
    student_details = Read_E_Side_Note_Marks_ods('./'+file_name)
    
    fuzz_list = upload_marks_optimized(9901024120 , 9901024120 , student_details ,empty=False)
    
    هذه الدالة خطيرة و تحتاج الى تغيير بعض الاشياء و التاكد من الاشياء التالية
    1- تحتاج الى التكاد من جسم الرد على الطلب ان يكون جيسون و ان لا يكون فيه خطاء
    2- تغير وايل لوب جملة التكرار بينما الى جملة التكرار فور و تكرار الكود داخلها اقصى حد خمس مرات و الخروج من جملة التكرار
    '''
    fuzz_postdata_list = []
    session = requests.Session()
    auth = get_auth(username , password)
    period_id = classess_data['custom_shapes']['period_id']
    school_id = classess_data['custom_shapes']['school_id']
    # term1_assessment_codes = ['S1A1', 'S1A2', 'S1A3', 'S1A4']
    # term2_assessment_codes = ['S2A1', 'S2A2', 'S2A3', 'S2A4']
    assessment_codes = ['S1A1', 'S1A2', 'S1A3', 'S1A4' , 'S2A1', 'S2A2', 'S2A3', 'S2A4']
    assessment_code_dic = {'S1A1': {'term' :'term1' , 'assess' : 'assessment1'},
                            'S1A2': {'term' :'term1' , 'assess' : 'assessment2'},
                            'S1A3': {'term' :'term1' , 'assess' : 'assessment3'},
                            'S1A4': {'term' :'term1' , 'assess' : 'assessment4'},
                            'S2A1': {'term' :'term2' , 'assess' : 'assessment1'},
                            'S2A2': {'term' :'term2' , 'assess' : 'assessment2'},
                            'S2A3': {'term' :'term2' , 'assess' : 'assessment3'},
                            'S2A4': {'term' :'term2' , 'assess' : 'assessment4'}}
    
    assessments_periods_data = classess_data['required_data_for_mrks_enter']
    for class_data in classess_data['file_data']:
        class_id = class_data['class_name'].split('=')[2] 
        class_subject = class_data['class_name'].split('=')[3]
        class_name = classess_data['file_data'][1]['class_name'].split('=')[0]
        if 'عشر' not in class_name : 
            students_marks_ids = class_data['students_data']
            assessment_grade_id = assessments_periods_data[int(class_id)]['assessment_grade_id']
            grade_id = assessments_periods_data[int(class_id)]['grade_id']
            assessment_periods = get_editable_assessments(auth,username,assessment_grade_id,class_subject,session=session)
            # assessment_ids = assessments_periods_data[int(class_id)]['assessments_period_ids']
            # s1a1, s1a2, s1a3, s1a4, s2a1, s2a2, s2a3, s2a4 = [assessment_ids[i] if i < len(assessment_ids) else None for i in range(8)]
            for student_info in students_marks_ids:
                for code in assessment_codes:
                    if len([i for i in assessment_periods if code in i['code']]) != 0:
                        assessment_period_id = [i for i in assessment_periods if code in i['code']][0]['AssesId']
                        term = assessment_code_dic[code]['term']
                        assess = assessment_code_dic[code]['assess']
                        term_marks = student_info[term]
                        mark = '' if empty else term_marks.get(assess)
                        fuzz_postdata = {
                                'marks': '' if mark== '' else str("{:.2f}".format(float(mark))),
                                'assessment_id': assessment_grade_id,
                                'education_subject_id': class_subject,
                                'education_grade_id': grade_id,
                                'institution_classes_id': class_id,
                                'student_id': student_info['id'],
                                'assessment_period_id': assessment_period_id,
                                'action_type': 'default'
                            }
                        fuzz_postdata_list.append(json.dumps(fuzz_postdata).replace('{','').replace('}',''))
                        
    body_postdata = json.dumps({
            'assessment_grading_option_id': 8,
            'institution_id': school_id,
            'academic_period_id': period_id,
            'student_status_id': 1,
            'action_type': 'default'}).replace('}',', FUZZ }')

    headers = [("User-Agent" , "python-requests/2.28.1"),("Accept-Encoding" , "gzip, deflate"),("Accept" , "*/*"),("Connection" , "close"),("Authorization" , f"{auth}"),("ControllerAction" , "Results"),("Content-Type" , "application/json")]
    
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentItemResults.json"
    
    unsuccessful_requests = wfuzz_function(url , fuzz_postdata_list,headers,body_postdata)

    while len(unsuccessful_requests) != 0:
        unsuccessful_requests = wfuzz_function(unsuccessful_requests,headers,body_postdata)

    print("All requests were successful!")
    
def read_json_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        dictionary = json.load(file)
    return dictionary

def save_dictionary_to_json_file(dictionary, file_path='./send_folder/output.json', indent=None):

    with open(file_path, 'w', encoding='utf-8') as file:
        json.dump(dictionary, file, indent=indent, ensure_ascii=False)
        
def create_coloured_certs_wrapper(username , password ,term2=False):
    session = requests.Session()
    auth = get_auth(username , password)
    student_info_marks = get_students_info_subjectsMarks( username , password , session=session)
    students_statistics_assesment_data = get_student_statistic_info(username,password , student_ids=[i['student_id'] for i in get_school_students_ids(auth)] , session=session)
    
    dic_list4 = student_info_marks
    grouped_list = group_students(dic_list4 )
    
    
    add_subject_sum_dictionary(grouped_list)

    add_averages_to_group_list(grouped_list ,skip_art_sport=False)
    students_statistics_assesment_data['assessments_data'] = grouped_list
    
    save_dictionary_to_json_file(dictionary=students_statistics_assesment_data)
    
    create_coloured_certs_ods(students_statistics_assesment_data , term2=term2)

def convert_files_to_pdf(outdir,pages_range):
    """داله تقوم بتحويل الملفات في مجلد الى صيغة pdf 

    Args:
        outdir (string): اسم المجلد الذي تريد تحويل الملفات منه
    """    
    files = os.listdir(outdir)

    for file in files:
        if not file.endswith(".json"):
            if not pages_range:
                subprocess.run(['soffice', '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', outdir, '--page-ranges', '1-2', f'{outdir}/{file}'])
            else:
                subprocess.run(['soffice', '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', outdir, f'{outdir}/{file}'])

def column_index_from_string(column_string):
    """Converts a column letter to a column index."""
    index = 0
    for i, char in enumerate(column_string):
        index = index * 26 + (ord(char.upper()) - ord('A')) + 1
    return index

def get_column_letter(column_index):
    """
    Convert a column index to the corresponding column letter.
    """
    div = column_index +1
    column_letter = ""
    while div > 0:
        modulo = (div - 1) % 26
        column_letter = chr(65 + modulo) + column_letter
        div = (div - modulo) // 26
    return column_letter

def strategy_ladder ( ):
    doc = ezodf.opendoc('./ods_strategy_ladder/SL-5,6,7.ods')
    sheet = doc.sheets[0] # Assuming you want to work with the first sheet

    # Access cell values
    cell = sheet['AM1']
    cells = cell.value.split('-')
    firstCell, secondCell, beforeLastCell = cells[0].split('/')[0], cells[0].split('/')[1], cells[0].split('/')[-2]
    thirdCell = column_index_from_string(secondCell)+1
    arb_fill_cells = cells[1].split('.')

    # function parameter variables
    outdir = './send_folder'
    file_name = 'عمرو المطارنة-9981048186(2).ods'
    student_details = Read_E_Side_Note_Marks_ods('./'+file_name)
    term = 2


    # Iterate over student data
    for students_data_list in student_details['file_data']:
        class_name = students_data_list['class_name'].split('=')[0].replace('الصف', '')
        
        for row_idx, student_info in enumerate(students_data_list['students_data'], start=9):
            assessments = [
                student_info[f'term{term}']['assessment1'],
                student_info[f'term{term}']['assessment2'],
                student_info[f'term{term}']['assessment3'],
                student_info[f'term{term}']['assessment4']
                ]
            name_conter = row_idx - 8
            sheet[row_idx ,0].set_value(name_conter)
            sheet[row_idx ,1].set_value(student_info['name'])
            sheet[row_idx ,column_index_from_string(firstCell)].set_value(student_info[f'term{term}']['assessment1'])
            sheet[row_idx ,column_index_from_string(secondCell)].set_value(student_info[f'term{term}']['assessment2'])
            sheet[row_idx ,thirdCell].set_value(student_info[f'term{term}']['assessment3'])
            sheet[row_idx ,column_index_from_string(beforeLastCell)].set_value(student_info[f'term{term}']['assessment4'])
            sheet[row_idx ,column_index_from_string(beforeLastCell)+1].set_value(sum(int(assessment) if assessment != '' else 0 for assessment in assessments) )

            Assess1_generator = RandomNumberGenerator(student_info[f'term{term}']['assessment1'], RandomNumberGenerator.convert_to_ranges(arb_fill_cells)).generate_numbers_with_sum()
            Assess1_generator = Assess1_generator if Assess1_generator is not None else [''] * len(arb_fill_cells)
            Assess1_start_column = column_index_from_string(firstCell) - len(arb_fill_cells)
            for column_idx, value in enumerate(Assess1_generator, start=Assess1_start_column):
                sheet[row_idx ,column_idx].set_value(value)

            Assess2_generator = RandomNumberGenerator(student_info[f'term{term}']['assessment2'], RandomNumberGenerator.convert_to_ranges(arb_fill_cells)).generate_numbers_with_sum() 
            Assess2_generator = Assess2_generator if Assess2_generator is not None else [''] * len(arb_fill_cells)
            Assess2_start = column_index_from_string(secondCell) - len(arb_fill_cells)
            for column_idx, value in enumerate(Assess2_generator, start=Assess2_start):
                sheet[row_idx ,column_idx].set_value(value)

            Assess4_generator = RandomNumberGenerator(student_info[f'term{term}']['assessment4'], RandomNumberGenerator.convert_to_ranges([int(i) * 2 for i in arb_fill_cells])).generate_numbers_with_sum() 
            Assess4_generator = Assess4_generator if Assess4_generator is not None else [''] * len(arb_fill_cells)
            Assess4_start = column_index_from_string(beforeLastCell) - len(arb_fill_cells)
            for column_idx, value in enumerate(Assess4_generator, start=Assess4_start):
                sheet[row_idx ,column_idx].set_value(value)
            
            
        doc.saveas(f'{outdir}/{class_name}.ods')    

class RandomNumberGenerator:
    '''
    total_sum = 18
    numbers = [3, 2, 4, 5, 1, 3]
    ranges = RandomNumberGenerator.convert_to_ranges(numbers)  # ranges = [(0, 3), (0, 3), (0, 5), (0, 5), (0, 2), (0, 2)]
    
    generator = RandomNumberGenerator(total_sum, ranges)
    result = generator.generate_numbers_with_sum()
    print(result)
    '''
    def __init__(self, total_sum, ranges):
        self.total_sum = total_sum
        self.ranges = ranges
    
    def generate_numbers(self):
        numbers = [random.randint(minimum, maximum) for minimum, maximum in self.ranges]
        return numbers
    
    def check_sum(self, numbers):
        return sum(numbers) == self.total_sum
    
    def generate_numbers_with_sum(self):
        for numbers in product(*(range(minimum, maximum + 1) for minimum, maximum in self.ranges)):
            if self.check_sum(numbers):
                return numbers
    
    @staticmethod
    def convert_to_ranges(numbers):
        ranges = [(1, int(number)) for number in numbers]
        return ranges

def fill_student_absent_doc_name_days_cover(student_details , ods_file, outdir):
    doc = ezodf.opendoc(ods_file)
        
    sheet_name = 'Sheet1'
    sheet = doc.sheets[sheet_name]

    students_data_lists = student_details['students_info']
    class_name = student_details['class_name']
    context = {27 : 'Y69=AP123', 2 : 'A69=V123' ,3 : 'Y128=AP182', 26 : 'A128=V182' ,25 : 'Y186=AP240', 4 : 'A186=V240' ,5 : 'Y244=AP298', 24 : 'A244=V298' ,
                    23 : 'Y302=AP356', 6 : 'A302=V356' ,7 : 'Y360=AP414', 22 : 'A360=V414' ,21 : 'Y418=AP472', 8 : 'A418=V472' ,9 : 'Y476=AP530', 20 : 'A476=V530' ,
                    19 : 'Y534=AP588', 10 : 'A534=V588' ,11 : 'Y592=AP646', 18 : 'A592=V646' ,17 : 'Y650=AP704', 12 : 'A650=V704' ,13 : 'Y708=AP762', 16 : 'A708=V762' ,
                    15 : 'Y766=AP820', 14 : 'A766=V820' }

    year1 , year2 = student_details['year_code'].split('-')
    for i in range(183,820,58):
        sheet[f"E{i}"].set_value(f'{class_name.replace("الصف" ,"")}')
        sheet[f"AN{i}"].set_value(f'{year2} / {year1}')

    for counter,student_info in enumerate(students_data_lists, start=0):
        
        # row_idx = counter + int(context[str(page)].split(':')[0][1:]) - 1  # compute the row index based on the counter
        row_idx = counter + 69
        row_idx2 = counter + 128
        birth_data = student_info['birth_date'].split('-')
        years, months, days = calculate_age(student_info['birth_date'],student_details['start_date'] )
        
        sheet[f"G{row_idx}"].set_value(student_info['first_name'])
        sheet[f"I{row_idx}"].set_value(student_info['second_name'])
        sheet[f"K{row_idx}"].set_value(student_info['third_name'])
        sheet[f"M{row_idx}"].set_value(student_info['last_name'])
        sheet[f"O{row_idx}"].set_value(int(birth_data[2]))
        sheet[f"P{row_idx}"].set_value(int(birth_data[1]))
        sheet[f"Q{row_idx}"].set_value(birth_data[0])
        sheet[f"S{row_idx}"].set_value(student_info['birthPlace_area'])
        sheet[f"U{row_idx}"].set_value(student_info['nationality'])
        sheet[f"Y{row_idx2}"].set_value(student_info['religion'])
        
        sheet[f"AA{row_idx2}"].set_value(days)
        sheet[f"AB{row_idx2}"].set_value(months)
        sheet[f"AC{row_idx2}"].set_value(years)
        
        sheet[f"AH{row_idx2}"].set_value(student_info['guardian_name'])
        sheet[f"AJ{row_idx2}"].set_value(student_info['guardian_student_relationship'])
        sheet[f"AL{row_idx2}"].set_value(student_info['guardian_employment'])
        sheet[f"AN{row_idx2}"].set_value(student_info['guardian_phone_number'])
        sheet[f"AO{row_idx2}"].set_value(student_info['address'])
        sheet[f"AP{row_idx2}"].set_value(student_info['student_id']       )
        
    months_range = [14,16,18,20,22,24,4,6,8,10,12]

    for  counter, month in enumerate(months_range , start =1):
        section_one_row_start , section_one_row_end = int(re.findall(r'\d+',context[month].split('=')[0])[0])-2 , int(re.findall(r'\d+',context[month].split('=')[1])[0]) 
        section_two_row_start , section_two_row_end = int(re.findall(r'\d+',context[month+1].split('=')[0])[0])-2 , int(re.findall(r'\d+',context[month+1].split('=')[1])[0])
        
        if month in [26]:
            print('')
        
        if counter < 7 :
            year = int(year2) 
        elif counter == 7:
            year = int(year1)
            counter+=1        
        else:
            year = int(year1)
            counter+=1
        
        for column in range(8,24):
            day = column-7
            try :
                sheet[section_one_row_start, column-2].set_value( get_day_name_from_date(year , counter , day ) )
                
                if not (day ==1 and counter ==1) :
                    if ("سبت" in get_day_name_from_date(year , counter , day )) or ("جمعة" in get_day_name_from_date(year , counter , day )) : 
                        
                        for row in range(section_one_row_start+1 , section_one_row_end ):
                            # FIXME: sheet[row, column].fill = PatternFill(start_color="c0c0c0", fill_type="solid")
                            sheet[row, column-2].set_value('▒▒▒')
            except ValueError:
                pass
            # except AttributeError:
            #     print(section_one_row_start)

        for column in range(24,40):
                day = column-23+16
                try:
                    sheet[section_two_row_start, column].set_value( get_day_name_from_date(year , counter , column-7) )
                    
                    if not (day ==25 and counter ==2) :
                        if ("سبت" in get_day_name_from_date(year , counter , day )) or ("جمعة" in get_day_name_from_date(year , counter , day )) : 
                            
                                for row in range(section_two_row_start+1 , section_two_row_end ):
                                    # FIXME: sheet[row, column).fill = PatternFill(start_color="c0c0c0", fill_type="solid")
                                    sheet[row, column].set_value('▒▒▒')
                
                except ValueError:
                    pass
                # except AttributeError:
                #     print(section_two_row_start)         

    doc.saveas(outdir+'one_step_more.ods' )
    
    modeeriah = student_details['school_bridge'].replace('لواء ' , '')
    school_name = student_details['school_name_code'].split(' - ')[1].replace('مدرسة ', '')
    class_name = student_details['class_name'].split('-')[0].replace('الصف' , '')
    sec = student_details['class_name'].split('-')[1]
    teacher = student_details['teacher_incharge_name']
    year1 , year2 = student_details['year_code'].split('-')
    custom_shapes = {
        'modeeriah': f'لواء {modeeriah}',
        'school': school_name,
        'class': class_name,
        'sec': sec,
        'murabee' : teacher,
        'year' : f'{year1}  /  {year2}'
    }

    fill_custom_shape(doc= outdir+'one_step_more.ods', sheet_name='الغلاف', custom_shape_values=custom_shapes, outfile= outdir+f'/{class_name}-{sec}.ods')
    
    delete_file(outdir+'one_step_more.ods')

    # outdir = './send_folder'
    filename = f'{class_name}-{sec}.ods'
    command = f'soffice --headless --convert-to pdf:writer_pdf_Export --outdir {outdir} "{outdir}/{filename}"'
    os.system(command)
    
def get_day_name_from_date(year, month, day):
    # Set the locale to Arabic (Egypt)
    locale.setlocale(locale.LC_ALL, 'ar_EG.utf8')

    # Create a datetime object from the given year, month, and day
    date_object = datetime.date(year, month, day)

    # Get the day name in Arabic
    day_name_arabic = date_object.strftime('%A')

    return day_name_arabic

class InvalidPageRangeError(Exception):
    pass

def print_page_pairs(pair_pages=None,start_page=1 , end_page=None ):
    if pair_pages is not None:
        if pair_pages % 2 != 0 :
            raise InvalidPageRangeError("Invalid pair pages range: pair_pages should be even")
        pages_length = pair_pages
        counter = pair_pages*2 -2
    else :
        range_ = range(start_page, end_page+1)
        pages_length = int(len(range_)/2)
        pages_list = [i for i in range_]
        counter = len(pages_list)-2
        if len(pages_list) % 2 != 0 :
            raise InvalidPageRangeError("Invalid page range: end_page should be even")
        
    for i in range(start_page,pages_length+1):  
        if i % 2 == 0 :
            print(i ,' ----->' , f'{i + counter+1} - {i}')      
        else :
            print(i ,' ----->', f'{i} - {i + counter+1 }')
        counter -=2

def calculate_age(birth_date, target_date):
    '''
    # Example usage
    birth_date = datetime.date(2007, 6, 29)
    target_date = datetime.date(2022, 9,1)
    years, months, days = calculate_age(birth_date, target_date)

    # Print the age in years, months, and days
    print(f"Age on {target_date.strftime('%Y-%m-%d')}: {years} years, {months} months, {days} days")
    '''
    birth_date = datetime.datetime.strptime(birth_date, '%Y-%m-%d').date()
    target_date = datetime.datetime.strptime(target_date, '%Y-%m-%d').date()
    age = relativedelta(target_date, birth_date)
    return age.years, age.months, age.days

def fill_Template_With_basic_Student_info(student_details,template='./templet_files/كشف البيانات الاساسية للطلاب.xlsx' ,outdir='./send_folder' ):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(template)

    sheet = workbook.active
    # Specify the page ranges and column ranges to enter data into
    page_ranges = [
        (8, 25),
        (39, 56),
        (70, 87),
        (101, 118)
    ]

    counter = 1
    
    class_data = student_details['class_name'].split('-')
    class_name , class_distribution = class_data[1] , class_data[0]
    
    school_name , school_code = student_details['school_name_code'].split(' - ')[1] , [int(digit) for digit in  student_details['school_name_code'].split(' - ')[0] ]
    teacher_incharge_name = student_details['teacher_incharge_name']
    principle_name = student_details['principle_name']
    student_details = student_details['students_info']
                                            
    
    # Iterate over each page range and insert data from the dictionary list
    for start_row, end_row in page_ranges:
        for row_number, dataFrame in zip(range(start_row, end_row+1), student_details):
            if counter > len(student_details):
                break
            sheet.cell(row=row_number, column=1).value = counter
            sheet.cell(row=row_number, column=2).value = dataFrame['student_id']
            sheet.cell(row=row_number, column=3).value = dataFrame['identity_type']
            sheet.cell(row=row_number, column=4).value = dataFrame['first_name']
            sheet.cell(row=row_number, column=5).value = dataFrame['second_name']
            sheet.cell(row=row_number, column=6).value = dataFrame['third_name']
            sheet.cell(row=row_number, column=7).value = dataFrame['last_name']
            sheet.cell(row=row_number, column=8).value = dataFrame['birthPlace_area']
            sheet.cell(row=row_number, column=9).value = dataFrame['birth_date'].replace('-','/')
            sheet.cell(row=row_number, column=10).value = dataFrame['nationality']
            sheet.cell(row=row_number, column=11).value = dataFrame['sex']
            sheet.cell(row=row_number, column=12).value = dataFrame['resident_governorate']
            sheet.cell(row=row_number, column=13).value = dataFrame['resident_district']
            sheet.cell(row=row_number, column=14).value = dataFrame['resident_quarter']
            sheet.cell(row=row_number, column=15).value = ''                                # dataFrame['identity_type']
            sheet.cell(row=row_number, column=16).value = ''
            sheet.cell(row=row_number, column=17).value = counter                                 # dataFrame['student_id']
            sheet.cell(row=row_number, column=18).value = dataFrame['student_id']
            sheet.cell(row=row_number, column=19).value = dataFrame['first_name']+' '+dataFrame['last_name']
            sheet.cell(row=row_number, column=20).value = dataFrame['marital_status']
            sheet.cell(row=row_number, column=21).value = dataFrame['mother_name']
            sheet.cell(row=row_number, column=22).value = dataFrame['study_type']
            sheet.cell(row=row_number, column=23).value = dataFrame['father_education_level']
            sheet.cell(row=row_number, column=24).value = dataFrame['mother_education_level']
            sheet.cell(row=row_number, column=25).value = dataFrame['guardian_name']
            sheet.cell(row=row_number, column=26).value = dataFrame['guardian_student_relationship']
            sheet.cell(row=row_number, column=27).value = dataFrame['guardian_employment']
            sheet.cell(row=row_number, column=28).value = dataFrame['family_size']
            sheet.cell(row=row_number, column=29).value = dataFrame['student_siblings_rank']
            sheet.cell(row=row_number, column=30).value = dataFrame['student_health_status']
            sheet.cell(row=row_number, column=31).value = dataFrame['student_academic_status']
            sheet.cell(row=row_number, column=32).value = dataFrame['external_aid_available']
            sheet.cell(row=row_number, column=33).value = dataFrame['monthly_family_income']
            sheet.cell(row=row_number, column=34).value = dataFrame['religion']
            sheet.cell(row=row_number, column=35).value = dataFrame['govt_card_attribute']
            sheet.cell(row=row_number, column=35).value = dataFrame['guardian_phone_number']
    
    sheet.cell(row=1, column=39).value = school_name
    sheet.cell(row=3, column=39).value = class_distribution
    sheet.cell(row=4, column=39).value = class_name.replace('الصف', '')
    sheet.cell(row=5, column=39).value = teacher_incharge_name
    sheet.cell(row=6, column=39).value = principle_name
    
    for column_idx , digit in enumerate(school_code,start=39):
        sheet.cell(row=2, column=column_idx).value = digit

        
        counter += 1
    
    # Save the workbook
    workbook.save(outdir+'/your_file.xlsx')

def get_student_statistic_info(username,password, identity_nos=None, students_openemis_nos=None, student_ids=None, session=None , teacher_full_name=False ):
    auth = get_auth(username,password)
    final_dict_info = []
    identity_types = get_IdentityTypes(auth, session=session)
    area_data = get_AreaAdministrativeLevels(auth, session=session)['data']
    nationality_data = {i['id']: i['name'] for i in make_request(auth=auth, url='https://emis.moe.gov.jo/openemis-core/restful/v2/User-NationalityNames')['data']}
    curr_period_data = get_curr_period(auth,session=session)
    curr_period = curr_period_data['data'][0]['id']
    inst_id = inst_name(auth,session=session)['data'][0]['Institutions']['id']
    class_data_url = f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-InstitutionClasses?_contain=Institutions&_fields=id,name,institution_id,academic_period_id,Institutions.code,Institutions.name,Institutions.area_administrative_id&_finder=classesByInstitutionAndAcademicPeriod[institution_id:{inst_id};academic_period_id:{curr_period}]'
    class_data = make_request(auth=auth , url=class_data_url ,session=session)['data'][0]
    academic_period_id, institution_id , institution_name_code ,modeeriah= class_data['academic_period_id'],class_data['institution_id'],class_data['institution']['code_name'], [i['name'] for i in area_data if i['id'] == class_data['institution']['area_administrative_id']][0]
    class_name ,teacher_incharge_name= '' , ['']*4
    
    # بيانات المدرسة للشهادات الملونة
    school_data = inst_name(auth,session=session)['data'][0]
    inst_id = school_data['Institutions']['id']
    school_info = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={inst_id}&_contain=InstitutionLands.CustomFieldValues,AreaAdministratives,Areas',session=session)['data'][0]
    
    school_address = school_info['address']
    school_phone_number = school_info['telephone']
    school_national_id = school_info['code']
    school_directorate = ' مديرية لواء '+school_info['area']['name']
    school_bridge = ' لواء '+school_info['area']['name']

    
    staff = get_school_teachers(auth,id=institution_id ,session=session)['staff'] 

    working_teachers = [teacher 
                        for teacher in staff 
                            if teacher['staff_status'] == 1
                        ]
    
    principle_name = [
                        i['name_list']
                        for i in working_teachers 
                        if '- مدير' in i['position']
                    ][0]

    
    if identity_nos is not None:
        for chunk in chunks(identity_nos, 20):
            joined_string = ','.join([f'identity_number:{i}' for i in chunk])
            url='https://emis.moe.gov.jo/openemis-core/restful/Institution-StudentUser?_limit=0&_contain=BirthplaceAreas,CustomFieldValues,Identities&_orWhere='+joined_string
            students_info_data = make_request(auth=auth , url=url,session=session)['data']
            final_dict_info.extend(process_students_info(students_info_data, identity_types, nationality_data , area_data))

    elif students_openemis_nos is not None:
        for chunk in chunks(students_openemis_nos, 20):
            joined_string = ','.join([f'openemis_no:{i}' for i in chunk])
            url='https://emis.moe.gov.jo/openemis-core/restful/Institution-StudentUser?_limit=0&_contain=BirthplaceAreas,CustomFieldValues,Identities&_orWhere='+joined_string
            students_info_data = make_request(auth=auth , url=url,session=session)['data']
            final_dict_info.extend(process_students_info(students_info_data, identity_types, nationality_data , area_data))            

    elif student_ids is not None:
        for chunk in chunks(student_ids, 20):
            joined_string = ','.join([f'id:{i}' for i in chunk])
            url='https://emis.moe.gov.jo/openemis-core/restful/Institution-StudentUser?_limit=0&_contain=BirthplaceAreas,CustomFieldValues,Identities&_orWhere='+joined_string
            students_info_data = make_request(auth=auth , url=url,session=session)['data']
            final_dict_info.extend(process_students_info(students_info_data, identity_types, nationality_data , area_data))

    else : 
        # احضر بيانات الصف الي مع المعلم
        institution_class_id ,class_name = class_data['id'] , class_data['name']
        
        try : 
            teacher_incharge_name = [
                            i['name_list'] 
                            for i in working_teachers 
                                if i['nat_id'] == str(username) or i['default_nat_id'] == str(username) 
                            ][0]
        except :
            t =make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={inst_id}&_contain=Staff.Users,Staff.Positions' )
            
            teacher_incharge_name = [
                [i['user']['first_name'],
                 i['user']['middle_name'],
                 i['user']['third_name'],
                 i['user']['last_name']] for i in t['data'][0]['staff'] 
                                                if '9892006970' in i['staff_name']
                                                ][0]

        # # احضر اسماء الطلاب في الصف
        url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id&_limit=0&academic_period_id={academic_period_id}&institution_class_id={institution_class_id}&institution_id={institution_id}"
        student_ids= make_request(url,auth,session=session)
        student_ids =list(set([i['student_id'] for i in student_ids['data'] ]))
        for chunk in chunks(student_ids, 20):
            joined_string = ','.join(str(i) for i in [f'id:{i}' for i in chunk])
            url='https://emis.moe.gov.jo/openemis-core/restful/Institution-StudentUser?_limit=0&_contain=BirthplaceAreas,CustomFieldValues,Identities&_orWhere='+joined_string
            students_info_data = make_request(auth=auth , url=url,session=session)['data']
            final_dict_info.extend(process_students_info(students_info_data, identity_types, nationality_data , area_data))
           
    sorted_final_dict_info=sorted(final_dict_info, key=lambda x: x['full_name'])
    
    # c['code'] ====> '2022-2023'
    # c['code'].split('-')[0]  ====> '2022'
    # c['code'].split('-')[1]  ====> '2023'

    year_code = curr_period_data['data'][0]['code']

    start_of_the_year_date = curr_period_data['data'][0]['start_date']
    end_of_the_year_date = curr_period_data['data'][0]['end_date']
    return {'students_info':sorted_final_dict_info ,
            'class_name':class_name ,
            'school_name_code':institution_name_code ,
            'modeeriah' : modeeriah,
            'principle_name': principle_name[0]+' '+principle_name[1]+' '+principle_name[3],
            'teacher_incharge_name': teacher_incharge_name[0]+' '+teacher_incharge_name[1]+' '+teacher_incharge_name[3] if not teacher_full_name else ' '.join(teacher_incharge_name),
            'year_code': year_code,
            'school_address' :school_address ,
            'school_phone_number' :school_phone_number ,
            'school_national_id' :school_national_id ,
            'school_directorate' :school_directorate ,
            'school_bridge' :school_bridge ,
            'academic_year_1':year_code.split('-')[0],
            'academic_year_2':year_code.split('-')[1],
            'start_date':start_of_the_year_date,
            'end_date': end_of_the_year_date
            }
    
def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def process_students_info(students_info_data, identity_types, nationality_data , area_data):
    dic_list=[]
    options_values_dic ={88: 'اعزب',89: 'متزوج',90: 'ارمل',91: 'مطلقة',124: 'نظامية',125: 'منزلية',80: 'امي',81: 'اساسي',82: 'ثانوي',83: 'كلية مجتمع',84: 'بكالوريوس',85: 'دلوم عالي'
                        ,86: 'ماجستير',87: 'دكتوراة',92: 'امي',93: 'اساسي',94: 'ثانوي',95: 'كلية مجتمع',96: 'بكالوريوس',97: 'دلوم عالي',98: 'ماجستير',99: 'دكتوراة',115: 'اب',116: 'ام'
                        ,117: 'نفسه',118: 'عم-عمه',119: 'جد-جدة',120: 'خال-خالة',121: 'اخ-اخت',136: 'اخرى',100: 'سليم',101: 'غير سليم',110: 'ناجح',111: 'معيد',112: 'متسرب'
                        ,122: 'لا يوجد',123: 'يوجد',144: 'نعم',145: 'لا',127: 'الاسلام',128: 'المسيحية',113: 'لا يحمل بطاقة',114: 'لاجئ',137: 'روضة 2 (تمهيدي)',138: 'روضة 1 (بستان)'
                        ,141: 'روضة 2 (تمهيدي) روضة 1 (بستان)',142: 'لم يلتحق',158: 'يرسم',159: 'الخط',160: 'الصوت الجميل',161: 'العزف',162: 'رياصية',164: 'التمثيل',165: 'الشعر'
                        ,166: 'الرواية',167: 'اخرى',168: 'التسريع الأكاديمي',169: 'مدارس الملك عبد الله الثاني للتميز',140: 'المراكز الريادية',171: 'غرف مصادر الطلبة الموهوبين',172: 'جائزة انتل'
                        ,173: 'جائزة روبوتكس',174: 'جائزة اخرى',175: 'اختراع',176: 'ابتكار',177: 'فكرة ابداعية',178: 'استكشاف مقصود', 179:'نعم' ,180 :'لا' ,1:"ذكر" ,2:"انثى" }
    variables = {'mother_name': 1,'guardian_employment': 5,'student_siblings_rank': 7,'family_size': 6,'address':'',
        'monthly_family_income': 9,'guardian_name': 11,'father_education_level': 12,'marital_status': 13,
        'mother_education_level': 14,'student_health_status': 15,'student_academic_status': 17,'govt_card_attribute': 18,
        'guardian_student_relationship': 19,'external_aid_available': 20,'study_type': 21,'religion': 22,'did_student_attend_kindergarten': 28,
        'is_family_registered_nationally': 30,'guardian_phone_number': 31,'mother_nationality': 32,'did_student_join_international_program': 37,
        'did_student_attend_kindergarten' : '','intelligence_giftedness' : '','talent_and_giftedness' : '','talent' : '',
        }
    
    for data_item in students_info_data:
        ''' 
        تفاصيل حقول البييانات الاحصائية للطالب custom_field_values keys
            1 ==> اسم الأم (الاسم الاول والعائلة)  variable: 'mother_name'
            5 ==> عمل ولي الأمر    variable: 'guardian_employment'
            6 ==> عدد أفراد الأسرة     variable: 'family_size'
            7 ==> ترتيب الطالب بين اخوانه  variable: 'student_siblings_rank'
            9 ==> دخل الأسرة الشهري    variable: 'monthly_family_income'
            11 ==> اسم ولي الأمر   variable: 'guardian_name'
            12 ==> مستوى تعليم الاب    variable: 'father_education_level'
            13 ==> الحالة الاجتماعية   variable: 'marital_status'
            14 ==> مستوى تعليم الام    variable: 'mother_education_level'
            15 ==> الوضع الصحي للطالب  variable: 'student_health_status'
            17 ==> الوضع الدراسي للطالب    variable: 'student_academic_status'
            18 ==> صفة بطاقة الغوث     variable: 'govt_card_attribute'
            19 ==> علاقة ولي الامر بالطالب     variable: 'guardian_student_relationship'
            20 ==> المساعدات الخارجية (إن وجدت)    variable: 'external_aid_available'
            21 ==> نوع الدراسة     variable: 'study_type'
            22 ==> الديانة     variable: 'religion'
            28 ==> هل التحق الطالب/الطالبة برياض الاطفال؟  variable: 'did_student_attend_kindergarten'
            30 ==> هل الاسرة مسجلة بالمعونة الوطنية؟   variable: 'is_family_registered_nationally'
            31 ==> هاتف ولي امر الطالب     variable: 'guardian_phone_number'
            32 ==> جنسية الام  variable: 'mother_nationality'
            37 ==> هل التحق الطالب بالبرنامج الدولي    variable: 'did_student_join_international_program'
                                    الموهبة    variable: 'talent'
                                التفوق و الموهبة   variable: 'excellence_and_talent'
                                التفوق العقلي     variable: 'intellectual_excellence'   
        '''

        custom_field_values = data_item['custom_field_values'] 
        custom_field_values_dict = {item['student_custom_field_id']: str(item[key]) for item in custom_field_values 
                                                                                        for key in ['text_value', 
                                                                                                    'number_value', 
                                                                                                    'decimal_value', 
                                                                                                    'textarea_value', 
                                                                                                    'date_value', 
                                                                                                    'time_value'] 
                                                                                            if item.get(key) is not None }

        result = {var_name: custom_field_values_dict.get(var_id, '') for var_name, var_id in variables.items()}
        result = {
            key: options_values_dic[int(val)] if val.isdigit() 
                and int(val) in options_values_dic  
                    and key not in ['family_size', 'monthly_family_income', 'student_siblings_rank', 'guardian_phone_number'] 
            else val
            for key, val in result.items()
        }

        area_chain = find_area_chain(data_item['address_area_id'], area_data).split(' - ')
        result['student_id'] = data_item['id']
        result['birthPlace_area'] = '' if data_item['birthplace_area'] is None else data_item['birthplace_area']['name'] 
        result['identity_type'] =  '' if data_item['identity_type_id'] != 825 else identity_types[data_item['identity_type_id']]
        result['identity_number'] = '' if len(data_item['identities']) ==0 else data_item['identities'][0]['number']
        result['full_name'] = '' if data_item['name'] is None else data_item['name'] 
        result['first_name'] = '' if data_item['first_name'] is None else data_item['first_name'] 
        result['second_name'] = '' if data_item['middle_name'] is None else data_item['middle_name'] 
        result['third_name'] = '' if data_item['third_name'] is None else data_item['third_name'] 
        result['last_name'] = '' if data_item['last_name'] is None else data_item['last_name'] 
        result['birth_date'] = '' if data_item['date_of_birth'] is None else data_item['date_of_birth'] 
        result['nationality'] = '' if data_item['nationality_id']  is None else nationality_data[data_item['nationality_id']]  
        result['sex'] = '' if data_item['gender_id']  is None else options_values_dic[data_item['gender_id']]  
        result['resident_governorate'] = area_chain[0] if len(area_chain) == 1 else ''
        result['resident_district'] = area_chain[1] if len(area_chain) == 2 else ''
        result['resident_quarter'] = area_chain[2] if len(area_chain) == 3 else ''    
        result['address'] = '' if data_item['address'] is None else data_item['address'] 
        dic_list.append(result)
    
    return dic_list

def find_parent_info(item_id ,area_data):
    for item in area_data:
        if item['id'] == item_id:
            parent_id = item['parent_id']
            name = item['name']
            if parent_id in [3, 4, 5 ,1]:
                return None, name
            return parent_id, name
    return None, None

def find_area_chain(id,area_data):
    names = []

    while id is not None:
        id, name = find_parent_info(id ,area_data)
        if name:
            names.append(name)
    
    names.reverse()  # Reverse the order of names            
    output = ' - '.join(names)
    return output

def get_AreaAdministrativeLevels(auth,session=None):
    url='https://emis.moe.gov.jo/openemis-core/restful/v2/Area-AreaAdministratives?_limit=0&_contain=AreaAdministrativeLevels&_fields=id,name,parent_id,area_administrative_level_id'
    return make_request(auth=auth , url= url,session=session)

def get_IdentityTypes(auth,session=None):
    url='https://emis.moe.gov.jo/openemis-core/restful/v2/FieldOption-IdentityTypes.json?_limit=0&_fields=id,name'
    return { i['id'] : i['name']  for i in make_request(auth=auth , url=url ,session=session)['data']}
    
def find_default_teachers_creds(auth ,id=None , nat_school=None ,session=None):
    if id == None:
        teachers = get_school_teachers(auth,nat_school=nat_school,session=session)['staff']
    else:
        teachers = get_school_teachers(auth,id=id,session=session)['staff']

    working_teachers = [(teacher['name'],teacher['nat_id']) for teacher in teachers if teacher['staff_status'] == 1]

    found_creds,unfound_creds = [],[]
    for teacher in working_teachers:
        if get_auth(teacher [1] , teacher [1]):
            found_creds.append(teacher[1])
            print('found password for this teacher ' + teacher[0]+' -----> '+teacher[1])
        else: 
            unfound_creds.append(str(teacher[0]+'-'+teacher[1]))
    return {'institution_staff': teachers , 'found_creds': found_creds ,'unfound_creds': unfound_creds}

def five_names_every_class_wrapper(auth , emp_number ,term=1 , session=None):
    data = five_names_every_class(auth , emp_number ,session=session)
    term = 'term1' if term == 1 else 'term2'
    long_text = ''

    for subject in data['row_data']:
        text =''
        middle_index = len(subject['marks_and_name']) // 2
        first_two = subject['marks_and_name'][:2]
        middle_one = subject['marks_and_name'][middle_index]
        last_two =data['row_data'][0]['marks_and_name'][-2:]
        for item_dic in first_two : 
            text +=  item_dic['name'] +'\n'+'\t'+ ' ت1 ---> ' + str(item_dic[term]['assessment1']) +'\n'+'\t'+ ' ت2 ---> ' + str(item_dic[term]['assessment2'])+'\n'+'\t'+' ت3 ---> ' +str(item_dic[term]['assessment3']) +'\n'+'\t'+'النهائي ---> ' +str(item_dic[term]['assessment4'])+'\n' 
        text += '[ .......... ]'+'\n'        
        text +=  middle_one['name'] +'\n'+'\t'+ ' ت1 ---> ' + str(middle_one[term]['assessment1']) +'\n'+'\t'+ ' ت2 ---> ' + str(middle_one[term]['assessment2'])+'\n'+'\t'+' ت3 ---> ' +str(middle_one[term]['assessment3']) +'\n'+'\t'+'النهائي ---> ' +str(middle_one[term]['assessment4']) +'\n' 
        text += '[ .......... ]'+'\n'            
        for item_dic in last_two : 
            text +=  item_dic['name'] +'\n'+'\t'+ ' ت1 ---> ' + str(item_dic[term]['assessment1']) +'\n'+'\t'+ ' ت2 ---> ' + str(item_dic[term]['assessment2'])+'\n'+'\t'+' ت3 ---> ' +str(item_dic[term]['assessment3']) +'\n'+'\t'+'النهائي ---> ' +str(item_dic[term]['assessment4'])+'\n' 
        long_text += '\n'+subject['subject']+'//'+subject['className']+'\n'+text + '-'*70
        
    return long_text

def five_names_every_class(auth, emp_username ,session=None ):
    period_id = get_curr_period(auth,session=session)['data'][0]['id']
    user = user_info(auth , emp_username,session=session)
    userInfo = user['data'][0]
    user_id , user_name = userInfo['id'] , userInfo['first_name']+' '+ userInfo['last_name']+'-' + str(emp_username)
    # years = get_curr_period(auth)
    school_data = inst_name(auth,session=session)['data'][0]
    inst_id = school_data['Institutions']['id']
    # school_name = school_data['Institutions']['name']
    # grades = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Education.EducationGrades?_limit=0')
    # school_year = get_curr_period(auth)['data']

    
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id,session=session)['data']]
    classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i],session=session)['data'] for i in range(len(classes_id_1))]
    assessments = ['assessment1','assessment2','assessment3','assessment4']
    terms = ['term1','term2']
    upload_percentage,modified_classes,classes_id_3 ,classes,mawad,row_data =[],[],[],[],[],[]
    row_d={}

    for class_info in classes_id_2:
        classes_id_3.append([{"institution_class_id": class_info[0]['institution_class_id'] ,"sub_name": class_info[0]['institution_subject']['name'],"class_name": class_info[0]['institution_class']['name'] , 'subject_id': class_info[0]['institution_subject']['education_subject_id']}])

    for v in range(len(classes_id_1)):
        # id
        # print (classes_id_3[v][0]['institution_class_id'])
        # id = classes_id_3[v][0]['institution_class_id']
        # subject name 
        # print (classes_id_3[v][0]['sub_name'])
        # class name
        # print (classes_id_3[v][0]['class_name'])
        # class_name = classes_id_3[v][0]['class_name']
        # subject id 
        # print (classes_id_3[v][0]['subject_id'])

        mawad.append(classes_id_3[v][0]['sub_name'])
        classes.append(classes_id_3[v][0]['class_name'])
        class_name = classes_id_3[v][0]['class_name'].split('-')[0].replace('الصف ' , '')
        # class_char = classes_id_3[v][0]['class_name'].split('-')[1]
        # sub_name = classes_id_3[v][0]['sub_name']    
        
        students = get_class_students(auth
                                    ,period_id
                                    ,classes_id_1[v]
                                    ,classes_id_3[v][0]['institution_class_id']
                                    ,inst_id
                                    ,session=session)
        students_names = sorted([i['user']['name'] for i in students['data']])
        # print(students_names)
        students_id_and_names = []
        
        for IdAndName in students['data']:
            students_id_and_names.append({'student_name': IdAndName['user']['name'] , 'student_id':IdAndName['student_id']})

        assessments_json = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id='+str(classes_id_3[v][0]['subject_id'])+'&institution_classes_id='+ str(classes_id_3[v][0]['institution_class_id'])+ f'&institution_id={inst_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects')

        marks_and_name = []

        dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''}}
        for student_data_item in students_id_and_names:   
            for student_assessment_item in assessments_json['data']:
                if student_assessment_item['student_id'] == student_data_item['student_id'] :  
                    # FIXME: غير الشرط اذا كان None استبدل القيمة بلا شيء                    
                    if student_assessment_item["marks"] is not None :
                        dic['id'] = student_data_item['student_id'] 
                        dic['name'] = student_data_item['student_name'] 
                        if student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment1'] = student_assessment_item["marks"] 
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment2']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment3']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment4']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment1']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment2']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment3']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment4']  = student_assessment_item["marks"]
            marks_and_name.append(dic)
            dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} }


        marks_and_name = [d for d in marks_and_name if d['name'] != '']
        marks_and_name = sorted(marks_and_name, key=lambda x: x['name'])
        
        percent_dict ={'subject': '' , 'className' :'', 'term1' : {'assessment1_percentage': '', 'assessment2_percentage': '', 'assessment3_percentage': '', 'assessment4_percentage': ''} ,
        'term2':{'assessment1_percentage': '', 'assessment2_percentage': '', 'assessment3_percentage': '', 'assessment4_percentage': ''}}
        row_d={}        
        
        if 'عشر' in class_name :
            percent_dict ={'subject': '' , 'className' :'', 'term1' : {'assessment1_percentage': 0, 'assessment2_percentage': 0, 'assessment3_percentage': 0, 'assessment4_percentage': 0} ,
                            'term2':{'assessment1_percentage': 0, 'assessment2_percentage': 0, 'assessment3_percentage': 0, 'assessment4_percentage': 0}}
            
            percent_dict['subject']= classes_id_3[v][0]['sub_name']
            percent_dict['className']= classes_id_3[v][0]['class_name']
            row_d['subject'] = classes_id_3[v][0]['sub_name']
            row_d['className'] = classes_id_3[v][0]['class_name']
            # row_d['marks_and_name'] = marks_and_name
        else:
            for term in terms : 
                for assessment in assessments :
                    total_marks ,marks_uploaded= len([i[term][assessment] for i in marks_and_name ]) , len([i[term][assessment] for i in marks_and_name if i[term][assessment] != ''])
                    percentage = int((marks_uploaded / total_marks) * 100)
                    percent_dict[term][assessment+'_percentage']=percentage
                    
            percent_dict['subject']= classes_id_3[v][0]['sub_name']
            percent_dict['className']= classes_id_3[v][0]['class_name']
            row_d['subject'] = classes_id_3[v][0]['sub_name']
            row_d['className'] = classes_id_3[v][0]['class_name']
            row_d['marks_and_name'] = marks_and_name
        
        row_data.append(row_d)
        upload_percentage.append(percent_dict)
    
    return {'teacher': userInfo['name'] ,'upload_percentage' :upload_percentage , 'row_data':row_data}

def convert_official_marks_doc(ods_name='send', outdir='./send_folder' ,ods_num=1,file_path=None, file_content=None):
    ods_file = f'{ods_name}{ods_num}.ods'
    
    if file_content is None:
        doc = ezodf.opendoc(file_path)
    else:
        # Save the file content to a temporary file
        with tempfile.NamedTemporaryFile(delete=True) as temp_file:
            temp_file.write(file_content.getvalue())
            temp_file.flush()
            doc = ezodf.opendoc(temp_file.name)
            
            shutil.copy(temp_file.name, outdir+'/final_'+ods_file)
            
    os.system(f'soffice --headless --convert-to pdf:writer_pdf_Export --outdir {outdir}  {outdir}/final_{ods_file} ')
    add_margins(f"{outdir}/final_{ods_name}{ods_num}.pdf", f"{outdir}/output_file.pdf",top_rec=30, bottom_rec=50, left_rec=68, right_rec=120)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/سجل العلامات الرسمي.pdf",page=1 , top_rec=60, bottom_rec=80, left_rec=70, right_rec=120)
    split_A3_pages(f"{outdir}/output_file.pdf" , outdir)
    reorder_official_marks_to_A4(f"{outdir}/output.pdf" , f"{outdir}/reordered.pdf")

    add_margins(f"{outdir}/reordered.pdf", f"{outdir}/output_file.pdf",top_rec=60, bottom_rec=50, left_rec=68, right_rec=20)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/output_file1.pdf",page=1 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120)
    add_margins(f"{outdir}/output_file1.pdf", f"{outdir}/output_file2.pdf",page=50 , top_rec=100, bottom_rec=80, left_rec=70, right_rec=60)    
    add_margins(f"{outdir}/output_file2.pdf", f"{outdir}/سجل العلامات الرسمي_A4.pdf",page=51 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120)  
    delete_files_except([f"سجل العلامات الرسمي.pdf",f"سجل العلامات الرسمي_A4.pdf"], outdir)
    
def check_file_if_official_marks_file(file_path=None, file_content=None):
    if file_content is None:
        doc = ezodf.opendoc(file_path)
    else:
        # Save the file content to a temporary file
        with tempfile.NamedTemporaryFile(delete=True) as temp_file:
            temp_file.write(file_content.getvalue())
            temp_file.flush()
            doc = ezodf.opendoc(temp_file.name)

    exists = 'sheet' in [sheet.name for sheet in doc.sheets]
    return exists

def teachers_marks_upload_percentage_wrapper(auth ,term=1 ,inst_id=None , inst_nat=None , session=None , template='./templet_files/كشف نسبة الادخال معدل.xlsx' ,outdir='./send_folder/' ):
    # percent_dict ={'subject': '' , 'className' :'', 'term1' : {'assessment1_percentage': 0, 'assessment2_percentage': 0, 'assessment3_percentage': 0, 'assessment4_percentage': 0} ,
    #                             'term2':{'assessment1_percentage': 0, 'assessment2_percentage': 0, 'assessment3_percentage': 0, 'assessment4_percentage': 0}}
    # assessments = ['assessment1','assessment2','assessment3','assessment4']
    # terms = ['term1','term2']
    # terms = 'term2'
    # session = requests.Session()
    term = 'term1' if term == 1 else 'term2'
    if inst_id is None and inst_nat is None : 
        inst_id = inst_name(auth ,session=session)['data'][0]['Institutions']['id']
        
    teachers_nats = [teacher['nat_id'] for teacher in get_school_teachers( auth , id=inst_id ,nat_school=inst_nat , session=session)['staff'] if 'معلم' in teacher['position']]
    all_teachers_data  = []
    for nat in teachers_nats :
        all_teachers_data.append(teachers_marks_upload_percentage(auth , nat ,session=session))

    # load the existing workbook
    existing_wb = load_workbook(template)
    # Select the worksheet
    existing_ws = existing_wb.active

    
    row_number=11

    for teacher in all_teachers_data : 
        row = teacher['row_data']
        flattened_list = []
        marks_and_name = [r['marks_and_name'] for r in row if "عشر" not in r['className']]
        
        for group in marks_and_name:
            flattened_list.extend(group)

        total_marks = len([i[term]['assessment1'] for i in flattened_list ]) 
        if total_marks > 0:

            existing_ws.cell(row=row_number, column=1).value = teacher['teacher']
            existing_ws.cell(row=row_number, column=3).value = total_marks
            
            first_marks_uploaded = len([i[term]['assessment1'] for i in flattened_list if i[term]['assessment1'] != '']) 
            existing_ws.cell(row=row_number, column=4).value = first_marks_uploaded
            
            second_marks_uploaded =  len([i[term]['assessment2'] for i in flattened_list if i[term]['assessment2'] != ''])
            existing_ws.cell(row=row_number, column=5).value = second_marks_uploaded
            
            third_marks_uploaded =  len([i[term]['assessment3'] for i in flattened_list if i[term]['assessment3'] != ''])
            existing_ws.cell(row=row_number, column=6).value = third_marks_uploaded
            
            fourth_marks_uploaded =  len([i[term]['assessment4'] for i in flattened_list if i[term]['assessment4'] != ''])
            existing_ws.cell(row=row_number, column=7).value = fourth_marks_uploaded
            
            try :
                percentage = int((first_marks_uploaded + second_marks_uploaded + third_marks_uploaded + fourth_marks_uploaded) / (total_marks * 4) * 100)
            except : 
                percentage = 0
                
            existing_ws.cell(row=row_number, column=2).value = percentage          
            row_number+=1
        else:
            row_number-=1
    existing_wb.save( outdir + f'output.xlsx')
    
    playsound()
        
def teachers_marks_upload_percentage(auth, emp_username, template='./templet_files/side_marks_note_2.docx' ,outdir='./send_folder/' ,first_page='side_mark_first_page.docx', template_dir='./templet_files/',term=1 ,session=None ):
    period_id = get_curr_period(auth,session=session)['data'][0]['id']
    user = user_info(auth , emp_username,session=session)
    userInfo = user['data'][0]
    user_id , user_name = userInfo['id'] , userInfo['first_name']+' '+ userInfo['last_name']+'-' + str(emp_username)
    # years = get_curr_period(auth)
    school_data = inst_name(auth,session=session)['data'][0]
    inst_id = school_data['Institutions']['id']
    # school_name = school_data['Institutions']['name']
    # grades = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Education.EducationGrades?_limit=0')
    # school_year = get_curr_period(auth)['data']

    
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id,session=session)['data']]
    classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i],session=session)['data'] for i in range(len(classes_id_1))]
    assessments = ['assessment1','assessment2','assessment3','assessment4']
    terms = ['term1','term2']
    upload_percentage,modified_classes,classes_id_3 ,classes,mawad,row_data =[],[],[],[],[],[]
    row_d={}

    for class_info in classes_id_2:
        classes_id_3.append([{"institution_class_id": class_info[0]['institution_class_id'] ,"sub_name": class_info[0]['institution_subject']['name'],"class_name": class_info[0]['institution_class']['name'] , 'subject_id': class_info[0]['institution_subject']['education_subject_id']}])

    for v in range(len(classes_id_1)):
        # id
        # print (classes_id_3[v][0]['institution_class_id'])
        # id = classes_id_3[v][0]['institution_class_id']
        # subject name 
        # print (classes_id_3[v][0]['sub_name'])
        # class name
        # print (classes_id_3[v][0]['class_name'])
        # class_name = classes_id_3[v][0]['class_name']
        # subject id 
        # print (classes_id_3[v][0]['subject_id'])

        mawad.append(classes_id_3[v][0]['sub_name'])
        classes.append(classes_id_3[v][0]['class_name'])
        class_name = classes_id_3[v][0]['class_name'].split('-')[0].replace('الصف ' , '')
        class_char = classes_id_3[v][0]['class_name'].split('-')[1]
        # sub_name = classes_id_3[v][0]['sub_name']    
        
        students = get_class_students(auth
                                    ,period_id
                                    ,classes_id_1[v]
                                    ,classes_id_3[v][0]['institution_class_id']
                                    ,inst_id
                                    ,session=session)
        students_names = sorted([i['user']['name'] for i in students['data']])
        # print(students_names)
        students_id_and_names = []
        
        for IdAndName in students['data']:
            students_id_and_names.append({'student_name': IdAndName['user']['name'] , 'student_id':IdAndName['student_id']})

        assessments_json = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id='+str(classes_id_3[v][0]['subject_id'])+'&institution_classes_id='+ str(classes_id_3[v][0]['institution_class_id'])+ f'&institution_id={inst_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects',session=session)

        marks_and_name = []

        dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''}}
        for student_data_item in students_id_and_names:   
            for student_assessment_item in assessments_json['data']:
                if student_assessment_item['student_id'] == student_data_item['student_id'] :  
                    # FIXME: غير الشرط اذا كان None استبدل القيمة بلا شيء                    
                    if student_assessment_item["marks"] is not None :
                        dic['id'] = student_data_item['student_id'] 
                        dic['name'] = student_data_item['student_name'] 
                        if student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment1'] = student_assessment_item["marks"] 
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment2']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment3']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment4']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment1']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment2']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment3']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment4']  = student_assessment_item["marks"]
            marks_and_name.append(dic)
            dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} }


        marks_and_name = [d for d in marks_and_name if d['name'] != '']
        marks_and_name = sorted(marks_and_name, key=lambda x: x['name'])
        
        percent_dict ={'subject': '' , 'className' :'', 'term1' : {'assessment1_percentage': '', 'assessment2_percentage': '', 'assessment3_percentage': '', 'assessment4_percentage': ''} ,
        'term2':{'assessment1_percentage': '', 'assessment2_percentage': '', 'assessment3_percentage': '', 'assessment4_percentage': ''}}
        row_d={}        
        
        if 'عشر' in class_name :
            percent_dict ={'subject': '' , 'className' :'', 'term1' : {'assessment1_percentage': 0, 'assessment2_percentage': 0, 'assessment3_percentage': 0, 'assessment4_percentage': 0} ,
                            'term2':{'assessment1_percentage': 0, 'assessment2_percentage': 0, 'assessment3_percentage': 0, 'assessment4_percentage': 0}}
            
            percent_dict['subject']= classes_id_3[v][0]['sub_name']
            percent_dict['className']= classes_id_3[v][0]['class_name']
            row_d['subject'] = classes_id_3[v][0]['sub_name']
            row_d['className'] = classes_id_3[v][0]['class_name']
            # row_d['marks_and_name'] = marks_and_name
        else:
            for term in terms : 
                for assessment in assessments :
                    total_marks ,marks_uploaded= len([i[term][assessment] for i in marks_and_name ]) , len([i[term][assessment] for i in marks_and_name if i[term][assessment] != ''])
                    percentage = int((marks_uploaded / total_marks) * 100)
                    percent_dict[term][assessment+'_percentage']=percentage
                    
            percent_dict['subject']= classes_id_3[v][0]['sub_name']
            percent_dict['className']= classes_id_3[v][0]['class_name']
            row_d['subject'] = classes_id_3[v][0]['sub_name']
            row_d['className'] = classes_id_3[v][0]['class_name']
            row_d['marks_and_name'] = marks_and_name
        
        row_data.append(row_d)
        upload_percentage.append(percent_dict)
    
    return {'teacher': userInfo['name'] ,'upload_percentage' :upload_percentage , 'row_data':row_data}

def side_marks_document_with_marks(username=None , password=None ,classes_data=None,template='./templet_files/side_marks_note_2.docx' ,outdir='./send_folder/' ,first_page='side_mark_first_page.docx', template_dir='./templet_files/',term=1 , names_only=False, session=None):
    """دالة تقوم بانشاء سجل علامات جانبي وتعبئة العلامات 

    Args:
        username (_type_, optional): اسم المستخدم. Defaults to None.
        password (_type_, optional): كلم السر. Defaults to None.
        classes_data (_type_, optional): بيانات الصفوف اذا استعملت الدالة اوفلاين. Defaults to None.
        template (str, optional): النموذج الذي يتم استخدامه. Defaults to './templet_files/side_marks_note_2.docx'.
        outdir (str, optional): المجلد الذي يتم الحفظ فيه . Defaults to './send_folder/'.
        first_page (str, optional): الصفحة الاولى للسجل الجانبي . Defaults to 'side_mark_first_page.docx'.
        template_dir (str, optional): مجلد النماذج. Defaults to './templet_files/'.
        term (int, optional): الفصل اما الاول او الثاني. Defaults to 1.
        names_only (bool, optional): خيار لطباعة الاسماء فقط. Defaults to False.
    """    
    
    classes=[]
    mawad=[]
    modified_classes=[]
    context={}
    
    if username is not None and password is not None : 
        auth = get_auth(username , password)
        period_id = get_curr_period(auth)['data'][0]['id']
        user = user_info(auth , username)
        userInfo = user['data'][0]
        user_id , user_name = userInfo['id'] , userInfo['first_name']+' '+ userInfo['last_name']+'-' + str(username)
        # years = get_curr_period(auth)
        school_data = inst_name(auth)['data'][0]
        inst_id = school_data['Institutions']['id']
        school_name = school_data['Institutions']['name']
        school_name_id = f'{school_name}={inst_id}'
        baldah = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={inst_id}&_contain=InstitutionLands.CustomFieldValues')['data'][0]['address'].split('-')[0]
        # grades = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Education.EducationGrades?_limit=0')
        modeeriah = inst_area(auth)['data'][0]['Areas']['name']
        school_year = get_curr_period(auth)['data']
        melady1 = str(school_year[0]['start_year'])
        melady2 = str(school_year[0]['end_year'])
        teacher = user['data'][0]['name'].split(' ')[0]+' '+user['data'][0]['name'].split(' ')[-1]
        
        # ما بعرف كيف سويتها لكن زبطت 
        
        classes_id_1 = sorted([[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id,session=session)['data']])
        classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i])['data'] for i in range(len(classes_id_1))]
        classes_id_3 = []  

        for class_info in classes_id_2:
            classes_id_3.append([{"institution_class_id": class_info[0]['institution_class_id'] ,"sub_name": class_info[0]['institution_subject']['name'],"class_name": class_info[0]['institution_class']['name'] , 'subject_id': class_info[0]['institution_subject']['education_subject_id']}])

        for v in range(len(classes_id_1)):
            # id
            print (classes_id_3[v][0]['institution_class_id'])
            id = classes_id_3[v][0]['institution_class_id']
            # subject name 
            print (classes_id_3[v][0]['sub_name'])
            # class name
            print (classes_id_3[v][0]['class_name'])
            # class_name = classes_id_3[v][0]['class_name']
            # subject id 
            print (classes_id_3[v][0]['subject_id'])

            mawad.append(classes_id_3[v][0]['sub_name'])
            classes.append(classes_id_3[v][0]['class_name'])
            class_name = classes_id_3[v][0]['class_name'].split('-')[0].replace('الصف ' , '')
            class_char = classes_id_3[v][0]['class_name'].split('-')[1]
            # sub_name = classes_id_3[v][0]['sub_name']    
            
            students = get_class_students(auth
                                        ,period_id
                                        ,classes_id_1[v]
                                        ,classes_id_3[v][0]['institution_class_id']
                                        ,inst_id)
            students_names = sorted([i['user']['name'] for i in students['data']])
            print(students_names)
            students_id_and_names = []
            
            for IdAndName in students['data']:
                students_id_and_names.append({'student_name': IdAndName['user']['name'] , 'student_id':IdAndName['student_id'] , 'term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''}})

            assessments_json = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id='+str(classes_id_3[v][0]['subject_id'])+'&institution_classes_id='+ str(classes_id_3[v][0]['institution_class_id'])+ f'&institution_id={inst_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects')

            marks_and_name = []

            dic = {'id':'' ,'student_name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''}}
            for student_data_item in students_id_and_names:
                for student_assessment_item in assessments_json['data']:
                    if student_assessment_item['student_id'] == student_data_item['student_id'] :  
                    # FIXME: غير الشرط اذا كان None استبدل القيمة بلا شيء                        
                        if student_assessment_item["marks"] is not None :
                            dic['id'] = student_data_item['student_id'] 
                            dic['student_name'] = student_data_item['student_name'] 
                            if student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                                dic['term1']['assessment1'] = student_assessment_item["marks"] 
                            elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                                dic['term1']['assessment2']  = student_assessment_item["marks"]
                            elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                                dic['term1']['assessment3']  = student_assessment_item["marks"]
                            elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                                dic['term1']['assessment4']  = student_assessment_item["marks"]
                            elif student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                                dic['term2']['assessment1']  = student_assessment_item["marks"]
                            elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                                dic['term2']['assessment2']  = student_assessment_item["marks"]
                            elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                                dic['term2']['assessment3']  = student_assessment_item["marks"]
                            elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                                dic['term2']['assessment4']  = student_assessment_item["marks"]
                marks_and_name.append(dic)
                dic = {'id':'' ,'student_name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} }

            marks_and_name = [d for d in marks_and_name if d['student_name'] != '']
            marks_and_name = sorted(students_id_and_names, key=lambda x: x['student_name'])  if  assessments_json  else sorted(marks_and_name, key=lambda x: x['name'])
            if 'عشر' in class_name :
                counter = 0
                for item in marks_and_name :
                    context[f'name{counter}'] = item['student_name']
                    counter+=1 
            else:
                counter = 0
                for item in marks_and_name :
                    context[f'name{counter}'] = item['student_name']
                    if not names_only :
                        assessments = [
                                    item[f'term{term}']['assessment1'],
                                    item[f'term{term}']['assessment2'],
                                    item[f'term{term}']['assessment3'],
                                    item[f'term{term}']['assessment4']
                                    ]
                        context[f'A1_{counter}'] = item[f'term{term}']['assessment1']
                        context[f'A2_{counter}'] = item[f'term{term}']['assessment2']
                        context[f'A3_{counter}'] = item[f'term{term}']['assessment3']
                        context[f'A4_{counter}'] = item[f'term{term}']['assessment4']
                        SUM = sum(int(assessment) if assessment != '' else 0 for assessment in assessments)                    
                        context[f'S_{counter}'] = SUM if SUM !=0 else ''
                        total = item[f'term{term}']['assessment3']

                        try :                    
                            variables = [random.randint(3, min(total, 5)) for _ in range(3) if total > 0]
                            variables.append(total - sum(variables))
                            context[f'M1_{counter}'] ,context[f'M2_{counter}'] ,context[f'M3_{counter}'] ,context[f'M4_{counter}'] = variables
                        except : 
                            context[f'M1_{counter}'] ,context[f'M2_{counter}'] ,context[f'M3_{counter}'] ,context[f'M4_{counter}'] =['']*4
                    counter+=1 
            context['teacher'] = userInfo['first_name']+' '+ userInfo['middle_name'] +' '+ userInfo['last_name']
            context[f'class_name'] = class_name+' '+ class_char
            context[f'term'] = 'الأول' if term == 1 else 'الثاني'
            context['school'] = school_name
            context['directory'] = modeeriah
            context['y1'] = melady1
            context['y2'] = melady2
            context['sub'] = classes_id_3[v][0]['sub_name']
            fill_doc(template , context , outdir+f'send{v}.docx' )
            context.clear()
            generate_pdf(outdir+f'send{v}.docx' , outdir ,v)
            delete_pdf_page(outdir+f'send{v}.pdf', outdir+f'SEND{v}.pdf', 1)
            delete_file(outdir+f'send{v}.pdf')

    else:
        student_details = classes_data
        school_name = student_details['custom_shapes']['school']
        modified_classes =student_details['custom_shapes']['classes']
        teacher = student_details['custom_shapes']['teacher'] 
        melady2 = student_details['custom_shapes']['melady1']
        melady1 = student_details['custom_shapes']['melady2']
        modeeriah = student_details['custom_shapes']['modeeriah']
        
        # Iterate over student data
        for v ,students_data_list in enumerate(student_details['file_data']):
            class_name = students_data_list['class_name'].split('=')[0].replace('الصف', '')
            mawad.append(students_data_list['class_name'].split('=')[1])
            classes.append(class_name)

            if 'عشر' in class_name :
                counter = 0
                for item in students_data_list['students_data'] :
                    context[f'name{counter}'] = item['name']
            else:
                counter = 0
                for item in students_data_list['students_data'] :
                    context[f'name{counter}'] = item['name']
                    if not names_only :
                        assessments = [
                                    item[f'term{term}']['assessment1'],
                                    item[f'term{term}']['assessment2'],
                                    item[f'term{term}']['assessment3'],
                                    item[f'term{term}']['assessment4']
                                    ]
                        context[f'A1_{counter}'] = item[f'term{term}']['assessment1']
                        context[f'A2_{counter}'] = item[f'term{term}']['assessment2']
                        context[f'A3_{counter}'] = item[f'term{term}']['assessment3']
                        context[f'A4_{counter}'] = item[f'term{term}']['assessment4']
                        SUM = sum(int(assessment) if assessment != '' else 0 for assessment in assessments)                    
                        context[f'S_{counter}'] = SUM if SUM !=0 else ''
                        total = item[f'term{term}']['assessment3']

                        try :                    
                            variables = [random.randint(3, min(total, 5)) for _ in range(3) if total > 0]
                            variables.append(total - sum(variables))       
                            context[f'M1_{counter}'] ,context[f'M2_{counter}'] ,context[f'M3_{counter}'] ,context[f'M4_{counter}'] = variables
                        except : 
                            context[f'M1_{counter}'] ,context[f'M2_{counter}'] ,context[f'M3_{counter}'] ,context[f'M4_{counter}'] =['']*4                        
                        
                    counter+=1 
                    context['teacher'] = teacher
            context[f'class_name'] = class_name
            context[f'term'] = 'الأول' if term == 1 else 'الثاني'
            context['school'] = school_name
            context['directory'] = modeeriah
            context['y1'] = melady1
            context['y2'] = melady2
            context['sub'] = students_data_list['class_name'].split('=')[1]
            fill_doc(template , context , outdir+f'send{v}.docx' )
            context.clear()
            generate_pdf(outdir+f'send{v}.docx' , outdir ,v)
            delete_pdf_page(outdir+f'send{v}.pdf', outdir+f'SEND{v}.pdf', 1)
            delete_file(outdir+f'send{v}.pdf')    

    for i in classes: 
        modified_classes.append(mawad_representations(i))
    
    # modified_classes = modified_classes if modified_classes else ' ، '.join(modified_classes)
    # modified_classes = sorted(set(modified_classes))
    mawad = [madah.replace('أ', 'ا').replace('إ', 'ا') for madah in mawad]
    mawad = sorted(set(mawad))
    mawad = ' ، '.join(mawad)
    context = {'school':school_name 
            ,'classes' : ' , '.join(modified_classes) 
            ,'subjects' : mawad
            ,'teacher' : teacher if teacher else userInfo['first_name']+' '+ userInfo['middle_name'] +' '+ userInfo['last_name'] 
            ,'y1' : melady2 
            ,'y2' : melady1}
    fill_doc(template_dir+first_page , context , outdir+first_page )
    generate_pdf(outdir+first_page , outdir ,first_page)
    
    input_files = get_pdf_files(outdir)
    # Put ready_side_mark_first_page first on the list.
    input_files.sort(reverse=False)
    input_files.insert(0, input_files.pop(input_files.index(outdir+first_page.replace('docx','pdf'))))
    output_path = "السجل الجانبي.pdf"
    merge_pdfs(input_files, outdir+output_path)
    [delete_file(i) for i in input_files]
    
def merge_pdfs(input_files, output_file):
    """Merges a list of PDF files into a single PDF file.

    Args:
    input_files: A list of PDF files to merge.
    output_file: The name of the output PDF file.

    Returns:
    None.
    """

    merger = PdfFileMerger()

    for file in input_files:
        merger.append(file)

    merger.write(output_file)

def get_pdf_files(directory):
    pdf_files = []

    for filename in os.listdir(directory):
        if filename.endswith('.pdf'):
            pdf_files.append(os.path.join(directory, filename))

    return pdf_files

def delete_pdf_page(input_path, output_path, page_number):
    with open(input_path, 'rb') as file:
        pdf_reader = PyPDF4.PdfFileReader(file)
        total_pages = len(pdf_reader.pages)

        if page_number < 0 or page_number >= total_pages:
            print("Invalid page number.")
            return

        pdf_writer = PyPDF4.PdfFileWriter()

        for page in range(total_pages):
            if page != page_number:
                pdf_writer.addPage(pdf_reader.pages[page])

        with open(output_path, 'wb') as output_file:
            pdf_writer.write(output_file)

        print("Page deletion completed.")

def create_zip(file_paths, zip_name='ملف مضغوط' , zip_path='./send_folder/', extension='.rar'):
    zip_file_path = os.path.join(zip_path, zip_name + extension)
    
    with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in file_paths:
            zipf.write(file_path , arcname=os.path.basename(file_path))
            
def Read_E_Side_Note_Marks_ods(file_path=None, file_content=None):
    if file_content is None:
        doc = ezodf.opendoc(file_path)
    else:
        # Save the file content to a temporary file
        with tempfile.NamedTemporaryFile(delete=True) as temp_file:
            temp_file.write(file_content.getvalue())
            temp_file.flush()
            doc = ezodf.opendoc(temp_file.name)

    sheets = [sheet for sheet in doc.sheets][:-1]
    info_sheet = [sheet for sheet in doc.sheets][-1]
    read_file_output_lists = []

    for sheet in sheets:
        data = []

        for i, row in enumerate(sheet.rows()):
            if i < 2:
                continue  # Skip the first two rows
            row_data = [cell.value if cell.value is not None else '' for cell in row]
            if row_data[1] != '':
                dic = {
                    'id': int(row_data[1]),
                    'name': row_data[2],
                    'term1': {'assessment1': int(row_data[3]) if row_data[3] != '' else '', 'assessment2': int(row_data[4]) if row_data[4] != '' else '', 'assessment3': int(row_data[5]) if row_data[5] != '' else '', 'assessment4': int(row_data[6]) if row_data[6] != '' else ''},
                    'term2': {'assessment1': int(row_data[8]) if row_data[8] != '' else '', 'assessment2': int(row_data[9]) if row_data[9] != '' else '', 'assessment3': int(row_data[10]) if row_data[10] != '' else '', 'assessment4': int(row_data[11]) if row_data[11] != '' else ''}
                }
                data.append(dic)

        temp_dic = {'class_name': sheet.name, "students_data": data}
        read_file_output_lists.append(temp_dic)

    modified_classes = []

    classes = [i['class_name'].split('=')[0] for i in read_file_output_lists]
    mawad = [i['class_name'].split('=')[1] for i in read_file_output_lists]
    for i in classes:
        modified_classes.append(mawad_representations(i))

    school_name = info_sheet['A1'].value.split('=')[0]
    school_id = info_sheet['A1'].value.split('=')[1]
    modeeriah = info_sheet['A2'].value
    hejri1 = info_sheet['A3'].value
    hejri2 = info_sheet['A4'].value
    melady1 = info_sheet['A5'].value
    melady2 = info_sheet['A6'].value
    baldah = info_sheet['A7'].value
    modified_classes = ' ، '.join(modified_classes)
    mawad = sorted(set(mawad))
    mawad = ' ، '.join(mawad)
    teacher = info_sheet['A8'].value
    required_data_mrks_text = info_sheet['A9'].value
    period_id = info_sheet['A10'].value

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
        'teacher': teacher,
        'modeeriah_20_2': f'لواء {modeeriah}',
        'hejri_20_1': hejri1,
        'hejri_20_2': hejri2,
        'melady_20_1': melady1,
        'melady_20_2': melady2,
        'baldah_20_2': baldah,
        'school_20_2': school_name,
        'classes_20_2': modified_classes,
        'mawad_20_2': mawad,
        'teacher_20_2': teacher,
        'modeeriah_20_1': f'لواء {modeeriah}',
        'hejri1': hejri1,
        'hejri2': hejri2,
        'melady1': melady1,
        'melady2': melady2,
        'baldah_20_1': baldah,
        'school_20_1': school_name,
        'classes_20_1': modified_classes,
        'mawad_20_1': mawad,
        'teacher_20_1': teacher,
        'period_id': period_id,
        'school_id': school_id
    }

    required_data_mrks_dic_list = {
        int(item.split('-')[0]):
            {
                'assessment_grade_id': int(item.split('-')[1].split(',')[0]),
                'grade_id': int(item.split(',')[0].split('-')[2]),
                'assessments_period_ids': item.split(',')[1:]
            }
        for item in required_data_mrks_text.split('\\\\')
    }

    read_file_output_dict = {'file_data': read_file_output_lists,
                             'custom_shapes': custom_shapes,
                             'required_data_for_mrks_enter': required_data_mrks_dic_list}

    return read_file_output_dict

def upload_marks(username , password , classess_data ):
        auth = get_auth(username , password)
        period_id = classess_data['custom_shapes']['period_id']
        school_id = classess_data['custom_shapes']['school_id']
        # term1_assessment_codes = ['S1A1', 'S1A2', 'S1A3', 'S1A4']
        # term2_assessment_codes = ['S2A1', 'S2A2', 'S2A3', 'S2A4']
        assessment_codes = ['S1A1', 'S1A2', 'S1A3', 'S1A4' , 'S2A1', 'S2A2', 'S2A3', 'S2A4']
        assessment_code_dic = {'S1A1': {'term' :'term1' , 'assess' : 'assessment1'},
                                'S1A2': {'term' :'term1' , 'assess' : 'assessment2'},
                                'S1A3': {'term' :'term1' , 'assess' : 'assessment3'},
                                'S1A4': {'term' :'term1' , 'assess' : 'assessment4'},
                                'S2A1': {'term' :'term2' , 'assess' : 'assessment1'},
                                'S2A2': {'term' :'term2' , 'assess' : 'assessment2'},
                                'S2A3': {'term' :'term2' , 'assess' : 'assessment3'},
                                'S2A4': {'term' :'term2' , 'assess' : 'assessment4'}}
        
        assessments_periods_data = classess_data['required_data_for_mrks_enter']
        for class_data in classess_data['file_data']:
            class_id = class_data['class_name'].split('=')[2] 
            class_subject = class_data['class_name'].split('=')[3]
            class_name = classess_data['file_data'][1]['class_name'].split('=')[0]
            if 'عشر' not in class_name : 
                students_marks_ids = class_data['students_data']
                assessment_grade_id = assessments_periods_data[int(class_id)]['assessment_grade_id']
                grade_id = assessments_periods_data[int(class_id)]['grade_id']
                assessment_periods = get_editable_assessments(auth,username,assessment_grade_id,class_subject)
                # assessment_ids = assessments_periods_data[int(class_id)]['assessments_period_ids']
                # s1a1, s1a2, s1a3, s1a4, s2a1, s2a2, s2a3, s2a4 = [assessment_ids[i] if i < len(assessment_ids) else None for i in range(8)]
                for student_info in students_marks_ids:
                    for code in assessment_codes:
                        if len([i for i in assessment_periods if code in i['code']]) != 0:
                            assessment_period_id = [i for i in assessment_periods if code in i['code']][0]['AssesId']
                            term = assessment_code_dic[code]['term']
                            assess = assessment_code_dic[code]['assess']
                            term_marks = student_info[term]
                            mark = term_marks.get(assess)
                            if mark != '':
                                enter_mark(
                                        auth,
                                        marks=str("{:.2f}".format(float(mark))),
                                        assessment_grading_option_id=8,
                                        assessment_id=assessment_grade_id,
                                        education_subject_id=class_subject,
                                        education_grade_id=grade_id,
                                        institution_id=school_id,
                                        academic_period_id=period_id,
                                        institution_classes_id=class_id,
                                        student_status_id=1,
                                        student_id=student_info['id'],
                                        assessment_period_id=assessment_period_id
                                        )                   

def scrape_schools(username, password , limit = 10, pages = 10*6 ,sector=11):
    dic_list = []
    for page in range(1,pages):
        auth = get_auth(username,password)
        institutions = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit={limit}&institution_sector_id={sector}&_page={page}&_fields=name,code,address,institution_sector_id,area_id,area_administrative_id,longitude,latitude')
        if len(institutions['data']) == 0:
            break
        else:
            dic_list.append(institutions['data'])
    return dic_list

def Vacancies (username , password , schools_nats):
    dic_list=[]
    faulty_inst_nat = []
    school_name_code = []
    error = []
    try:
        for school_nat in schools_nats:
            auth = get_auth(username,password)
            school_name_staff = get_school_teachers(auth,nat_school=school_nat)
            teachers = school_name_staff['staff']
            school_name = school_name_staff['school_code_name']
            school_id = school_name_staff['school_id']
            school_load = get_school_load(auth, school_id)
            teachers_load = get_school_teachers_load(auth , school_id)


            working_teachers = [teacher['name'] for teacher in teachers if teacher['staff_status'] == 1]
            sub_teachers = [teacher['name'] for teacher in teachers if teacher['staff_type'] == 197605]
            english_teachers = [name for name in [ i for i in get_teacher_load_with_name(teachers_load , 1)] if name[0] in working_teachers]
            arabic_teachers = [name for name in [ i for i in get_teacher_load_with_name(teachers_load , 2)] if name[0] in working_teachers]
            math_teachers = [name for name in [ i for i in get_teacher_load_with_name(teachers_load , 3)] if name[0] in working_teachers]
            english_teachers_final = [[name[0]+'**',name[1],name[2]] if name[0] in sub_teachers else name for name in english_teachers]
            arabic_teachers_final = [[name[0]+'**',name[1],name[2]] if name[0] in sub_teachers else name for name in arabic_teachers]
            math_teachers_final = [[name[0]+'**',name[1],name[2]] if name[0] in sub_teachers else name for name in math_teachers]

            string = str(school_load['english_school_sum'])+' <== نصاب الانجليزي \n'+str(school_load['arabic_school_sum'])+' <== نصاب العربي \n'+str(school_load['math_school_sum'])+' <== نصاب الرياضيات \n'
            classes = ' ,\n'.join(str(i).replace('الصف', '') for i in school_load['classes'])

            long_string = '--------------معلمين الانجليزي--------------\n'
            for item in english_teachers_final:
                long_string += item[0]+' =======>> '+ str(item[1]) + ' =======>> ' +  ' , '.join(str(i).replace('الصف', '') for i in item[2])+'\n'
            long_string += '--------------معلمين العربي--------------\n'
            for item in arabic_teachers_final:
                long_string += item[0]+' =======>> '+ str(item[1]) + ' =======>> ' +  ' , '.join(str(i).replace('الصف', '') for i in item[2])+'\n'
            long_string += '--------------معلمين الرياضيات--------------\n'
            for item in math_teachers_final:
                long_string += item[0]+' =======>> '+ str(item[1]) + ' =======>> ' +  ' , '.join(str(i).replace('الصف', '') for i in item[2])+'\n'

            dic = { 'school_name' :school_name , 'school_load' : string , 'teachers' : long_string , 'classes': classes }
            dic_list.append(dic)
            
    except Exception as e : 
        faulty_inst_nat.append(school_nat)
        error.append(e)
        try:
            school_name_code.append(school_name_staff['school_code_name'])
        except:
            pass
    return dic_list

def get_school_load(auth , inst_id ,academic_period_id=13):
    student_classess = make_request(auth=auth, url=f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-InstitutionClassStudents.json?institution_id={inst_id}&academic_period_id={academic_period_id}&_limit=0&_contain=Users.Genders')['data']
    institution_class_ids = list(set([i['institution_class_id'] for i in student_classess]))
    joined_string = ','.join(str(i) for i in [f'institution_class_id:{i}' for i in institution_class_ids])
    classes_data = make_request(auth=auth,url='https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere='+joined_string)['data']
    class_list = []
    for i in classes_data:
        class_list.append({'class_id': i['institution_class_id'] , 'class_name': i['institution_class']['name'] })
        class_dict = {i['class_id']: i['class_name'] for i in class_list if i['class_id'] != ''}
        classes = [key for value,key in class_dict.items()]

    arabic_class_sum = 0
    english_class_sum = 0
    math_class_sum = 0

    for school_class in classes:
        if 'اول' in school_class or 'ثاني' in school_class or 'ثالث'in school_class or  'رابع' in school_class or 'عشر' in school_class:
            if 'دبي' in school_class:
                math_class_sum+=3
            else:
                math_class_sum+=5
            english_class_sum += 4
        else: 
            english_class_sum+= 5
            math_class_sum += 5

    for school_class in classes:
        if 'سابع' in school_class or 'ثامن' in school_class or 'تاسع' in school_class or  'عاشر' in school_class :
            arabic_class_sum += 6
        else:
            if 'عشر' in school_class:
                if 'دبي' in school_class:
                    if  'حادي' in school_class:
                        arabic_class_sum+=5
                    elif 'ثاني' in school_class:
                        arabic_class_sum+=4
                else:
                    arabic_class_sum+=4
            else:
                arabic_class_sum+=7
    return {'english_school_sum' : english_class_sum , 'arabic_school_sum' : arabic_class_sum , 'math_school_sum' :  math_class_sum , 'classes' :  classes}

def get_school_teachers(auth ,id=None , nat_school=None ,session=None ,row=False):
    if id == None:
        teachers =make_request(auth=auth, url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&_orWhere=code:{nat_school}&_contain=Staff.Users,Staff.Positions',session=session)
    else:
        teachers =make_request(auth=auth, url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={id}&_contain=Staff.Users,Staff.Positions',session=session)
    dic_list=[]
    for teacher in teachers['data'][0]['staff'] : 
        if teacher['staff_status_id'] == 1:
        # print(counter ,'-' , each['staff_name'])
            # dic_list.append(teacher['staff_name'])
            dic_list.append({'staffId':teacher['staff_id'],'name':teacher['name'],'name_list':[teacher['user']['first_name'], teacher['user']['middle_name'],teacher['user']['third_name'],teacher['user']['last_name']],'position':teacher['position']['name'],'birthDate':teacher['user']['date_of_birth'], 'nat_id':teacher['user']['identity_number'],'default_nat_id':teacher['user']['default_identity_type'],'staff_type':teacher['staff_type_id'] , 'staff_status': teacher['staff_status_id']})
    if row :
        return teachers
    else:
        return {'school_code_name' : teachers['data'][0]['code_name'], 'staff' : dic_list}
    
def get_school_teachers_load(auth , inst_id , academic_period_id=13):
    school_load = make_request(auth=auth,url=f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-InstitutionSubjectStaff.json?institution_id={inst_id}&_contain=Users,InstitutionSubjects&academic_period_id={academic_period_id}&_limit=0')['data']
    
    institution_class_ids = list(set([i['institution_subject_id'] for i in school_load]))
    class_list = []

    for i in range(0, len(institution_class_ids), 20):
        start = i
        end = i+19 if i+19 < len(institution_class_ids) else i+(len(institution_class_ids)-i)-1
        class_ids = [i for i in  institution_class_ids[start:end]]
        joined_string = ','.join(str(i) for i in [f'institution_subject_id:{i}' for i in class_ids])
        classes_data = make_request(auth=auth,url='https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere='+joined_string)['data']
        for i in classes_data:
            class_list.append({'class_id': i['institution_subject_id'] , 'class_name': i['institution_class']['name'] })
            class_dict = {i['class_id']: i['class_name'] for i in class_list if i['class_id'] != ''}
            classes = [key for value,key in class_dict.items()]
            
    grade_data = get_grade_info(auth)
    grade_list = []
    for i in grade_data:
        grade_list.append({'grade_id': i['education_grade_id'] , 'grade_name': re.sub('.*للصف','الصف', i['name']) })
        grade_dict = {i['grade_id']: i['grade_name'] for i in grade_list if i['grade_name'] != ''}
    grade_dict

    school_load_dictionary = []
    for load in school_load:
        try:
            school_load_dictionary.append({'name':load['user']['name'],'subject':load['institution_subject']['name'],'grade':class_dict[load['institution_subject_id']]})
        except:
            try:
                school_load_dictionary.append({'name':load['user']['name'],'subject':load['institution_subject']['name'],'grade':grade_dict[load['institution_subject']['education_grade_id']]})
            except:
                school_load_dictionary.append({'name':load['user']['name'],'subject':load['institution_subject']['name'],'grade':load['institution_subject_id']})
    return school_load_dictionary   

def get_grade_info(auth):
    
    my_list = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0')['data']
    return my_list

def count_teachers_grades(teachers_load):
    english_teachers = [item for item in teachers_load if 'الانجليزية' in item['subject']]
    arabic_teachers = [item for item in teachers_load if 'العربية' in item['subject']]
    math_teachers = [item for item in teachers_load if 'رياضيات' in item['subject']]

    unique_english_teachers = set(item['name'] for item in english_teachers )
    unique_arabic_teachers = set(item['name'] for item in arabic_teachers)
    unique_math_teachers = set(item['name'] for item in math_teachers)
    
    loads = {'english': [], 'arabic': [], 'math': []}
    teachers = {'english': english_teachers, 'arabic': arabic_teachers, 'math': math_teachers}
    unique_teachers = {'english': unique_english_teachers, 'arabic': unique_arabic_teachers, 'math': unique_math_teachers}

    for subject in loads.keys():
        for u_name in list(unique_teachers[subject]):
            load = [teacher for teacher in teachers[subject] if teacher['name'] == u_name]
            loads[subject].append(load)
    return loads

def get_teacher_load_with_name(teachers_load , subject):
    if subject == 1 :
        subject = 'english'
        subject_sum = 'english_class_sum'
    elif subject == 2 :
        subject = 'arabic'
        subject_sum = 'arabic_class_sum'
    elif subject == 3 :
        subject = 'math'
        subject_sum = 'math_class_sum'
    teachers = []
    for group in count_teachers_grades(teachers_load)[subject]:
        grades = [grade['grade'] for grade in group ] 
        teachers.append((group[0]['name'],count_teacher_load(grades)[subject_sum],count_teacher_load(grades)['classes']))
    return teachers

def count_teacher_load(classes):
    
    arabic_class_sum = 0
    english_class_sum = 0
    math_class_sum = 0

    for school_class in classes:
        if 'اول' in school_class or 'ثاني' in school_class or 'ثالث'in school_class or  'رابع' in school_class or 'عشر' in school_class:
            if 'دبي' in school_class:
                math_class_sum+=3
            else:
                math_class_sum+=5
            english_class_sum += 4
        else: 
            english_class_sum+= 5
            math_class_sum += 5

    for school_class in classes:
        if 'سابع' in school_class or 'ثامن' in school_class or 'تاسع' in school_class or  'عاشر' in school_class :
            arabic_class_sum += 6
        else:
            if 'عشر' in school_class:
                if 'دبي' in school_class and 'خصص' in school_class:
                    if  'حادي' in school_class:
                        arabic_class_sum+=5
                    elif 'ثاني' in school_class:
                        arabic_class_sum+=4
                else:
                    arabic_class_sum+=4
            else:
                arabic_class_sum+=7
    return {'english_class_sum' : english_class_sum , 'arabic_class_sum' : arabic_class_sum , 'math_class_sum' :  math_class_sum , 'classes' :  classes}

def create_tables_wrapper(username , password ,term2=False): 
    session = requests.Session()
    auth = get_auth(username, password)
    student_info_marks = get_students_info_subjectsMarks( username , password ,session)
    dic_list4 = student_info_marks
    grouped_list = group_students(dic_list4 )
    

    add_subject_sum_dictionary(grouped_list)
    add_averages_to_group_list(grouped_list ,skip_art_sport=False)
    
    # save_dictionary_to_json_file(dictionary={'grouped_list':grouped_list})
    create_tables(auth , grouped_list ,term2=term2 )
        
def create_certs_wrapper(username , password ,term2=False):
    student_info_marks = get_students_info_subjectsMarks( username , password )
    dic_list4 = student_info_marks
    grouped_list = group_students(dic_list4 )
    

    add_subject_sum_dictionary(grouped_list)
    add_averages_to_group_list(grouped_list ,skip_art_sport=False)
    
    create_certs(grouped_list , term2=term2)

def create_tables(auth , grouped_list ,term2=False ,template='./templet_files/tamplete_table.xlsx'  , outdir='./send_folder/'):
    # auth = get_auth(username , password)
    institution_area_data = inst_area(auth)
    institution_data = inst_name(auth)
    curr_year_code = get_curr_period(auth)['data'][0]['code']

    for group in grouped_list:
        
        if 'عشر' not in group[0]['student_grade_name']:
            template_file = openpyxl.load_workbook(template)
            marks_sheet = template_file.worksheets[2]

        
            for row_number, dataFrame in enumerate(sort_dictionary_list_based_on(group ,dictionary_key='student__full_name',simple=False,reverse=False), start=4):
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
                marks_sheet.cell(row=row_number, column=9).value = islam_subject[0][1] if term2 and islam_subject and len(islam_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=10).value = islam_subject[0][0]+islam_subject[0][1] if term2 and islam_subject and len(islam_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=11).value = arabic_subject[0][0] if arabic_subject and len(arabic_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=12).value = arabic_subject[0][1] if term2 and arabic_subject and len(arabic_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=13).value = arabic_subject[0][0]+arabic_subject[0][1] if term2 and arabic_subject and len(arabic_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=14).value = english_subject[0][0] if english_subject and len(english_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=15).value = english_subject[0][1] if term2 and english_subject and len(english_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=16).value = english_subject[0][0]+english_subject[0][1] if term2 and english_subject and len(english_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=17).value = math_subject[0][0] if math_subject and len(math_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=18).value = math_subject[0][1] if term2 and math_subject and len(math_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=19).value = math_subject[0][0]+math_subject[0][1] if term2 and math_subject and len(math_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=20).value = social_subjects[0][0] if social_subjects and len(social_subjects[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=21).value = social_subjects[0][1] if term2 and social_subjects and len(social_subjects[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=22).value = social_subjects[0][0]+social_subjects[0][1] if term2 and social_subjects and len(social_subjects[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=23).value = science_subjects[0][0] if science_subjects and len(science_subjects[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=24).value = science_subjects[0][1] if term2 and science_subjects and len(science_subjects[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=25).value = science_subjects[0][0]+science_subjects[0][1] if term2 and science_subjects and len(science_subjects[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=26).value = art_subject[0][0] if art_subject and len(art_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=27).value = art_subject[0][1] if term2 and art_subject and len(art_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=28).value = art_subject[0][0]+art_subject[0][1] if term2 and art_subject and len(art_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=29).value = sport_subject[0][0] if sport_subject and len(sport_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=30).value = sport_subject[0][1] if term2 and sport_subject and len(sport_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=31).value = sport_subject[0][0]+sport_subject[0][1] if term2 and sport_subject and len(sport_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=32).value = financial_subject[0][0] if financial_subject and len(financial_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=33).value = financial_subject[0][1] if term2 and financial_subject and len(financial_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=34).value = financial_subject[0][0]+financial_subject[0][1] if term2 and financial_subject and len(financial_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=35).value = vocational_subject[0][0] if vocational_subject and len(vocational_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=36).value = vocational_subject[0][1] if term2 and vocational_subject and len(vocational_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=37).value = vocational_subject[0][0]+vocational_subject[0][1] if term2 and vocational_subject and len(vocational_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=38).value = computer_subject[0][0] if computer_subject and len(computer_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=39).value = computer_subject[0][1] if term2 and computer_subject and len(computer_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=40).value = computer_subject[0][0]+computer_subject[0][1] if term2 and computer_subject and len(computer_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=41).value = franch_subject[0][0] if franch_subject and len(franch_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=42).value = franch_subject[0][1] if term2 and franch_subject and len(franch_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=43).value = franch_subject[0][0]+franch_subject[0][1] if term2 and franch_subject and len(franch_subject[0]) > 0 else ''
                # marks_sheet.cell(row=row_number, column=44).value = dataFrame[0][] if = and len(=[0]) > 0 else ''
                # marks_sheet.cell(row=row_number, column=45).value = dataFrame[0][] if = and len(=[0]) > 0 else ''
                # marks_sheet.cell(row=row_number, column=46).value = dataFrame[0][] if = and len(=[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=47).value = christian_subject[0][0] if christian_subject and len(christian_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=48).value = christian_subject[0][1] if term2 and christian_subject and len(christian_subject[0]) > 0 else ''
                marks_sheet.cell(row=row_number, column=49).value = christian_subject[0][0]+christian_subject[0][1] if term2 and christian_subject and len(christian_subject[0]) > 0 else ''

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
                marks_sheet['h3'],marks_sheet['i3'],marks_sheet['j3'] = [100]*3
                # عربية 
                # k/l/m
                marks_sheet['k3'],marks_sheet['l3'],marks_sheet['m3'] = [100]*3
                # انجليزية 
                # n/o/p
                marks_sheet['n3'],marks_sheet['o3'],marks_sheet['p3'] = [100]*3
                # رياضيات
                # q/r/s
                marks_sheet['q3'],marks_sheet['r3'],marks_sheet['s3'] = [100]*3
                # اجتماعيات 
                # t/u/v
                marks_sheet['t3'],marks_sheet['u3'],marks_sheet['v3'] = [100]*3
                # علوم
                # w/x/y
                marks_sheet['w3'],marks_sheet['x3'],marks_sheet['y3'] = [100]*3    
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
            
            template_file.save(outdir+' جدول '+group[0]['student_class_name_letter']+'.xlsx')
        
def create_certs(grouped_list , term2=False ,template='./templet_files/a4_gray_cert.xlsx' ,image='./templet_files/Pasted image.png' , outdir='./send_folder/'):
    gray_cert_cell_positions_context = {
    'student_name': 'E7',
    'class_section': 'B11',
    'nationality': 'H9',
    'national_id': 'B9',
    'birthplace_date': 'H7',
    'religion': '',
    'student_address': '',
    'school_name': 'G11',
    'school_address': '',
    'directorate': 'B13',
    'brigade': 'I13',
    'supervising_authority': '',
    'school_phone_number': 'C18',
    'school_national_id': 'C19',
    'academic_year': '',
    'academic_year_1': '',
    'academic_year_2': '',
    'class': '',
    'islamic_education': 'c18',
    'arabic_language': 'c19',
    'english_language': 'c20',
    'mathematics': 'c21',
    'social_studies': 'c22',
    'science': 'c23',
    'visual_arts': 'c24',
    'physical_education': 'c25',
    'vocational_education': 'c26',
    'computer': 'c27',
    'financial_culture': 'c28',
    'french_language': 'c29',
    'christian_religion': 'C30',
    'result': '',
    'average': 'C32',
    'school_days': 'G35',
    'class_teacher_name': 'J35',
    'principal_name': 'I36',
    'semester_1_student_absent': '',
    'semester_2_student_absent': '',
    'student_name': '',
    }
    for group in grouped_list:
        if 'عشر' not in group[0]['student_grade_name']:
            template_file = load_workbook(template)
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

                img = openpyxl.drawing.image.Image(image)
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
                sheet2['E18'] = islam_subject[0][1] if term2 and islam_subject and len(islam_subject[0]) != 0 else ''
                sheet2['F18'] = (islam_subject[0][value_item] + islam_subject[0][1])/2 if term2 and islam_subject and len(islam_subject[0]) != 0 else ''
                if term2:
                    sheet2['G18'] = convert_avarage_to_words((islam_subject[0][value_item] + islam_subject[0][1])/2) if islam_subject else ''
                    sheet2['J18'] = score_in_words(((islam_subject[0][value_item] + islam_subject[0][1])/2),max_mark=maxMark) if islam_subject else ''
                else:
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
                sheet2['E19'] = arabic_subject[0][1] if term2 and arabic_subject and len(arabic_subject[0]) != 0 else ''
                sheet2['F19'] = (arabic_subject[0][value_item] + arabic_subject[0][1])/2 if term2 and arabic_subject and len(arabic_subject[0]) != 0 else ''
                if term2:
                    sheet2['G19'] = convert_avarage_to_words((arabic_subject[0][value_item] + arabic_subject[0][1])/2) if arabic_subject else ''
                    sheet2['J19'] = score_in_words(((arabic_subject[0][value_item] + arabic_subject[0][1])/2),max_mark=maxMark) if arabic_subject else ''
                else:
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
                sheet2['E20'] = english_subject[0][1] if term2 and english_subject and len(english_subject[0]) != 0 else ''
                sheet2['F20'] = (english_subject[0][value_item] + english_subject[0][1])/2 if term2 and english_subject and len(english_subject[0]) != 0 else ''
                if term2:
                    sheet2['G20'] = convert_avarage_to_words((english_subject[0][value_item] + english_subject[0][1])/2) if english_subject else ''
                    sheet2['J20'] = score_in_words(((english_subject[0][value_item] + english_subject[0][1])/2),max_mark=maxMark) if english_subject else ''
                else:
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
                sheet2['E21'] = math_subject[0][1] if term2 and math_subject and len(math_subject[0]) != 0 else ''
                sheet2['F21'] = (math_subject[0][value_item] + math_subject[0][1])/2 if term2 and math_subject and len(math_subject[0]) != 0 else ''
                if term2:
                    sheet2['G21'] = convert_avarage_to_words((math_subject[0][value_item] + math_subject[0][1])/2) if math_subject else ''
                    sheet2['J21'] = score_in_words(((math_subject[0][value_item] + math_subject[0][1])/2),max_mark=maxMark) if math_subject else ''
                else:
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
                sheet2['E22'] = social_subjects[0][1] if term2 and social_subjects and len(social_subjects[0]) != 0 else ''
                sheet2['F22'] = (social_subjects[0][value_item] + social_subjects[0][1])/2 if term2 and social_subjects and len(social_subjects[0]) != 0 else ''
                if term2:
                    sheet2['G22'] = convert_avarage_to_words((social_subjects[0][value_item] + social_subjects[0][1])/2) if social_subjects else ''
                    sheet2['J22'] = score_in_words(int(((social_subjects[0][value_item] + social_subjects[0][1])/2)*(2/3)),max_mark=maxMark) if social_subjects else ''
                else:
                    sheet2['G22'] = convert_avarage_to_words(social_subjects[0][value_item]) if social_subjects else ''
                    sheet2['J22'] = score_in_words(int(social_subjects[0][value_item]*(2/3)),max_mark=maxMark) if social_subjects else ''

                # العلوم
                science_subjects = [value for key ,value in group_item['subject_sums'].items() if 'علوم' in key]
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
                sheet2['E23'] = science_subjects[0][1] if term2 and science_subjects and len(science_subjects[0]) != 0 else ''
                sheet2['F23'] = (science_subjects[0][value_item] + science_subjects[0][1])/2 if term2 and science_subjects and len(science_subjects[0]) != 0 else ''
                if term2:
                    sheet2['G23'] = convert_avarage_to_words((science_subjects[0][value_item] + science_subjects[0][1])/2) if science_subjects else ''
                    sheet2['J23'] = score_in_words( ((science_subjects[0][value_item] + science_subjects[0][1])/2),max_mark=maxMark) if  science_subjects else ''
                else:
                    sheet2['G23'] = convert_avarage_to_words(science_subjects[0][value_item]) if science_subjects else ''
                    sheet2['J23'] = score_in_words( science_subjects[0][value_item],max_mark=maxMark) if  science_subjects else ''

                # التربية الفنية والموسيقية
                art_subject = [value for key ,value in group_item['subject_sums'].items() if 'الفنية والموس' in key]
                sheet2['C24'] = 100 if art_subject and len(art_subject[0]) != 0 else ''
                sheet2['D24'] = art_subject[0][value_item] if art_subject and len(art_subject[0]) != 0 else ''
                sheet2['E24'] = art_subject[0][1] if term2 and art_subject and len(art_subject[0]) != 0 else ''
                sheet2['F24'] = (art_subject[0][value_item] + art_subject[0][1])/2 if term2 and art_subject and len(art_subject[0]) != 0 else ''
                if term2:
                    sheet2['G24'] = convert_avarage_to_words((art_subject[0][value_item] + art_subject[0][1])/2) if art_subject else ''
                    sheet2['J24'] = score_in_words(((art_subject[0][value_item] + art_subject[0][1])/2) ) if art_subject else ''
                else:
                    sheet2['G24'] = convert_avarage_to_words(art_subject[0][value_item]) if art_subject else ''
                    sheet2['J24'] = score_in_words(art_subject[0][value_item] ) if art_subject else ''

                # التربية الرياضية
                sport_subject = [value for key ,value in group_item['subject_sums'].items() if 'رياضية' in key]
                sheet2['C25'] = 100 if sport_subject and len(sport_subject[0]) != 0 else ''
                sheet2['D25'] = sport_subject[0][value_item] if sport_subject and len(sport_subject[0]) != 0 else ''
                sheet2['E25'] = sport_subject[0][1] if term2 and sport_subject and len(sport_subject[0]) != 0 else ''
                sheet2['F25'] = (sport_subject[0][value_item] + sport_subject[0][1])/2 if term2 and sport_subject and len(sport_subject[0]) != 0 else ''
                if term2:
                    sheet2['G25'] = convert_avarage_to_words((sport_subject[0][value_item] + sport_subject[0][1])/2) if sport_subject else ''
                    sheet2['J25'] = score_in_words(((sport_subject[0][value_item] + sport_subject[0][1])/2) ) if sport_subject else ''
                else:
                    sheet2['G25'] = convert_avarage_to_words(sport_subject[0][value_item]) if sport_subject else ''
                    sheet2['J25'] = score_in_words(sport_subject[0][value_item] ) if sport_subject else ''

                # التربية المهنية 
                vocational_subject = [value for key ,value in group_item['subject_sums'].items() if 'مهنية' in key]
                sheet2['C26'] = 100 if vocational_subject and len(vocational_subject[0]) != 0 else ''
                sheet2['D26'] = vocational_subject[0][value_item] if vocational_subject and len(vocational_subject[0]) != 0 else ''
                sheet2['E26'] = vocational_subject[0][1] if term2 and vocational_subject and len(vocational_subject[0]) != 0 else ''
                sheet2['F26'] = (vocational_subject[0][value_item] + vocational_subject[0][1])/2 if term2 and vocational_subject and len(vocational_subject[0]) != 0 else ''
                if term2:
                    sheet2['G26'] = convert_avarage_to_words((vocational_subject[0][value_item] + vocational_subject[0][1])/2) if vocational_subject else ''
                    sheet2['J26'] = score_in_words(((vocational_subject[0][value_item] + vocational_subject[0][1])/2) ) if vocational_subject else ''
                else:
                    sheet2['G26'] = convert_avarage_to_words(vocational_subject[0][value_item]) if vocational_subject else ''
                    sheet2['J26'] = score_in_words(vocational_subject[0][value_item] ) if vocational_subject else ''

                # الحاسوب
                computer_subject = [value for key ,value in group_item['subject_sums'].items() if 'حاسوب' in key]
                sheet2['C27'] = 100 if computer_subject and len(computer_subject[0]) != 0 else ''
                sheet2['D27'] = computer_subject[0][value_item] if computer_subject and len(computer_subject[0]) != 0 else ''
                sheet2['E27'] = computer_subject[0][1] if term2 and computer_subject and len(computer_subject[0]) != 0 else ''
                sheet2['F27'] = (computer_subject[0][value_item] + computer_subject[0][1])/2 if term2 and computer_subject and len(computer_subject[0]) != 0 else ''
                if term2:
                    sheet2['G27'] = convert_avarage_to_words((computer_subject[0][value_item] + computer_subject[0][1])/2) if computer_subject else ''
                    sheet2['J27'] = score_in_words(((computer_subject[0][value_item] + computer_subject[0][1])/2) ) if computer_subject else ''
                else:
                    sheet2['G27'] = convert_avarage_to_words(computer_subject[0][value_item]) if computer_subject else ''
                    sheet2['J27'] = score_in_words(computer_subject[0][value_item] ) if computer_subject else ''

                # الثقافة المالية
                financial_subject = [value for key ,value in group_item['subject_sums'].items() if 'مالية' in key]
                sheet2['C28'] = 100 if financial_subject and len(financial_subject[0]) != 0 else ''
                sheet2['D28'] = financial_subject[0][value_item] if financial_subject and len(financial_subject[0]) != 0 else ''
                sheet2['E28'] = financial_subject[0][1] if term2 and financial_subject and len(financial_subject[0]) != 0 else ''
                sheet2['F28'] = (financial_subject[0][value_item] + financial_subject[0][1])/2 if term2 and financial_subject and len(financial_subject[0]) != 0 else ''
                if term2:
                    sheet2['G28'] = convert_avarage_to_words((financial_subject[0][value_item] + financial_subject[0][1])/2) if financial_subject else ''
                    sheet2['J28'] = score_in_words(((financial_subject[0][value_item] + financial_subject[0][1])/2) ) if financial_subject else ''
                else:
                    sheet2['G28'] = convert_avarage_to_words(financial_subject[0][value_item]) if financial_subject else ''
                    sheet2['J28'] = score_in_words(financial_subject[0][value_item] ) if financial_subject else ''

                # اللغة الفرنسية 
                franch_subject = [value for key ,value in group_item['subject_sums'].items() if 'فرنسية' in key]
                sheet2['C29'] = 100 if franch_subject and len(franch_subject[0]) != 0 else ''
                sheet2['D29'] = franch_subject[0][value_item] if franch_subject and len(franch_subject[0]) != 0 else ''
                sheet2['E29'] = franch_subject[0][1] if term2 and franch_subject and len(franch_subject[0]) != 0 else ''
                sheet2['F29'] = (franch_subject[0][value_item] + franch_subject[0][1])/2 if term2 and franch_subject and len(franch_subject[0]) != 0 else ''
                if term2:
                    sheet2['G29'] = convert_avarage_to_words((franch_subject[0][value_item] + franch_subject[0][1])/2) if franch_subject else ''
                    sheet2['J29'] = score_in_words(((franch_subject[0][value_item] + franch_subject[0][1])/2) ) if franch_subject else ''
                else:
                    sheet2['G29'] = convert_avarage_to_words(franch_subject[0][value_item]) if franch_subject else ''
                    sheet2['J29'] = score_in_words(franch_subject[0][value_item] ) if franch_subject else ''

                # الدين المسيحي
                christian_subject = [value for key ,value in group_item['subject_sums'].items() if 'الدين المسيحي' in key]
                sheet2['C30'] = 100 if christian_subject and len(christian_subject[0]) != 0 else ''
                sheet2['D30'] = christian_subject[0][value_item] if christian_subject and len(christian_subject[0]) != 0 else ''
                sheet2['E30'] = christian_subject[0][1] if term2 and christian_subject and len(christian_subject[0]) != 0 else ''
                sheet2['F30'] = (christian_subject[0][value_item] + christian_subject[0][1])/2 if term2 and christian_subject and len(christian_subject[0]) != 0 else ''
                if term2:
                    sheet2['G30'] = convert_avarage_to_words((christian_subject[0][value_item] + christian_subject[0][1])/2) if christian_subject else ''
                    sheet2['J30'] = score_in_words(((christian_subject[0][value_item] + christian_subject[0][1])/2) ) if christian_subject else ''
                else:
                    sheet2['G30'] = convert_avarage_to_words(christian_subject[0][value_item]) if christian_subject else ''
                    sheet2['J30'] = score_in_words(christian_subject[0][value_item] ) if christian_subject else ''

                if term2 :
                    # عدل المئوي بالرقام 
                    sheet2['c32']= group_item['t1+t2+year_avarage'][2]
                    #بالحروف
                    sheet2['e32']= convert_avarage_to_words(group_item['t1+t2+year_avarage'][2]) if group_item else ''
                    #ترتيب الطالب على الصف 
                    sheet2['j32']= counter-1
                    #النتيجة 
                    sheet2['b33']= 'مقصر' if any(int(sum(item)/2) > 49 for item in [value for key , value in group_item['subject_sums'].items()] ) else score_in_words(int(group_item['t1+t2+year_avarage'][2]))
                else:
                    
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
            template_file.save(outdir+group[0]['student_class_name_letter']+'.xlsx')

def create_coloured_certs_excel(grouped_list , term2=False ,template='./templet_files/نموذج شهادات ملونة.xlsx' , outdir='./send_folder/'):
    
    colored_cert_cells_position_context = {
                                        'student_name': 'E11',
                                        'student_name2': 'B44',
                                        'class_section': 'E12',
                                        'class_name':'B45',
                                        'nationality': 'D13',
                                        'national_id': 'E14',
                                        'birthplace_date': 'E15',
                                        'religion': 'E16',
                                        'student_address': 'E17',
                                        'school_name': 'E18',
                                        'school_address': 'E19',
                                        'directorate': 'D20',
                                        'school_bridge': 'D21',
                                        'supervising_authority': 'E22',
                                        'school_phone_number': 'E23',
                                        'academic_year_1': 'F42',
                                        'academic_year_2': 'G42',
                                        'class': 'B45',
                                        'islamic_education': 'E50',
                                        'arabic_language': 'E51',
                                        'english_language': 'E52',
                                        'mathematics': 'E53',
                                        'social_studies': 'E54',
                                        'science': 'E55',
                                        'visual_arts': 'E56',
                                        'physical_education': 'E57',
                                        'vocational_education': 'E58',
                                        'computer': 'E59',
                                        'financial_culture': 'E60',
                                        'french_language': 'E61',
                                        'christian_religion': 'E62',
                                        'average': 'B67',
                                        'school_days': 'I64',
                                        'class_teacher_name': 'B69',
                                        'principal_name': 'L69',
                                        'semester_1_student_absent': 'H66',
                                        'semester_2_student_absent': 'L66',
                                        }
    
    c = colored_cert_cells_position_context
    result_cell_positions = ['B64','B65','B66']
    
    statistic_data = grouped_list['students_info']
    assessments_data_groups = grouped_list['assessments_data']
    term = 1 if term2 else 0
    
    
    for group in assessments_data_groups:
        if not any("عشر" in i['student_grade_name'] for i in group) : 
            template_file = load_workbook(template)
            sheet1 = template_file.worksheets[0]
            
            names_averages =  sort_dictionary_list_based_on(group,item_in_list=term)

            group = sort_dictionary_list_based_on(group ,simple=False,item_in_list=term)

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
                
                wanted_id = group_item['student_id']
                student_statistic_info = [i for i in statistic_data if i['student_id'] == wanted_id ][0]

                # img = openpyxl.drawing.image.Image(image)
                # img.anchor = 'e2'
                # sheet2.add_image(img)

                # print(sheet2)
                sheet2[c['student_name']] = group_item['student__full_name']
                # مكان و تاريخ الولادة
                sheet2[c['birthplace_date']]= str(group_item['student_birth_place']) + ' ' + str(group_item['student_birth_date'])
                #الرقم الوطني
                sheet2[c['national_id']]= group_item['student_nat_id']
                #الجنسية
                sheet2[c['nationality']]= group_item['student_nat']
                #الصف و الشعبة 
                sheet2[c['class_section']]= group_item['student_class_name_letter']
                # صف الطالب
                sheet2[c['class_name']]= group_item['student_grade_name'].replace('الصف' ,'')
                #المدرسة و رقمها الوطني
                sheet2[c['school_name']]= group_item['student_school_name']
                #المنطقة التعليمية  او السلطة المشرفة
                sheet2[c['supervising_authority']]= 'وزارة التربية و التعليم'
                # مديرية المدرسة 
                sheet2[c['directorate']]= grouped_list['school_directorate']
                # لواء المدرسة
                sheet2[c['school_bridge']]= grouped_list['school_bridge']
                # الديانة
                sheet2[c['religion']]= student_statistic_info['religion']
                # عنوان الطالب
                sheet2[c['student_address']]= student_statistic_info['address']
                # عنوان المدرسة
                sheet2[c['school_address']]= grouped_list['school_address']
                # العام الدراسي الاول 
                sheet2[c['academic_year_1']]= grouped_list['academic_year_1']
                # العام الدراسي الثاني 
                sheet2[c['academic_year_2']]= grouped_list['academic_year_2']
                # الاسم على الوجه الثاني
                sheet2[c['student_name2']] = group_item['student__full_name']
                # غياب الفصل الاول 
                sheet2[c['semester_1_student_absent']]= ''
                # غياب الفصل الثاني
                sheet2[c['semester_2_student_absent']]= ''
                
                # put the subjects cells inder here 
                i ='c,d,e,g,j,f'.split(',')
                r = range(18,32)


                # التربية الاسلامية
                islam_subject = [value for key ,value in group_item['subject_sums'].items() if 'سلامية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    islam_subject_maxMark = 200
                    sheet2[c['islamic_education']] = islam_subject_maxMark /2
                    sheet2['F'+c['islamic_education'][1:]] = islam_subject_maxMark
                else:
                    islam_subject_maxMark = 100
                    sheet2[c['islamic_education']] = islam_subject_maxMark /2
                    sheet2['F'+c['islamic_education'][1:]] = islam_subject_maxMark
                sheet2['H'+c['islamic_education'][1:]] = islam_subject[0][0] if islam_subject and len(islam_subject[0]) != 0 else ''
                sheet2['I'+c['islamic_education'][1:]] = islam_subject[0][1] if term2 and islam_subject and len(islam_subject[0]) != 0 else ''
                sheet2['K'+c['islamic_education'][1:]] = (islam_subject[0][0] + islam_subject[0][1])/2 if term2 and islam_subject and len(islam_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G'+c['islamic_education']] = convert_avarage_to_words((islam_subject[0][0] + islam_subject[0][1])/2) if islam_subject else ''
                #     sheet2['J'+c['islamic_education']] = score_in_words(((islam_subject[0][0] + islam_subject[0][1])/2),max_mark=maxMark) if islam_subject else ''
                # else:
                #     sheet2['G'+c['islamic_education']] = convert_avarage_to_words(islam_subject[0][0]) if islam_subject else ''
                #     sheet2['J'+c['islamic_education']] = score_in_words(islam_subject[0][0],max_mark=maxMark) if islam_subject else ''

                # اللغة العربية
                arabic_subject = [value for key ,value in group_item['subject_sums'].items() if 'عربية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    arabic_subject_maxMark = 300
                    sheet2[c['arabic_language']] = arabic_subject_maxMark / 2
                    sheet2['F'+c['arabic_language'][1:]] = arabic_subject_maxMark
                else:
                    arabic_subject_maxMark = 100
                    sheet2[c['arabic_language']] = arabic_subject_maxMark / 2
                    sheet2['F'+c['arabic_language'][1:]] = arabic_subject_maxMark
                sheet2['H' + c['arabic_language'][1:]] = arabic_subject[0][0] if arabic_subject and len(arabic_subject[0]) != 0 else ''
                sheet2['I' + c['arabic_language'][1:]] = arabic_subject[0][1] if term2 and arabic_subject and len(arabic_subject[0]) != 0 else ''
                sheet2['K' + c['arabic_language'][1:]] = (arabic_subject[0][0] + arabic_subject[0][1])/2 if term2 and arabic_subject and len(arabic_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['J19'] = score_in_words(((arabic_subject[0][0] + arabic_subject[0][1])/2),max_mark=maxMark) if arabic_subject else ''
                #     sheet2['F'+] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words((arabic_subject[0][0] + arabic_subject[0][1])/2) if arabic_subject else maxMark
                # else:
                #     sheet2['J19'] = score_in_words(arabic_subject[0][0],max_mark=maxMark) if arabic_subject else ''
                #     sheet2['F'+] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words(arabic_subject[0][0]) if arabic_subject else maxMark

                # اللغة الانجليزية 
                english_subject = [value for key ,value in group_item['subject_sums'].items() if 'جليزية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    english_subject_maxMark = 200
                    sheet2[c['english_language']] = english_subject_maxMark / 2
                    sheet2['F'+c['english_language'][1:]] = english_subject_maxMark
                else:
                    english_subject_maxMark = 100
                    sheet2[c['english_language']] = english_subject_maxMark / 2
                    sheet2['F'+c['english_language'][1:]] = english_subject_maxMark
                sheet2['H'+c['english_language'][1:]] = english_subject[0][0] if english_subject and len(english_subject[0]) != 0 else ''
                sheet2['I'+c['english_language'][1:]] = english_subject[0][1] if term2 and english_subject and len(english_subject[0]) != 0 else ''
                sheet2['K'+c['english_language'][1:]] = (english_subject[0][0] + english_subject[0][1])/2 if term2 and english_subject and len(english_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['J20'] = score_in_words(((english_subject[0][0] + english_subject[0][1])/2),max_mark=maxMark) if english_subject else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words((english_subject[0][0] + english_subject[0][1])/2) if english_subject else maxMark
                # else:
                #     sheet2['J20'] = score_in_words(english_subject[0][0],max_mark=maxMark) if english_subject else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words(english_subject[0][0]) if english_subject else maxMark

                # الرياضيات 
                math_subject = [value for key ,value in group_item['subject_sums'].items() if 'رياضيات' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    math_subject_maxMark = 200
                    sheet2[c['mathematics']] = math_subject_maxMark / 2
                    sheet2['F'+c['mathematics'][1:]] = math_subject_maxMark
                else:
                    math_subject_maxMark = 100
                    sheet2[c['mathematics']] = math_subject_maxMark / 2
                    sheet2['F'+c['mathematics'][1:]] = math_subject_maxMark
                sheet2['H'+c['mathematics'][1:]] = math_subject[0][0] if math_subject and len(math_subject[0]) != 0 else ''
                sheet2['I'+c['mathematics'][1:]] = math_subject[0][1] if term2 and math_subject and len(math_subject[0]) != 0 else ''
                sheet2['K'+c['mathematics'][1:]] = (math_subject[0][0] + math_subject[0][1])/2 if term2 and math_subject and len(math_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['J21'] = score_in_words(((math_subject[0][0] + math_subject[0][1])/2),max_mark=maxMark) if math_subject else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words((math_subject[0][0] + math_subject[0][1])/2) if math_subject else maxMark
                # else:
                #     sheet2['J21'] = score_in_words(math_subject[0][0],max_mark=maxMark) if math_subject else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words(math_subject[0][0]) if math_subject else maxMark

                # التربية الاجتماعية و الوطنية 
                social_subjects = [value for key ,value in group_item['subject_sums'].items() if 'اجتماعية و الوطنية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    social_subjects_maxMark = 200
                    sheet2[c['social_studies']] = social_subjects_maxMark / 2
                    sheet2['F'+c['social_studies'][1:]] = social_subjects_maxMark
                    sheet2['H'+c['social_studies'][1:]] = int(social_subjects[0][0]*(2/3)) if social_subjects and len(social_subjects[0]) != 0 else ''
                    sheet2['I'+c['social_studies'][1:]] = int(social_subjects[0][1]*(2/3)) if term2 and social_subjects and len(social_subjects[0]) != 0 else ''
                elif 'سادس' in group_item['student_grade_name'] or 'سابع' in group_item['student_grade_name']:
                    social_subjects_maxMark = 100                
                    sheet2[c['social_studies']] = social_subjects_maxMark / 2
                    sheet2['F'+c['social_studies'][1:]] = social_subjects_maxMark
                    sheet2['H'+c['social_studies'][1:]] = int(social_subjects[0][0]/3) if social_subjects and len(social_subjects[0]) != 0 else ''
                    sheet2['I'+c['social_studies'][1:]] = int(social_subjects[0][1]*(2/3)) if term2 and social_subjects and len(social_subjects[0]) != 0 else ''
                else:
                    social_subjects_maxMark = 100
                    sheet2[c['social_studies']] = social_subjects_maxMark / 2
                    sheet2['F'+c['social_studies'][1:]] = social_subjects_maxMark
                    sheet2['H'+c['social_studies'][1:]] = int(social_subjects[0][0]) if social_subjects and len(social_subjects[0]) != 0 else ''
                    sheet2['I'+c['social_studies'][1:]] = int(social_subjects[0][1]*(2/3)) if term2 and social_subjects and len(social_subjects[0]) != 0 else ''
                    
                sheet2['K'+c['social_studies'][1:]] = (social_subjects[0][0] + social_subjects[0][1])/2 if term2 and social_subjects and len(social_subjects[0]) != 0 else ''
                # if term2:
                #     sheet2['J22'] = score_in_words(int(((social_subjects[0][0] + social_subjects[0][1])/2)*(2/3)),max_mark=maxMark) if social_subjects else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words((social_subjects[0][0] + social_subjects[0][1])/2) if social_subjects else maxMark
                # else:
                #     sheet2['J22'] = score_in_words(int(social_subjects[0][0]*(2/3)),max_mark=maxMark) if social_subjects else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words(social_subjects[0][0]) if social_subjects else maxMark

                # العلوم
                science_subjects = [value for key ,value in group_item['subject_sums'].items() if 'العلوم' in key]
                if 'ثامن' in group_item['student_grade_name'] :
                    science_subjects_maxMark = 200
                    sheet2[c['science']] = science_subjects_maxMark / 2
                    sheet2['F'+c['science'][1:]] = science_subjects_maxMark
                elif 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    science_subjects_maxMark = 400
                    sheet2[c['science']] = science_subjects_maxMark / 2
                    sheet2['F'+c['science'][1:]] = science_subjects_maxMark
                else:
                    science_subjects_maxMark = 100
                    sheet2[c['science']] = science_subjects_maxMark / 2
                    sheet2['F'+c['science'][1:]] = science_subjects_maxMark
                sheet2['H'+c['science'][1:]] = science_subjects[0][0] if science_subjects and len(science_subjects[0]) != 0 else ''
                sheet2['I'+c['science'][1:]] = science_subjects[0][1] if term2 and science_subjects and len(science_subjects[0]) != 0 else ''
                sheet2['K'+c['science'][1:]] = (science_subjects[0][0] + science_subjects[0][1])/2 if term2 and science_subjects and len(science_subjects[0]) != 0 else ''
                # if term2:
                #     sheet2['J23'] = score_in_words( ((science_subjects[0][0] + science_subjects[0][1])/2),max_mark=maxMark) if  science_subjects else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words((science_subjects[0][0] + science_subjects[0][1])/2) if science_subjects else maxMark
                # else:
                #     sheet2['J23'] = score_in_words( science_subjects[0][0],max_mark=maxMark) if  science_subjects else ''
                #     sheet2['F'+c[''][1:]] = maxMark / 2
                #     sheet2[] = convert_avarage_to_words(science_subjects[0][0]) if science_subjects else maxMark

                # التربية الفنية والموسيقية
                art_subject = [value for key ,value in group_item['subject_sums'].items() if 'الفنية والموس' in key]
                sheet2[c['visual_arts']] = 50 if art_subject and len(art_subject[0]) != 0 else ''
                sheet2['F'+c['visual_arts'][1:]] = 100 if art_subject and len(art_subject[0]) != 0 else ''
                sheet2['H'+c['visual_arts'][1:]] = art_subject[0][0] if art_subject and len(art_subject[0]) != 0 else ''
                sheet2['I'+c['visual_arts'][1:]] = art_subject[0][1] if term2 and art_subject and len(art_subject[0]) != 0 else ''
                sheet2['K'+c['visual_arts'][1:]] = (art_subject[0][0] + art_subject[0][1])/2 if term2 and art_subject and len(art_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G24'] = convert_avarage_to_words((art_subject[0][0] + art_subject[0][1])/2) if art_subject else ''
                #     sheet2['J24'] = score_in_words(((art_subject[0][0] + art_subject[0][1])/2) ) if art_subject else ''
                # else:
                #     sheet2['G24'] = convert_avarage_to_words(art_subject[0][0]) if art_subject else ''
                #     sheet2['J24'] = score_in_words(art_subject[0][0] ) if art_subject else ''

                # التربية الرياضية
                sport_subject = [value for key ,value in group_item['subject_sums'].items() if 'رياضية' in key]
                sheet2[c['physical_education']] = 50 if sport_subject and len(sport_subject[0]) != 0 else ''
                sheet2['F'+c['physical_education'][1:]] = 100 if sport_subject and len(sport_subject[0]) != 0 else ''
                sheet2['H'+c['physical_education'][1:]] = sport_subject[0][0] if sport_subject and len(sport_subject[0]) != 0 else ''
                sheet2['I'+c['physical_education'][1:]] = sport_subject[0][1] if term2 and sport_subject and len(sport_subject[0]) != 0 else ''
                sheet2['K'+c['physical_education'][1:]] = (sport_subject[0][0] + sport_subject[0][1])/2 if term2 and sport_subject and len(sport_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G25'] = convert_avarage_to_words((sport_subject[0][0] + sport_subject[0][1])/2) if sport_subject else ''
                #     sheet2['J25'] = score_in_words(((sport_subject[0][0] + sport_subject[0][1])/2) ) if sport_subject else ''
                # else:
                #     sheet2['G25'] = convert_avarage_to_words(sport_subject[0][0]) if sport_subject else ''
                #     sheet2['J25'] = score_in_words(sport_subject[0][0] ) if sport_subject else ''

                # التربية المهنية 
                vocational_subject = [value for key ,value in group_item['subject_sums'].items() if 'مهنية' in key]
                sheet2[c['vocational_education']] = 50 if vocational_subject and len(vocational_subject[0]) != 0 else ''
                sheet2['F'+c['vocational_education'][1:]] = 100 if vocational_subject and len(vocational_subject[0]) != 0 else ''
                sheet2['H'+c['vocational_education'][1:]] = vocational_subject[0][0] if vocational_subject and len(vocational_subject[0]) != 0 else ''
                sheet2['I'+c['vocational_education'][1:]] = vocational_subject[0][1] if term2 and vocational_subject and len(vocational_subject[0]) != 0 else ''
                sheet2['K'+c['vocational_education'][1:]] = (vocational_subject[0][0] + vocational_subject[0][1])/2 if term2 and vocational_subject and len(vocational_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G26'] = convert_avarage_to_words((vocational_subject[0][0] + vocational_subject[0][1])/2) if vocational_subject else ''
                #     sheet2['J26'] = score_in_words(((vocational_subject[0][0] + vocational_subject[0][1])/2) ) if vocational_subject else ''
                # else:
                #     sheet2['G26'] = convert_avarage_to_words(vocational_subject[0][0]) if vocational_subject else ''
                #     sheet2['J26'] = score_in_words(vocational_subject[0][0] ) if vocational_subject else ''

                # الحاسوب
                computer_subject = [value for key ,value in group_item['subject_sums'].items() if 'حاسوب' in key]
                sheet2[c['computer']] = 50 if computer_subject and len(computer_subject[0]) != 0 else ''
                sheet2['F'+c['computer'][1:]] = 100 if computer_subject and len(computer_subject[0]) != 0 else ''
                sheet2['H'+c['computer'][1:]] = computer_subject[0][0] if computer_subject and len(computer_subject[0]) != 0 else ''
                sheet2['I'+c['computer'][1:]] = computer_subject[0][1] if term2 and computer_subject and len(computer_subject[0]) != 0 else ''
                sheet2['K'+c['computer'][1:]] = (computer_subject[0][0] + computer_subject[0][1])/2 if term2 and computer_subject and len(computer_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G27'] = convert_avarage_to_words((computer_subject[0][0] + computer_subject[0][1])/2) if computer_subject else ''
                #     sheet2['J27'] = score_in_words(((computer_subject[0][0] + computer_subject[0][1])/2) ) if computer_subject else ''
                # else:
                #     sheet2['G27'] = convert_avarage_to_words(computer_subject[0][0]) if computer_subject else ''
                #     sheet2['J27'] = score_in_words(computer_subject[0][0] ) if computer_subject else ''

                # الثقافة المالية
                financial_subject = [value for key ,value in group_item['subject_sums'].items() if 'مالية' in key]
                sheet2[c['financial_culture']] = 50 if financial_subject and len(financial_subject[0]) != 0 else ''
                sheet2['F'+c['financial_culture'][1:]] = 100 if financial_subject and len(financial_subject[0]) != 0 else ''
                sheet2['H'+c['financial_culture'][1:]] = financial_subject[0][0] if financial_subject and len(financial_subject[0]) != 0 else ''
                sheet2['I'+c['financial_culture'][1:]] = financial_subject[0][1] if term2 and financial_subject and len(financial_subject[0]) != 0 else ''
                sheet2['K'+c['financial_culture'][1:]] = (financial_subject[0][0] + financial_subject[0][1])/2 if term2 and financial_subject and len(financial_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G28'] = convert_avarage_to_words((financial_subject[0][0] + financial_subject[0][1])/2) if financial_subject else ''
                #     sheet2['J28'] = score_in_words(((financial_subject[0][0] + financial_subject[0][1])/2) ) if financial_subject else ''
                # else:
                #     sheet2['G28'] = convert_avarage_to_words(financial_subject[0][0]) if financial_subject else ''
                #     sheet2['J28'] = score_in_words(financial_subject[0][0] ) if financial_subject else ''

                # اللغة الفرنسية 
                franch_subject = [value for key ,value in group_item['subject_sums'].items() if 'فرنسية' in key]
                sheet2[c['french_language']] = 50 if franch_subject and len(franch_subject[0]) != 0 else ''
                sheet2['F'+c['french_language'][1:]] = 100 if franch_subject and len(franch_subject[0]) != 0 else ''
                sheet2['H'+c['french_language'][1:]] = franch_subject[0][0] if franch_subject and len(franch_subject[0]) != 0 else ''
                sheet2['I'+c['french_language'][1:]] = franch_subject[0][1] if term2 and franch_subject and len(franch_subject[0]) != 0 else ''
                sheet2['K'+c['french_language'][1:]] = (franch_subject[0][0] + franch_subject[0][1])/2 if term2 and franch_subject and len(franch_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G29'] = convert_avarage_to_words((franch_subject[0][0] + franch_subject[0][1])/2) if franch_subject else ''
                #     sheet2['J29'] = score_in_words(((franch_subject[0][0] + franch_subject[0][1])/2) ) if franch_subject else ''
                # else:
                #     sheet2['G29'] = convert_avarage_to_words(franch_subject[0][0]) if franch_subject else ''
                #     sheet2['J29'] = score_in_words(franch_subject[0][0] ) if franch_subject else ''

                # الدين المسيحي
                christian_subject = [value for key ,value in group_item['subject_sums'].items() if 'الدين المسيحي' in key]
                sheet2[c['christian_religion']] = 50 if christian_subject and len(christian_subject[0]) != 0 else ''
                sheet2['F'+c['christian_religion'][1:]] = 100 if christian_subject and len(christian_subject[0]) != 0 else ''
                sheet2['H'+c['christian_religion'][1:]] = christian_subject[0][0] if christian_subject and len(christian_subject[0]) != 0 else ''
                sheet2['I'+c['christian_religion'][1:]] = christian_subject[0][1] if term2 and christian_subject and len(christian_subject[0]) != 0 else ''
                sheet2['K'+c['christian_religion'][1:]] = (christian_subject[0][0] + christian_subject[0][1])/2 if term2 and christian_subject and len(christian_subject[0]) != 0 else ''
                # if term2:
                #     sheet2['G30'] = convert_avarage_to_words((christian_subject[0][0] + christian_subject[0][1])/2) if christian_subject else ''
                #     sheet2['J30'] = score_in_words(((christian_subject[0][0] + christian_subject[0][1])/2) ) if christian_subject else ''
                # else:
                #     sheet2['G30'] = convert_avarage_to_words(christian_subject[0][0]) if christian_subject else ''
                #     sheet2['J30'] = score_in_words(christian_subject[0][0] ) if christian_subject else ''

                counter = 0
                for subject_name  ,S1_S2 in group_item['subject_sums'].items():
                    average = (S1_S2[0]+S1_S2[1])/2
                    print( subject_name, S1_S2)
                    if 'سلامي' in subject_name and average < islam_subject_maxMark / 2 : 
                        counter+=1
                    elif "عربية"  in subject_name and average < arabic_subject_maxMark / 2 : 
                        counter+=1
                    elif "نجليزي"  in subject_name and average < english_subject_maxMark / 2 : 
                        counter+=1
                    elif "رياضيات"  in subject_name and average < math_subject_maxMark / 2 : 
                        counter+=1
                    elif "جتماعية"  in subject_name and average < social_subjects_maxMark / 2 : 
                        counter+=1
                    elif "علوم"  in subject_name and average < science_subjects_maxMark / 2 : 
                        counter+=1
                    elif  average < 50: 
                        counter+=1
                    # طريقة طباعة الرقم صحيح اذا كان بدون اعشار 
                    # print(subject_name , int((S1_S2[0]+S1_S2[1])/2) if str((S1_S2[0]+S1_S2[1])/2).split('.')[1] == '0' else (S1_S2[0]+S1_S2[1])/2 )
                    
                # print(counter)
                if counter > 4 : 
                    print('يبقى في صفه')
                    result = 2
                elif counter == 0 :
                    print("ناجح")
                    result = 0
                else :     
                    print('مكمل')
                    result = 1
                
                if term2 :
                    # المعدل المئوي بالرقام 
                    sheet2[c['average']]= group_item['t1+t2+year_avarage'][2]
                    # #بالحروف
                    # sheet2['e32']= convert_avarage_to_words(group_item['t1+t2+year_avarage'][2]) if group_item else ''
                    # #ترتيب الطالب على الصف 
                    # sheet2['j32']= counter-1

                    #النتيجة 
                    sheet2[result_cell_positions[result]]= '✓'
                else:
                    
                    #المعدل المئوي بالرقام 
                    sheet2[c['average']]= group_item['t1+t2+year_avarage'][term]
                    # #بالحروف
                    # sheet2['e32']= convert_avarage_to_words(group_item['t1+t2+year_avarage'][0]) if group_item else ''
                    # #ترتيب الطالب على الصف 
                    # sheet2['j32']= counter-1
                    # #النتيجة 
                    if result == 2 : # اذا كان مكمل في صفه الفصل الاول خليها اله بس راسب لانه بجوز الفصل الثاني يتحسن 
                        sheet2[result_cell_positions[1]]= '✓'
                    else:    
                        sheet2[result_cell_positions[result]]= '✓'
                
                # #عدد ايام غياب الطالب 
                # sheet2['c35']= ''
                # #عدد ايام الدوام الرسمي الكامل 
                sheet2[c['school_days']]= ''
                #اسم و توقيع مربي الصف 
                sheet2[c['class_teacher_name']]= grouped_list['teacher_incharge_name']
                # #التاريخ
                # sheet2['b36']= ''
                #اسم و توقيع مدير المدرسة
                sheet2[c['principal_name']]= grouped_list['principle_name'] 
            template_file.remove(template_file['Sheet1'])
            template_file.save(outdir+group[0]['student_class_name_letter']+'.xlsx')

def create_coloured_certs_ods(grouped_list , term2=False ,template='./templet_files/حشو  شهادات الكرتون.ods' , outdir='./send_folder/'):
    
    colored_cert_cells_position_context = {
                                        'student_name': 'E11',
                                        'student_name2': 'B44',
                                        'class_section': 'E12',
                                        'class_name':'B45',
                                        'nationality': 'D13',
                                        'national_id': 'E14',
                                        'birthplace_date': 'E15',
                                        'religion': 'D16',
                                        'student_address': 'E17',
                                        'school_name': 'E18',
                                        'school_address': 'D19',
                                        'directorate': 'D20',
                                        'school_bridge': 'D21',
                                        'supervising_authority': 'E22',
                                        'school_phone_number': 'E23',
                                        'school_natioanl_id': 'E24',
                                        'academic_year_1': 'F42',
                                        'academic_year_2': 'G42',
                                        'class': 'B45',
                                        'islamic_education': 'E50',
                                        'arabic_language': 'E51',
                                        'english_language': 'E52',
                                        'mathematics': 'E53',
                                        'social_studies': 'E54',
                                        'science': 'E55',
                                        'visual_arts': 'E56',
                                        'physical_education': 'E57',
                                        'vocational_education': 'E58',
                                        'computer': 'E59',
                                        'financial_culture': 'E60',
                                        'french_language': 'E61',
                                        'christian_religion': 'E62',
                                        'average': 'B67',
                                        'school_days': 'I64',
                                        'class_teacher_name': 'B69',
                                        'principal_name': 'L69',
                                        'semester_1_student_absent': 'H66',
                                        'semester_2_student_absent': 'L66',
                                        }
    
    c = colored_cert_cells_position_context
    result_cell_positions = ['B64','B65','B66']
    
    statistic_data = grouped_list['students_info']
    assessments_data_groups = grouped_list['assessments_data']
    term = 1 if term2 else 0
    
    
    for group in assessments_data_groups:
        if not any("عشر" in i['student_grade_name'] for i in group) : 
            template_file = ezodf.opendoc(template)
            sheet1 = template_file.sheets[0]
            filler_sheet = template_file.sheets[1]

            sheet2 =filler_sheet
            
            names_averages =  sort_dictionary_list_based_on(group,item_in_list=term)

            group = sort_dictionary_list_based_on(group ,simple=False,item_in_list=term)

            for row_number, dataFrame in enumerate(names_averages, start=4):
                sheet1[row_number, 1].set_value( dataFrame[1])
                sheet1[row_number, 3].set_value( dataFrame[0])
                
            for sheet_number , group_item in enumerate(group , start=2):

                # sheet2 = template_file.copy_worksheet(template_file.worksheets[1])
                # sheet2.title = str(counter)
                # counter += 1
                # sheet2.sheet_view.rightToLeft = True    
                # sheet2.sheet_view.rightToLeft = True   
                
                wanted_id = group_item['student_id']
                student_statistic_info = [i for i in statistic_data if i['student_id'] == wanted_id ][0]

                # img = openpyxl.drawing.image.Image(image)
                # img.anchor = 'e2'
                # sheet2.add_image(img)

                # print(sheet2)
                sheet2[c['student_name']].set_value( group_item['student__full_name'])
                # مكان و تاريخ الولادة
                sheet2[c['birthplace_date']].set_value( str(group_item['student_birth_place']) + ' ' + str(group_item['student_birth_date']))
                # الديانة
                sheet2[c['religion']].set_value( student_statistic_info['religion'])
                # عنوان الطالب
                sheet2[c['student_address']].set_value( student_statistic_info['address'])
                #الرقم الوطني
                sheet2[c['national_id']].set_value( group_item['student_nat_id'])
                #الجنسية
                sheet2[c['nationality']].set_value( group_item['student_nat'])
                #الصف و الشعبة 
                sheet2[c['class_section']].set_value( group_item['student_class_name_letter'])
                # صف الطالب
                sheet2[c['class_name']].set_value( group_item['student_grade_name'].replace('الصف' ,''))
                # اسم المدرسة
                sheet2[c['school_name']].set_value( group_item['student_school_name'])
                # عنوان المدرسة
                sheet2[c['school_address']].set_value( grouped_list['school_address'])
                # مديرية المدرسة 
                sheet2[c['directorate']].set_value( grouped_list['school_directorate'])
                # لواء المدرسة
                sheet2[c['school_bridge']].set_value( grouped_list['school_bridge'])
                #المنطقة التعليمية  او السلطة المشرفة
                sheet2[c['supervising_authority']].set_value( 'وزارة التربية و التعليم')
                # هاتف المدرسة
                sheet2[c['school_phone_number']].set_value( grouped_list['school_phone_number'])
                # رقم المدرسة الوطني
                sheet2[c['school_natioanl_id']].set_value( grouped_list['school_national_id'])
                # العام الدراسي الاول 
                sheet2[c['academic_year_1']].set_value( grouped_list['academic_year_1'])
                # العام الدراسي الثاني 
                sheet2[c['academic_year_2']].set_value( grouped_list['academic_year_2'])
                # الاسم على الوجه الثاني
                sheet2[c['student_name2']].set_value( group_item['student__full_name'])
                # غياب الفصل الاول 
                sheet2[c['semester_1_student_absent']].set_value( '')
                # غياب الفصل الثاني
                sheet2[c['semester_2_student_absent']].set_value( '')
                
                # put the subjects cells inder here 
                i ='c,d,e,g,j,f'.split(',')
                r = range(18,32)


                # التربية الاسلامية
                islam_subject = [value for key ,value in group_item['subject_sums'].items() if 'سلامية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    islam_subject_maxMark = 200
                    sheet2[c['islamic_education']].set_value( islam_subject_maxMark /2)
                    sheet2['F'+c['islamic_education'][1:]].set_value( islam_subject_maxMark)
                else:
                    islam_subject_maxMark = 100
                    sheet2[c['islamic_education']].set_value( islam_subject_maxMark /2)
                    sheet2['F'+c['islamic_education'][1:]].set_value( islam_subject_maxMark)
                sheet2['H'+c['islamic_education'][1:]].set_value( islam_subject[0][0] if islam_subject and len(islam_subject[0]) != 0 else '')
                sheet2['I'+c['islamic_education'][1:]].set_value( islam_subject[0][1] if term2 and islam_subject and len(islam_subject[0]) != 0 else '')
                sheet2['K'+c['islamic_education'][1:]].set_value( (islam_subject[0][0] + islam_subject[0][1])/2 if term2 and islam_subject and len(islam_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G'+c['islamic_education']].set_value( convert_avarage_to_words((islam_subject[0][0] + islam_subject[0][1])/2) if islam_subject else '')
                #     sheet2['J'+c['islamic_education']].set_value( score_in_words(((islam_subject[0][0] + islam_subject[0][1])/2),max_mark=maxMark) if islam_subject else '')
                # else:
                #     sheet2['G'+c['islamic_education']].set_value( convert_avarage_to_words(islam_subject[0][0]) if islam_subject else '')
                #     sheet2['J'+c['islamic_education']].set_value( score_in_words(islam_subject[0][0],max_mark=maxMark) if islam_subject else '')

                # اللغة العربية
                arabic_subject = [value for key ,value in group_item['subject_sums'].items() if 'عربية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    arabic_subject_maxMark = 300
                    sheet2[c['arabic_language']].set_value( arabic_subject_maxMark / 2)
                    sheet2['F'+c['arabic_language'][1:]].set_value( arabic_subject_maxMark)
                else:
                    arabic_subject_maxMark = 100
                    sheet2[c['arabic_language']].set_value( arabic_subject_maxMark / 2)
                    sheet2['F'+c['arabic_language'][1:]].set_value( arabic_subject_maxMark)
                sheet2['H' + c['arabic_language'][1:]].set_value( arabic_subject[0][0] if arabic_subject and len(arabic_subject[0]) != 0 else '')
                sheet2['I' + c['arabic_language'][1:]].set_value( arabic_subject[0][1] if term2 and arabic_subject and len(arabic_subject[0]) != 0 else '')
                sheet2['K' + c['arabic_language'][1:]].set_value( (arabic_subject[0][0] + arabic_subject[0][1])/2 if term2 and arabic_subject and len(arabic_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['J19'].set_value( score_in_words(((arabic_subject[0][0] + arabic_subject[0][1])/2),max_mark=maxMark) if arabic_subject else '')
                #     sheet2['F'+].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words((arabic_subject[0][0] + arabic_subject[0][1])/2) if arabic_subject else maxMark)
                # else:
                #     sheet2['J19'].set_value( score_in_words(arabic_subject[0][0],max_mark=maxMark) if arabic_subject else '')
                #     sheet2['F'+].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words(arabic_subject[0][0]) if arabic_subject else maxMark)

                # اللغة الانجليزية 
                english_subject = [value for key ,value in group_item['subject_sums'].items() if 'جليزية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    english_subject_maxMark = 200
                    sheet2[c['english_language']].set_value( english_subject_maxMark / 2)
                    sheet2['F'+c['english_language'][1:]].set_value( english_subject_maxMark)
                else:
                    english_subject_maxMark = 100
                    sheet2[c['english_language']].set_value( english_subject_maxMark / 2)
                    sheet2['F'+c['english_language'][1:]].set_value( english_subject_maxMark)
                sheet2['H'+c['english_language'][1:]].set_value( english_subject[0][0] if english_subject and len(english_subject[0]) != 0 else '')
                sheet2['I'+c['english_language'][1:]].set_value( english_subject[0][1] if term2 and english_subject and len(english_subject[0]) != 0 else '')
                sheet2['K'+c['english_language'][1:]].set_value( (english_subject[0][0] + english_subject[0][1])/2 if term2 and english_subject and len(english_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['J20'].set_value( score_in_words(((english_subject[0][0] + english_subject[0][1])/2),max_mark=maxMark) if english_subject else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words((english_subject[0][0] + english_subject[0][1])/2) if english_subject else maxMark)
                # else:
                #     sheet2['J20'].set_value( score_in_words(english_subject[0][0],max_mark=maxMark) if english_subject else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words(english_subject[0][0]) if english_subject else maxMark)

                # الرياضيات 
                math_subject = [value for key ,value in group_item['subject_sums'].items() if 'رياضيات' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    math_subject_maxMark = 200
                    sheet2[c['mathematics']].set_value( math_subject_maxMark / 2)
                    sheet2['F'+c['mathematics'][1:]].set_value( math_subject_maxMark)
                else:
                    math_subject_maxMark = 100
                    sheet2[c['mathematics']].set_value( math_subject_maxMark / 2)
                    sheet2['F'+c['mathematics'][1:]].set_value( math_subject_maxMark)
                sheet2['H'+c['mathematics'][1:]].set_value( math_subject[0][0] if math_subject and len(math_subject[0]) != 0 else '')
                sheet2['I'+c['mathematics'][1:]].set_value( math_subject[0][1] if term2 and math_subject and len(math_subject[0]) != 0 else '')
                sheet2['K'+c['mathematics'][1:]].set_value( (math_subject[0][0] + math_subject[0][1])/2 if term2 and math_subject and len(math_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['J21'].set_value( score_in_words(((math_subject[0][0] + math_subject[0][1])/2),max_mark=maxMark) if math_subject else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words((math_subject[0][0] + math_subject[0][1])/2) if math_subject else maxMark)
                # else:
                #     sheet2['J21'].set_value( score_in_words(math_subject[0][0],max_mark=maxMark) if math_subject else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words(math_subject[0][0]) if math_subject else maxMark)

                # التربية الاجتماعية و الوطنية 
                social_subjects = [value for key ,value in group_item['subject_sums'].items() if 'اجتماعية و الوطنية' in key]
                if 'ثامن' in group_item['student_grade_name'] or 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    social_subjects_maxMark = 200
                    sheet2[c['social_studies']].set_value( social_subjects_maxMark / 2)
                    sheet2['F'+c['social_studies'][1:]].set_value( social_subjects_maxMark)
                    sheet2['H'+c['social_studies'][1:]].set_value( int(social_subjects[0][0]*(2/3)) if social_subjects and len(social_subjects[0]) != 0 else '')
                    sheet2['I'+c['social_studies'][1:]].set_value( int(social_subjects[0][1]*(2/3)) if term2 and social_subjects and len(social_subjects[0]) != 0 else '')
                    sheet2['K'+c['social_studies'][1:]].set_value( round((((social_subjects[0][0] + social_subjects[0][1])/2)*(2/3)),1) if term2 and social_subjects and len(social_subjects[0]) != 0 else '')
                elif 'سادس' in group_item['student_grade_name'] or 'سابع' in group_item['student_grade_name']:
                    social_subjects_maxMark = 100                
                    sheet2[c['social_studies']].set_value( social_subjects_maxMark / 2)
                    sheet2['F'+c['social_studies'][1:]].set_value( social_subjects_maxMark)
                    sheet2['H'+c['social_studies'][1:]].set_value( int(social_subjects[0][0]/3) if social_subjects and len(social_subjects[0]) != 0 else '')
                    sheet2['I'+c['social_studies'][1:]].set_value( int(social_subjects[0][1]/3)) if term2 and social_subjects and len(social_subjects[0]) != 0 else ''
                    sheet2['K'+c['social_studies'][1:]].set_value( round(((social_subjects[0][0] + social_subjects[0][1])/2)/3 , 1) if term2 and social_subjects and len(social_subjects[0]) != 0 else '')
                else:
                    social_subjects_maxMark = 100
                    sheet2[c['social_studies']].set_value( social_subjects_maxMark / 2)
                    sheet2['F'+c['social_studies'][1:]].set_value( social_subjects_maxMark)
                    sheet2['H'+c['social_studies'][1:]].set_value( int(social_subjects[0][0]) if social_subjects and len(social_subjects[0]) != 0 else '')
                    sheet2['I'+c['social_studies'][1:]].set_value( int(social_subjects[0][1]) if term2 and social_subjects and len(social_subjects[0]) != 0 else '')
                    sheet2['K'+c['social_studies'][1:]].set_value( (social_subjects[0][0] + social_subjects[0][1])/2 if term2 and social_subjects and len(social_subjects[0]) != 0 else '')
                    
                # if term2:
                #     sheet2['J22'].set_value( score_in_words(int(((social_subjects[0][0] + social_subjects[0][1])/2)*(2/3)),max_mark=maxMark) if social_subjects else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words((social_subjects[0][0] + social_subjects[0][1])/2) if social_subjects else maxMark)
                # else:
                #     sheet2['J22'].set_value( score_in_words(int(social_subjects[0][0]*(2/3)),max_mark=maxMark) if social_subjects else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words(social_subjects[0][0]) if social_subjects else maxMark)

                # العلوم
                science_subjects = [value for key ,value in group_item['subject_sums'].items() if 'العلوم' in key]
                if 'ثامن' in group_item['student_grade_name'] :
                    science_subjects_maxMark = 200
                    sheet2[c['science']].set_value( science_subjects_maxMark / 2)
                    sheet2['F'+c['science'][1:]].set_value( science_subjects_maxMark)
                elif 'تاسع' in group_item['student_grade_name'] or 'عاشر' in group_item['student_grade_name']:
                    science_subjects_maxMark = 400
                    sheet2[c['science']].set_value( science_subjects_maxMark / 2)
                    sheet2['F'+c['science'][1:]].set_value( science_subjects_maxMark)
                else:
                    science_subjects_maxMark = 100
                    sheet2[c['science']].set_value( science_subjects_maxMark / 2)
                    sheet2['F'+c['science'][1:]].set_value( science_subjects_maxMark)
                sheet2['H'+c['science'][1:]].set_value( science_subjects[0][0] if science_subjects and len(science_subjects[0]) != 0 else '')
                sheet2['I'+c['science'][1:]].set_value( science_subjects[0][1] if term2 and science_subjects and len(science_subjects[0]) != 0 else '')
                sheet2['K'+c['science'][1:]].set_value( (science_subjects[0][0] + science_subjects[0][1])/2 if term2 and science_subjects and len(science_subjects[0]) != 0 else '')
                # if term2:
                #     sheet2['J23'].set_value( score_in_words( ((science_subjects[0][0] + science_subjects[0][1])/2),max_mark=maxMark) if  science_subjects else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words((science_subjects[0][0] + science_subjects[0][1])/2) if science_subjects else maxMark)
                # else:
                #     sheet2['J23'].set_value( score_in_words( science_subjects[0][0],max_mark=maxMark) if  science_subjects else '')
                #     sheet2['F'+c[''][1:]].set_value( maxMark / 2)
                #     sheet2[].set_value( convert_avarage_to_words(science_subjects[0][0]) if science_subjects else maxMark)

                # التربية الفنية والموسيقية
                art_subject = [value for key ,value in group_item['subject_sums'].items() if 'الفنية والموس' in key]
                sheet2[c['visual_arts']].set_value( 50 if art_subject and len(art_subject[0]) != 0 else '')
                sheet2['F'+c['visual_arts'][1:]].set_value( 100 if art_subject and len(art_subject[0]) != 0 else '')
                sheet2['H'+c['visual_arts'][1:]].set_value( art_subject[0][0] if art_subject and len(art_subject[0]) != 0 else '')
                sheet2['I'+c['visual_arts'][1:]].set_value( art_subject[0][1] if term2 and art_subject and len(art_subject[0]) != 0 else '')
                sheet2['K'+c['visual_arts'][1:]].set_value( (art_subject[0][0] + art_subject[0][1])/2 if term2 and art_subject and len(art_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G24'].set_value( convert_avarage_to_words((art_subject[0][0] + art_subject[0][1])/2) if art_subject else '')
                #     sheet2['J24'].set_value( score_in_words(((art_subject[0][0] + art_subject[0][1])/2) ) if art_subject else '')
                # else:
                #     sheet2['G24'].set_value( convert_avarage_to_words(art_subject[0][0]) if art_subject else '')
                #     sheet2['J24'].set_value( score_in_words(art_subject[0][0] ) if art_subject else '')

                # التربية الرياضية
                sport_subject = [value for key ,value in group_item['subject_sums'].items() if 'رياضية' in key]
                sheet2[c['physical_education']].set_value( 50 if sport_subject and len(sport_subject[0]) != 0 else '')
                sheet2['F'+c['physical_education'][1:]].set_value( 100 if sport_subject and len(sport_subject[0]) != 0 else '')
                sheet2['H'+c['physical_education'][1:]].set_value( sport_subject[0][0] if sport_subject and len(sport_subject[0]) != 0 else '')
                sheet2['I'+c['physical_education'][1:]].set_value( sport_subject[0][1] if term2 and sport_subject and len(sport_subject[0]) != 0 else '')
                sheet2['K'+c['physical_education'][1:]].set_value( (sport_subject[0][0] + sport_subject[0][1])/2 if term2 and sport_subject and len(sport_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G25'].set_value( convert_avarage_to_words((sport_subject[0][0] + sport_subject[0][1])/2) if sport_subject else '')
                #     sheet2['J25'].set_value( score_in_words(((sport_subject[0][0] + sport_subject[0][1])/2) ) if sport_subject else '')
                # else:
                #     sheet2['G25'].set_value( convert_avarage_to_words(sport_subject[0][0]) if sport_subject else '')
                #     sheet2['J25'].set_value( score_in_words(sport_subject[0][0] ) if sport_subject else '')

                # التربية المهنية 
                vocational_subject = [value for key ,value in group_item['subject_sums'].items() if 'مهنية' in key]
                sheet2[c['vocational_education']].set_value( 50 if vocational_subject and len(vocational_subject[0]) != 0 else '')
                sheet2['F'+c['vocational_education'][1:]].set_value( 100 if vocational_subject and len(vocational_subject[0]) != 0 else '')
                sheet2['H'+c['vocational_education'][1:]].set_value( vocational_subject[0][0] if vocational_subject and len(vocational_subject[0]) != 0 else '')
                sheet2['I'+c['vocational_education'][1:]].set_value( vocational_subject[0][1] if term2 and vocational_subject and len(vocational_subject[0]) != 0 else '')
                sheet2['K'+c['vocational_education'][1:]].set_value( (vocational_subject[0][0] + vocational_subject[0][1])/2 if term2 and vocational_subject and len(vocational_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G26'].set_value( convert_avarage_to_words((vocational_subject[0][0] + vocational_subject[0][1])/2) if vocational_subject else '')
                #     sheet2['J26'].set_value( score_in_words(((vocational_subject[0][0] + vocational_subject[0][1])/2) ) if vocational_subject else '')
                # else:
                #     sheet2['G26'].set_value( convert_avarage_to_words(vocational_subject[0][0]) if vocational_subject else '')
                #     sheet2['J26'].set_value( score_in_words(vocational_subject[0][0] ) if vocational_subject else '')

                # الحاسوب
                computer_subject = [value for key ,value in group_item['subject_sums'].items() if 'حاسوب' in key]
                sheet2[c['computer']].set_value( 50 if computer_subject and len(computer_subject[0]) != 0 else '')
                sheet2['F'+c['computer'][1:]].set_value( 100 if computer_subject and len(computer_subject[0]) != 0 else '')
                sheet2['H'+c['computer'][1:]].set_value( computer_subject[0][0] if computer_subject and len(computer_subject[0]) != 0 else '')
                sheet2['I'+c['computer'][1:]].set_value( computer_subject[0][1] if term2 and computer_subject and len(computer_subject[0]) != 0 else '')
                sheet2['K'+c['computer'][1:]].set_value( (computer_subject[0][0] + computer_subject[0][1])/2 if term2 and computer_subject and len(computer_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G27'].set_value( convert_avarage_to_words((computer_subject[0][0] + computer_subject[0][1])/2) if computer_subject else '')
                #     sheet2['J27'].set_value( score_in_words(((computer_subject[0][0] + computer_subject[0][1])/2) ) if computer_subject else '')
                # else:
                #     sheet2['G27'].set_value( convert_avarage_to_words(computer_subject[0][0]) if computer_subject else '')
                #     sheet2['J27'].set_value( score_in_words(computer_subject[0][0] ) if computer_subject else '')

                # الثقافة المالية
                financial_subject = [value for key ,value in group_item['subject_sums'].items() if 'مالية' in key]
                sheet2[c['financial_culture']].set_value( 50 if financial_subject and len(financial_subject[0]) != 0 else '')
                sheet2['F'+c['financial_culture'][1:]].set_value( 100 if financial_subject and len(financial_subject[0]) != 0 else '')
                sheet2['H'+c['financial_culture'][1:]].set_value( financial_subject[0][0] if financial_subject and len(financial_subject[0]) != 0 else '')
                sheet2['I'+c['financial_culture'][1:]].set_value( financial_subject[0][1] if term2 and financial_subject and len(financial_subject[0]) != 0 else '')
                sheet2['K'+c['financial_culture'][1:]].set_value( (financial_subject[0][0] + financial_subject[0][1])/2 if term2 and financial_subject and len(financial_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G28'].set_value( convert_avarage_to_words((financial_subject[0][0] + financial_subject[0][1])/2) if financial_subject else '')
                #     sheet2['J28'].set_value( score_in_words(((financial_subject[0][0] + financial_subject[0][1])/2) ) if financial_subject else '')
                # else:
                #     sheet2['G28'].set_value( convert_avarage_to_words(financial_subject[0][0]) if financial_subject else '')
                #     sheet2['J28'].set_value( score_in_words(financial_subject[0][0] ) if financial_subject else '')

                # اللغة الفرنسية 
                franch_subject = [value for key ,value in group_item['subject_sums'].items() if 'فرنسية' in key]
                sheet2[c['french_language']].set_value( 50 if franch_subject and len(franch_subject[0]) != 0 else '')
                sheet2['F'+c['french_language'][1:]].set_value( 100 if franch_subject and len(franch_subject[0]) != 0 else '')
                sheet2['H'+c['french_language'][1:]].set_value( franch_subject[0][0] if franch_subject and len(franch_subject[0]) != 0 else '')
                sheet2['I'+c['french_language'][1:]].set_value( franch_subject[0][1] if term2 and franch_subject and len(franch_subject[0]) != 0 else '')
                sheet2['K'+c['french_language'][1:]].set_value( (franch_subject[0][0] + franch_subject[0][1])/2 if term2 and franch_subject and len(franch_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G29'].set_value( convert_avarage_to_words((franch_subject[0][0] + franch_subject[0][1])/2) if franch_subject else '')
                #     sheet2['J29'].set_value( score_in_words(((franch_subject[0][0] + franch_subject[0][1])/2) ) if franch_subject else '')
                # else:
                #     sheet2['G29'].set_value( convert_avarage_to_words(franch_subject[0][0]) if franch_subject else '')
                #     sheet2['J29'].set_value( score_in_words(franch_subject[0][0] ) if franch_subject else '')

                # الدين المسيحي
                christian_subject = [value for key ,value in group_item['subject_sums'].items() if 'الدين المسيحي' in key]
                sheet2[c['christian_religion']].set_value( 50 if christian_subject and len(christian_subject[0]) != 0 else '')
                sheet2['F'+c['christian_religion'][1:]].set_value( 100 if christian_subject and len(christian_subject[0]) != 0 else '')
                sheet2['H'+c['christian_religion'][1:]].set_value( christian_subject[0][0] if christian_subject and len(christian_subject[0]) != 0 else '')
                sheet2['I'+c['christian_religion'][1:]].set_value( christian_subject[0][1] if term2 and christian_subject and len(christian_subject[0]) != 0 else '')
                sheet2['K'+c['christian_religion'][1:]].set_value( (christian_subject[0][0] + christian_subject[0][1])/2 if term2 and christian_subject and len(christian_subject[0]) != 0 else '')
                # if term2:
                #     sheet2['G30'].set_value( convert_avarage_to_words((christian_subject[0][0] + christian_subject[0][1])/2) if christian_subject else '')
                #     sheet2['J30'].set_value( score_in_words(((christian_subject[0][0] + christian_subject[0][1])/2) ) if christian_subject else '')
                # else:
                #     sheet2['G30'].set_value( convert_avarage_to_words(christian_subject[0][0]) if christian_subject else '')
                #     sheet2['J30'].set_value( score_in_words(christian_subject[0][0] ) if christian_subject else '')

                counter = 0
                for subject_name  ,S1_S2 in group_item['subject_sums'].items():
                    average = (S1_S2[0]+S1_S2[1])/2
                    print( subject_name, S1_S2)
                    if 'سلامي' in subject_name and average < islam_subject_maxMark / 2 : 
                        counter+=1
                    elif "عربية"  in subject_name and average < arabic_subject_maxMark / 2 : 
                        counter+=1
                    elif "نجليزي"  in subject_name and average < english_subject_maxMark / 2 : 
                        counter+=1
                    elif "رياضيات"  in subject_name and average < math_subject_maxMark / 2 : 
                        counter+=1
                    elif "جتماعية"  in subject_name and average < social_subjects_maxMark / 2 : 
                        counter+=1
                    elif "علوم"  in subject_name and average < science_subjects_maxMark / 2 : 
                        counter+=1
                    elif  average < 50: 
                        counter+=1
                    # طريقة طباعة الرقم صحيح اذا كان بدون اعشار 
                    # print(subject_name , int((S1_S2[0]+S1_S2[1])/2) if str((S1_S2[0]+S1_S2[1])/2).split('.')[1] == '0' else (S1_S2[0]+S1_S2[1])/2 )
                    
                # print(counter)
                if counter > 4 : 
                    print('يبقى في صفه')
                    result = 2
                elif counter == 0 :
                    print("ناجح")
                    result = 0
                else :     
                    print('مكمل')
                    result = 1
                
                if term2 :
                    # المعدل المئوي بالرقام 
                    sheet2[c['average']].set_value( group_item['t1+t2+year_avarage'][2])
                    # #بالحروف
                    # sheet2['e32'].set_value( convert_avarage_to_words(group_item['t1+t2+year_avarage'][2]) if group_item else '')
                    # #ترتيب الطالب على الصف 
                    # sheet2['j32'].set_value( counter-1)

                    #النتيجة 
                    sheet2[result_cell_positions[result]].set_value( '✓')
                else:
                    
                    #المعدل المئوي بالرقام 
                    sheet2[c['average']].set_value( group_item['t1+t2+year_avarage'][term])
                    # #بالحروف
                    # sheet2['e32'].set_value( convert_avarage_to_words(group_item['t1+t2+year_avarage'][0]) if group_item else '')
                    # #ترتيب الطالب على الصف 
                    # sheet2['j32'].set_value( counter-1)
                    # #النتيجة 
                    if result == 2 : # اذا كان مكمل في صفه الفصل الاول خليها اله بس راسب لانه بجوز الفصل الثاني يتحسن 
                        sheet2[result_cell_positions[1]].set_value( '✓')
                    else:    
                        sheet2[result_cell_positions[result]].set_value( '✓')
                
                # #عدد ايام غياب الطالب 
                # sheet2['c35'].set_value( '')
                # #عدد ايام الدوام الرسمي الكامل 
                sheet2[c['school_days']].set_value( '')
                
                if grouped_list['teacher_incharge_name'] != '  ':   # اذا لم يكن اسم المعلم فارغ عبي اسم المعلم
                    #اسم و توقيع مربي الصف 
                    sheet2[c['class_teacher_name']].set_value( grouped_list['teacher_incharge_name'])
                # #التاريخ
                # sheet2['b36'].set_value( '')
                #اسم و توقيع مدير المدرسة
                
                sheet2[c['principal_name']].set_value( grouped_list['principle_name'] )
                sheet2 = filler_sheet.copy(newname=str(sheet_number))
                template_file.sheets += sheet2  
            del template_file.sheets[-1]
            # template_file.remove(template_file['Sheet1'])
            template_file.saveas(outdir+group[0]['student_class_name_letter']+'.ods')

def sort_dictionary_list_based_on(dictionary_list, dictionary_key='t1+t2+year_avarage', item_in_list=0, reverse=True, simple=True):
    if simple:
        return [
            (
                0 if len(i.get('subjects_assessments_info', [])) == 0 else i.get(dictionary_key, [''])[item_in_list],
                i.get('student__full_name', '')
            )
            for i in sorted(
                dictionary_list,
                key=lambda x: 0 if len(x.get('subjects_assessments_info', [])) == 0 else x.get(dictionary_key, [''])[item_in_list],
                reverse=reverse
            )
        ]
    else:
        return sorted(
                dictionary_list,
                key=lambda x: 0 if len(x.get('subjects_assessments_info', [])) == 0 else x.get(dictionary_key, [''])[item_in_list],
                reverse=reverse
            )

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
        if fraction == 0:
            return number_in_words.replace(' و', ' و ')
        else:
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
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 900)* 100,1) , round((term_2_avarage / 900)* 100,1) , round((((term_1_avarage+term_2_avarage)/2) / 900)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'سابع' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]/3,1)
                        term_2_avarage +=round(value[1]/3,1)
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 1100)* 100,1) , round((term_2_avarage / 1100)* 100,1) , round((((term_1_avarage+term_2_avarage)/2) / 1100)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'ثامن' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage += round(value[0]*2/3,1)
                        term_2_avarage += round(value[1]*2/3,1)
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 1800)* 100,1) , round((term_2_avarage / 1800)* 100,1) , round((((term_1_avarage+term_2_avarage)/2) / 1800)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'تاسع' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]*2/3,1)
                        term_2_avarage +=round(value[1]*2/3,1)
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 2000)* 100,1) , round((term_2_avarage / 2000)* 100,1) , round((((term_1_avarage+term_2_avarage)/2) / 2000)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            elif 'عاشر' in  item['student_grade_name']:
                for key, value in item['subject_sums'].items():
                    if 'ربية الاجتماعية و الوطنية' in key :
                        # print(key ,round(value[0]*2/3),1)
                        term_1_avarage +=round(value[0]*2/3,1)
                        term_2_avarage +=round(value[1]*2/3,1)
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    elif skip_art_sport :
                        if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                            pass                        
                    else:
                        # print(key , value[0])
                        term_1_avarage += value[0]
                        term_2_avarage += value[1]
                        # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 2000)* 100,1) , round((term_2_avarage / 2000)* 100,1) , round((((term_1_avarage+term_2_avarage)/2) / 2000)* 100,1)
                item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

            else:
                if 'عشر' not in item['student_grade_name']:
                    for key, value in item['subject_sums'].items():
                        if 'ربية الاجتماعية و الوطنية' in key :
                            # print(key ,round(value[0]*2/3),1)
                            term_1_avarage +=round(value[0]*2/3,1)
                            term_2_avarage +=round(value[1]*2/3,1)
                            # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                        elif skip_art_sport :
                            if 'التربية الفنية والموسيقية' in key or 'التربية الرياضية' in key:
                                pass                        
                        else:
                            # print(key , value[0])
                            term_1_avarage += value[0]
                            term_2_avarage += value[1]
                            # year_avarage += round((term_1_avarage + term_2_avarage)/2,1)
                    term_1_avarage ,term_2_avarage ,year_avarage =round((term_1_avarage / 800)* 100,1) , round((term_2_avarage / 800)* 100,1) , round((((term_1_avarage+term_2_avarage)/2) / 800)* 100,1)
                    item['t1+t2+year_avarage'] = [term_1_avarage ,term_2_avarage ,year_avarage ]

def add_subject_sum_dictionary (grouped_dict_list):
    subject_sums = {}
    for group in grouped_dict_list:
        for items in group:
            if len(items['subjects_assessments_info']) > 0 :
                science_sum ,social_sum ,subject_sum ,science_sum_t2 ,social_sum_t2 ,subject_sum_t2 =  [0] * 6
                for i in items['subjects_assessments_info'][0]:
                    if "علوم الأرض" in i['subject_name'] or 'الكيمياء' in i['subject_name'] or 'الحياتية' in i['subject_name'] or 'الفيزياء' in i['subject_name'] or 'العلوم' in i['subject_name']:
                        # compute sum for science subjects
                        science_sum +=  sum(int(i['term1'][key]) for key in i['term1'] if re.compile(r'^assessment\d+$').match(key) and 'max_mark' not in key and i['term1'][key])
                        science_sum_t2 +=  sum(int(i['term2'][key]) for key in i['term2'] if re.compile(r'^assessment\d+$').match(key) and 'max_mark' not in key and i['term2'][key])
                    elif 'التربية الوطنية و المدنية' in i['subject_name'] or 'الجغرافيا' in i['subject_name'] or 'تاريخ' in i['subject_name'] or 'التربية الاجتماعية' in i['subject_name']:
                        # compute sum for social subjects
                        social_sum +=  sum(int(i['term1'][key]) for key in i['term1'] if re.compile(r'^assessment\d+$').match(key) and 'max_mark' not in key and i['term1'][key])
                        social_sum_t2 +=  sum(int(i['term2'][key]) for key in i['term2'] if re.compile(r'^assessment\d+$').match(key) and 'max_mark' not in key and i['term2'][key])
                    else:
                        # compute sum for other subjects
                        subject_sum = sum(int(i['term1'][key]) for key in i['term1'] if re.compile(r'^assessment\d+$').match(key) and 'max_mark' not in key and i['term1'][key])
                        subject_sum_t2 = sum(int(i['term2'][key]) for key in i['term2'] if re.compile(r'^assessment\d+$').match(key) and 'max_mark' not in key and i['term2'][key])
                        # update dictionary with other subject sum
                        subject_sums[i['subject_name']] = [subject_sum,subject_sum_t2]
                    # اضف قيم العلوم و الاجتماعيات مهما كانت قيم المواد
                    subject_sums['العلوم'] = [science_sum ,science_sum_t2]
                    subject_sums['التربية الاجتماعية و الوطنية'] = [social_sum,social_sum_t2]

                    # update dictionary with social subject sum
                # print (items['student__full_name'],items['student_class_name_letter'],subject_sums)
                items['subject_sums'] = subject_sums
                subject_sums={}

def playsound(debug =False):
    # Execute the shell command to play a sine wave sound with frequency 440Hz for 2 seconds
    subprocess.run(['play', '-n', 'synth', '2', 'sin', '440'])
    if debug :
        pdb.set_trace()
    
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
    
def get_students_info_subjectsMarks(username,password,session=None):
    '''
    دالة لاستخراج معلومات و علامات الطلاب لاستخدامها لاحقا في انشاء الجداول و العلامات
    '''
    auth=get_auth(username,password)
    dic_list=[]
    target_student_marks=[]
    school_name = inst_name(session=session,auth=auth)['data'][0]['Institutions']['name']
    edu_directory = inst_area(session=session,auth=auth)['data'][0]['Areas']['name']
    curr_year = get_curr_period(auth,session)['data'][0]['id']
    
    for i in get_school_students_ids(session=session,auth=auth):
        dic_list.append(
                    {
                        'student_id':i['student_id'],
                        'student__full_name':i['user']['name'],
                        'student_nat':i['user']['nationality_id'],
                        'student_birth_place':i['user']['birthplace_area_id'] if i['user']['birthplace_area_id'] is not None and i['user']['birthplace_area_id'] != 'None' else '' ,
                        'student_birth_date' : i['user']['date_of_birth'] ,
                        'student_nat_id': '' if i['user']['identity_number'] is None else i['user']['identity_number'],
                        'student_grade_id':i['education_grade_id'],
                        'student_grade_name' : i['education_grade_id'] ,
                        'student_class_name_letter': '' if not isinstance(i['institution_class_id'], int) else i['institution_class_id'],
                        'student_edu_place' : edu_directory ,
                        'student_directory':edu_directory,
                        'student_school_name':school_name,
                        'subjects_assessments_info':[] ,
                    }
                    )
            
    sub_dic = {'subject_name':'','subject_number':'','term1':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''} ,'term2':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''}}
    subjects_assessments_info=[]
    # target_student_subjects = list(set(d['education_subject_id'] for d in target_student_marks))

    for chunk in chunks(dic_list, 7):
        student_ids = [i['student_id'] for i in chunk]
        joined_string = ','.join([f'student_id:{i}' for i in student_ids])
        marks = make_request(session=session,auth=auth,url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&academic_period_id={curr_year}&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects&_limit=0&_orWhere='+joined_string)['data']
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
                sub_dic['term1']['assessment1'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A1' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['assessment2'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A2' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['assessment3'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A3' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['assessment4'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A4' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A4' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment1'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A1' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment2'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A2' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment3'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A3' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['assessment4'] = [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A4' in assessments['assessment_period']['code']][0] if [assessments['marks'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A4' in assessments['assessment_period']['code']] else ''
                
                sub_dic['term1']['max_mark_assessment1'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A1' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['max_mark_assessment2'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A2' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['max_mark_assessment3'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A3' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term1']['max_mark_assessment4'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A4' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S1A4' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment1'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A1' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A1' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment2'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A2' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A2' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment3'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A3' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A3' in assessments['assessment_period']['code']] else ''
                sub_dic['term2']['max_mark_assessment4'] = [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A4' in assessments['assessment_period']['code']][0] if [assessments['assessment_grading_option']['max'] for assessments in dictionaries if assessments['assessment_period']  and 'S2A4' in assessments['assessment_period']['code']] else ''
                subjects_assessments_info.append(sub_dic)   
                sub_dic = {'subject_name':'','subject_number':'','term1':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''} ,'term2':{ 'assessment1': '','max_mark_assessment1':'' ,'assessment2': '','max_mark_assessment2':'' , 'assessment3': '','max_mark_assessment3':'' , 'assessment4': '','max_mark_assessment4':''}}
                # [dic for dic in dic_list if dic['student_id']==3439303][0]['subjects_assessments_info']
            target_index = next((i for i, dic in enumerate(dic_list) if dic['student_id'] == student_id != 0 ), None)
            if target_index is not None and len(target_student_subjects) != 0:
                dic_list[target_index]['subjects_assessments_info'].append(subjects_assessments_info)
                # dic_list[target_index]['student_class_name_letter'] = dictionaries[0]['institution_classes_id']
                # print(dic_list[target_index])
                subjects_assessments_info=[]
                target_student_marks = []

    class_name_letter = list(set([i['student_class_name_letter'] for i in dic_list if i['student_class_name_letter'] != '' ]))
    joined_string = ','.join(str(i) for i in [f'institution_class_id:{i}' for i in class_name_letter])
    classes_data = make_request(session=session,auth=auth,url='https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere='+joined_string)['data']            
    class_list = []
    for i in classes_data:
        class_list.append({'class_id': i['institution_class_id'] , 'class_name': i['institution_class']['name'] })
        class_dict = {i['class_id']: i['class_name'] for i in class_list if i['class_id'] != ''}
    for i in dic_list:
        class_id = i['student_class_name_letter']
        if class_id != '':
            i['student_class_name_letter'] = class_dict.get(class_id, class_id)
    grade_id = list(set([i['student_grade_id'] for i in dic_list if i['student_grade_id'] != '' ]))
    grade_data = get_grade_info(auth,session)
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
    birth_place_data = make_request(session=session,auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Area-AreaAdministratives?_limit=0&_contain=AreaAdministrativeLevels')['data']
    nationality_data = make_request(session=session,auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/User-NationalityNames')['data']
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

def get_school_students_ids(auth,session=None):
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    curr_year = get_curr_period(auth)['data'][0]['id']
    students = [
                i['student_id'] 
                for i in make_request(session=session ,auth=auth,url=f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.Students?_limit=0&_finder=Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_number,Users.external_reference,Users.status&institution_id={inst_id}&academic_period_id={curr_year}&_contain=Users')['data']
                    
                    if i['student_status_id'] == 1
                ]
    InstitutionClassStudents = [
                                i 
                                for i in make_request(auth=auth, url=f'https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-InstitutionClassStudents.json?_limit=0&_finder=Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_number,Users.external_reference,Users.status&institution_id={inst_id}&academic_period_id={curr_year}&_contain=Users')['data'] 
                                
                                    if i['student_status_id'] == 1 and i['student_id'] in students
                                ]
    return [
            i
            for i in InstitutionClassStudents
            
            ]

def fill_official_marks_a3_two_face_doc2_offline_version(students_data_lists, ods_file ):
    '''
    doc is the copy that you want to send 
    '''
    context = {'46': 'A6:A30', '4': 'A39:A63', '3': 'L6:L30', '45': 'L39:L63', '44': 'A71:A95', '6': 'A103:A127', '5': 'L71:L95', '43': 'L103:L127', '42': 'A135:A159', '8': 'A167:A191', '7': 'L135:L159', '41': 'L167:L191', '40': 'A199:A223', '10': 'A231:A255', '9': 'L199:L223', '39': 'L231:L255', '38': 'A263:A287', '12': 'A295:A319', '11': 'L263:L287', '37': 'L295:L319', '36': 'A327:A351', '14': 'A359:A383', '13': 'L327:L351', '35': 'L359:L383', '34': 'A391:A415', '16': 'A423:A447', '15': 'L391:L415', '33': 'L423:L447', '32': 'A455:A479', '18': 'A487:A511', '17': 'L455:L479', '31': 'L487:L511', '30': 'A519:A543', '20': 'A551:A575', '19': 'L519:L543', '29': 'L551:L575', '28': 'A583:A607', '22': 'A615:A639', '21': 'L583:L607', '27': 'L615:L639', '26': 'A647:A671', '24': 'A679:A703', '23': 'L647:L671', '25': 'L679:L703'}
    
    page = 4
    name_counter = 1
    name_counter = 1
    
    # classes=[]
    # mawad=[]
    # modified_classes=[]
    
    # Open the ODS file and load the sheet you want to fill
    doc = ezodf.opendoc(ods_file) 
       
    sheet_name = 'sheet'
    sheet = doc.sheets[sheet_name]


    for students_data_list in students_data_lists:
        
#         ['الصف السابع', 'أ', 'اللغة الانجليزية', '786118']
        
        class_data = students_data_list['class_name'].split('=')
        # mawad.append(class_data[2])
        # classes.append('-'.join([class_data[0],class_data[1]]))
        class_name = class_data[0].replace('الصف ' , '').split('-')[0]
        class_char = class_data[0].split('-')[1]
        sub_name = class_data[1]

        
        sheet[f"D{int(context[str(page)].split(':')[0][1:])-5 }"].set_value(f' الصف: {class_name}')
        sheet[f"I{int(context[str(page)].split(':')[0][1:])-5 }"].set_value(f'الشعبة (   {class_char}    )')    
        sheet[f"O{int(context[str(page+1)].split(':')[0][1:])-5}"].set_value(sub_name)
              
        for counter,student_info in enumerate(students_data_list['students_data'], start=1):
            if counter >= 26:
                page += 2
                counter = 1
                
                sheet[f"D{int(context[str(page)].split(':')[0][1:])-5}"].set_value(f' الصف: {class_name}')
                sheet[f"I{int(context[str(page)].split(':')[0][1:])-5}"].set_value(f'الشعبة (   {class_char}    )')  
                sheet[f"O{int(context[str(page+1)].split(':')[0][1:])-5}"].set_value(sub_name)
                #    المادة الدراسية     
                
                # {'id': 3824166, 'name': 'نورالدين محمود راضي الدغيمات', 'term1': {'assessment1': 9, 'assessment2': 10, 'assessment3': 11, 'assessment4': 20}}
                
                for student_info in students_data_list['students_data'][25:] :
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

    # FIXME: make the customshapes crop _20_ to the rest of the key in the custom_shapes
    # Iterate through the cells of the sheet and fill in the values you want
    doc.save()
      
    return 0            
    # return custom_shapes 

def Read_E_Side_Note_Marks_xlsx(file_path=None , file_content=None):
    if file_content is None:
        # Load the workbook
        wb = load_workbook(file_path)
    else:
        wb = load_workbook(filename=file_content)
        
    sheets = wb.sheetnames[:-1]
    # sheet = wb[wb.sheetnames[0]]
    info_sheet = wb[wb.sheetnames[-1]]
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
            if row[1] != '' or row[2] != '': 
                dic = {
                    'id': row[1], 
                    'name':  row[2],
                    'term1': {'assessment1':  row[3], 'assessment2':row[4], 'assessment3': row[5], 'assessment4': row[6]},
                    'term2': {'assessment1': row[8], 'assessment2': row[9], 'assessment3': row[10], 'assessment4': row[11]}
                        }
                data.append(dic)
        temp_dic = {'class_name':sheet ,"students_data": data}
        read_file_output_lists.append(temp_dic)
    
    modified_classes = []

    classes = [i['class_name'].split('=')[0] for i in read_file_output_lists]
    mawad = [i['class_name'].split('=')[1] for i in read_file_output_lists]
    for i in classes: 
        modified_classes.append(mawad_representations(i))
        
    school_name = info_sheet['A1'].value.split('=')[0]
    school_id = info_sheet['A1'].value.split('=')[1]
    modeeriah = info_sheet['A2'].value
    hejri1 = info_sheet['A3'].value
    hejri2 = info_sheet['A4'].value
    melady1 = info_sheet['A5'].value
    melady2 = info_sheet['A6'].value
    baldah = info_sheet['A7'].value
    modified_classes = ' ، '.join(modified_classes)
    mawad = sorted(set(mawad))
    mawad = ' ، '.join(mawad)
    teacher = info_sheet['A8'].value
    required_data_mrks_text = info_sheet['A9'].value
    period_id = info_sheet['A10'].value
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
    'teacher_20_1': teacher,
    'period_id': period_id,
    'school_id' : school_id 
    }
    
    try:
        required_data_mrks_dic_list = {
                                        int(item.split('-')[0]): 
                                            {
                                                'assessment_grade_id': int(item.split('-')[1].split(',')[0]),
                                                'grade_id': int(item.split(',')[0].split('-')[2]), 
                                                'assessments_period_ids': item.split(',')[1:]
                                            }
                                        for item in required_data_mrks_text.split('\\\\')
                                    }
    except Exception as e:
        required_data_mrks_dic_list = {
                                        0:
                                            {
                                                'assessment_grade_id': 0,
                                                'grade_id': 0, 
                                                'assessments_period_ids': 0
                                            }
                                        }

    read_file_output_dict = {'file_data': read_file_output_lists ,
                            'custom_shapes' : custom_shapes ,
                            'required_data_for_mrks_enter' : required_data_mrks_dic_list }
    
    return read_file_output_dict

def enter_marks_arbitrary_controlled_version(username , password , required_data_list ,AssessId, range1='' , range2=''):
    auth = get_auth(username , password)
    period_id = get_curr_period(auth)['data'][0]['id']
    inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    
    for item in required_data_list : 
        for Student_id in item['students_ids']:
            enter_mark(auth 
                ,marks= str("{:.2f}".format(float(random.randint(range1, range2)))) if range1 !='' and range2 !=''  else ''
                ,assessment_grading_option_id= 8
                ,assessment_id= item['assessment_id']
                ,education_subject_id= item['education_subject_id']
                ,education_grade_id= item['education_grade_id']
                ,institution_id= inst_id
                ,academic_period_id= period_id
                ,institution_classes_id= item['institution_classes_id']
                ,student_status_id= 1
                ,student_id= Student_id
                ,assessment_period_id= AssessId)
                        
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
        text += '/All_asses تعبئة كل الامتحانات المتوفرة تلقائيا'
        return text
    
def get_editable_assessments( auth , username ,assessment_grade_id=None , class_subject=None,session=None):
    if assessment_grade_id is None or class_subject is None:
        required_data_list = get_required_data_to_enter_marks(auth=auth ,username=username,session=session)
        ass_data = [[y['assessment_id'],y['education_subject_id']] for y in required_data_list ]
        ass_data = [item for sublist in [get_all_assessments_periods_data2(auth, i[0],i[1],session=session) for i in ass_data] for item in sublist if item.get('editable')==True]
        # unique_lst = [dict(t) for t in {tuple(sorted(d.items())) for d in lst}]
        unique_dict_list = [dict(t) for t in {tuple(sorted(d.items())) for d in ass_data}]
        sorted_dict = sorted(unique_dict_list , key=lambda x: x['code'])
        return sorted_dict
    else:
        ass_data = [item for sublist in [get_all_assessments_periods_data2(auth, assessment_grade_id ,class_subject ,session=session)] for item in sublist if item.get('editable')==True]
        # unique_lst = [dict(t) for t in {tuple(sorted(d.items())) for d in lst}]
        unique_dict_list = [dict(t) for t in {tuple(sorted(d.items())) for d in ass_data}]
        sorted_dict = sorted(unique_dict_list , key=lambda x: x['code'])
        return sorted_dict    
       
def assessments_periods_min_max_mark(auth , assessment_id , education_subject_id ,session=None):
    '''
         استعلام عن القيمة القصوى و الدنيا لكل التقويمات  
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود بمعلومات عن تقيمات الصفوف في السنة الدراسية  
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentItemsGradingTypes.json?_contain=EducationSubjects,AssessmentGradingTypes.GradingOptions&assessment_id={assessment_id}&education_subject_id={education_subject_id}&_limit=0"
    return make_request(url,auth,session=session)

def get_all_assessments_periods_data2(auth , assessment_id ,education_subject_id,session=None):
    '''
         استعلام عن تعريفات التقويمات في السنة الدراسية و امكانية تحرير التقويم و  العلامة القصوى و الدنيا
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود تعريفات التقويمات في السنة الدراسية و امكانية تحرير التقويم و  العلامة القصوى و الدنيا  
    '''
    terms = get_AcademicTerms(auth=auth , assessment_id=assessment_id,session=session)['data']
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

def get_class_students_ids(auth,academic_period_id,institution_subject_id,institution_class_id,institution_id,session=None):
    '''
    استدعاء معلومات عن الطلاب في الصف
    عوامل الدالة هي الرابط و التوكن و تعريفي الفترة الاكاديمية و تعريفي مادة المؤسسة و تعريفي صف المؤسسة و تعريفي المؤسسة
    تعود بمعلومات تفصيلية عن كل طالب في الصف بما في ذلك اسمه الرباعي و التعريفي و مكان سكنه
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id&_limit=0&academic_period_id={academic_period_id}&institution_subject_id={institution_subject_id}&institution_class_id={institution_class_id}&institution_id={institution_id}&_contain=Users"
    student_ids = [student['student_id'] for student in make_request(url,auth,session=session)['data']]
    return student_ids

def get_required_data_to_enter_marks(auth ,username,session=None):
    period_id = get_curr_period(auth,session=session)['data'][0]['id']
    inst_id = inst_name(auth,session=session)['data'][0]['Institutions']['id']
    user_id = user_info(auth,username,session=session)['data'][0]['id']
    years = get_curr_period(auth,session=session)
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = [[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id)['data']]
    required_data_to_enter_marks = []
    
    for class_id in classes_id_1 : 
        class_info = get_teacher_classes2( auth , class_id,session=session)['data']
        dic = {'assessment_id':'','education_subject_id':'' ,'education_grade_id':'','institution_classes_id':'','students_ids':[] }
        dic['assessment_id'] = get_assessment_id_from_grade_id(auth , class_info[0]['institution_subject']['education_grade_id'],session=session)
        dic['education_subject_id'] = class_info[0]['institution_subject']['education_subject_id']
        dic['education_grade_id'] = class_info[0]['institution_subject']['education_grade_id']
        dic['institution_classes_id'] = class_info[0]['institution_class_id']
        dic['class_name'] = class_info[0]['institution_class']['name']
        dic['students_ids'] = get_class_students_ids(auth,period_id,class_info[0]['institution_subject_id'],class_info[0]['institution_class_id'],inst_id,session=session)

        required_data_to_enter_marks.append(dic)
    
    return required_data_to_enter_marks

def get_grade_info(auth,session=None):    
    my_list = make_request(session=session ,auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0')['data']
    return my_list

def get_grade_name_from_grade_id(auth , grade_id):
    
    my_list = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0')['data']

    return [d['name'] for d in my_list if d.get('education_grade_id') == grade_id][0].replace('الفترات التقويمية ل','ا')

def get_assessment_id_from_grade_id(auth , grade_id,session=None):
    
    my_list = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-Assessments.json?_limit=0',session=session)['data']

    return [d['id'] for d in my_list if d.get('education_grade_id') == grade_id][0]

def create_e_side_marks_doc(username , password ,template='./templet_files/e_side_marks.xlsx' ,outdir='./send_folder' ,session=None):
    auth = get_auth(username , password )
    period_id = get_curr_period(auth,session=session)['data'][0]['id']
    user = user_info(auth , username,session=session)
    userInfo = user['data'][0]
    user_id , user_name = userInfo['id'] , userInfo['first_name']+' '+ userInfo['last_name']+'-' + str(username)
    # years = get_curr_period(auth)
    school_data = inst_name(auth,session=session)['data'][0]
    inst_id = school_data['Institutions']['id']
    school_name = school_data['Institutions']['name']
    school_name_id = f'{school_name}={inst_id}'
    baldah = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Institution-Institutions.json?_limit=1&id={inst_id}&_contain=InstitutionLands.CustomFieldValues',session=session)['data'][0]['address'].split('-')[0]
    # grades = make_request(auth=auth , url='https://emis.moe.gov.jo/openemis-core/restful/Education.EducationGrades?_limit=0')
    modeeriah = inst_area(auth)['data'][0]['Areas']['name']
    school_year = get_curr_period(auth,session=session)['data']
    hejri1 = str(hijri_converter.convert.Gregorian(school_year[0]['start_year'], 1, 1).to_hijri().year)
    hejri2 =  str(hijri_converter.convert.Gregorian(school_year[0]['end_year'], 1, 1).to_hijri().year)
    melady1 = str(school_year[0]['start_year'])
    melady2 = str(school_year[0]['end_year'])
    teacher = user['data'][0]['name'].split(' ')[0]+' '+user['data'][0]['name'].split(' ')[-1]
    
    # ما بعرف كيف سويتها لكن زبطت 
    classes_id_1 = sorted([[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id,session=session)['data']])
    classes_id_2 =[get_teacher_classes2( auth , classes_id_1[i],session=session)['data'] for i in range(len(classes_id_1))]
    classes_id_2 =[lst for lst in classes_id_2 if lst]
    classes_id_3 = []  
    assessments_period_data = []
    
    # load the existing workbook
    existing_wb = load_workbook(template)

    # Select the worksheet
    existing_ws = existing_wb.active

    for class_info in classes_id_2:
        classes_id_3.append([{'institution_class_id': class_info[0]['institution_class_id'] ,'sub_name': class_info[0]['institution_subject']['name'],'class_name': class_info[0]['institution_class']['name'] , 'subject_id': class_info[0]['institution_subject']['education_subject_id'] , 'education_grade_id':class_info[0]['institution_subject']['education_grade_id']}])

    for v in range(len(classes_id_2)):
        # id
        print (classes_id_3[v][0]['institution_class_id'])
        id = classes_id_3[v][0]['institution_class_id']
        # subject name 
        print (classes_id_3[v][0]['sub_name'])
        # class name
        print (classes_id_3[v][0]['class_name'])
        class_name = classes_id_3[v][0]['class_name']
        # subject id 
        print (classes_id_3[v][0]['subject_id'])

        
        # copy the worksheet
        new_ws = existing_wb.copy_worksheet(existing_ws)

        # rename the new worksheet
        new_ws.title = (classes_id_3[v][0]['class_name'].replace("الصف",'')+'='+classes_id_3[v][0]['sub_name'].replace('\\','_')+'='+str(classes_id_3[v][0]['institution_class_id'])+'='+str(classes_id_3[v][0]['subject_id'])).replace('/','~')
        new_ws.sheet_view.rightToLeft = True    
        existing_ws.sheet_view.rightToLeft = True   


        students = get_class_students(auth
                                    ,period_id
                                    ,classes_id_1[v]
                                    ,classes_id_3[v][0]['institution_class_id']
                                    ,inst_id
                                    ,classes_id_3[v][0]['education_grade_id'])
        students_names = sorted([i['user']['name'] for i in students['data']])
        print(students_names)
        students_id_and_names = []
        for IdAndName in students['data']:
            students_id_and_names.append({'student_name': IdAndName['user']['name'] , 'student_id':IdAndName['student_id']})

        assessments_json = make_request(auth=auth , url=f'https://emis.moe.gov.jo/openemis-core/restful/Assessment.AssessmentItemResults?academic_period_id={period_id}&education_subject_id='+str(classes_id_3[v][0]['subject_id'])+'&institution_classes_id='+ str(classes_id_3[v][0]['institution_class_id'])+ f'&institution_id={inst_id}&_limit=0&_fields=AssessmentGradingOptions.name,AssessmentGradingOptions.min,AssessmentGradingOptions.max,EducationSubjects.name,EducationSubjects.code,AssessmentPeriods.code,AssessmentPeriods.name,AssessmentPeriods.academic_term,marks,assessment_grading_option_id,student_id,assessment_id,education_subject_id,education_grade_id,assessment_period_id,institution_classes_id&_contain=AssessmentPeriods,AssessmentGradingOptions,EducationSubjects')

        marks_and_name = []

        dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'assessments_periods_ides':[]}
        for student_data_item in students_id_and_names:   
            dic['id'] = student_data_item['student_id'] 
            dic['name'] = student_data_item['student_name'] 
            for student_assessment_item in assessments_json['data']:
                if student_assessment_item['student_id'] == student_data_item['student_id'] :  
                    # FIXME: غير الشرط اذا كان None استبدل القيمة بلا شيء  
                    # FIXME: ضع في الحسبان انه لا توجد علامات و تريد سحب الاسماء في اول الفصل                  
                    if student_assessment_item["marks"] is not None :
                        # dic['id'] = student_data_item['student_id'] 
                        # dic['name'] = student_data_item['student_name'] 
                        if student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment1'] = student_assessment_item["marks"] 
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment2']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment3']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الأول':
                            dic['term1']['assessment4']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الأول' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment1']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثاني' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment2']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الثالث' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment3']  = student_assessment_item["marks"]
                        elif student_assessment_item['assessment_period']['name'] == 'التقويم الرابع' and student_assessment_item['assessment_period']['academic_term'] == 'الفصل الثاني':
                            dic['term2']['assessment4']  = student_assessment_item["marks"]

            marks_and_name.append(dic)
            dic = {'id':'' ,'name': '','term1':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'term2':{ 'assessment1': '' ,'assessment2': '' , 'assessment3': '' , 'assessment4': ''} ,'assessments_periods_ides':[]}
        # Set the font for the data rows
        data_font = Font(name='Arial', size=16, bold=False)

        marks_and_name = [d for d in marks_and_name if d['name'] != '']
        marks_and_name = sorted(marks_and_name, key=lambda x: x['name'])
        if 'عشر' in class_name :
            assessments_period_data_text = ''
            students_id_and_names = sorted(students_id_and_names, key=lambda x: x['student_name'])
            for row_number, dataFrame in enumerate(students_id_and_names, start=3):
                new_ws.cell(row=row_number, column=1).value = row_number-2
                new_ws.cell(row=row_number, column=2).value = dataFrame['student_id']
                new_ws.cell(row=row_number, column=3).value = dataFrame['student_name']
        else:
            assessment_data = '' if  not assessments_json['data'] else assessments_json['data'][0]
            assessment_id = '' if  not assessments_json['data'] else assessment_data['assessment_id']
            education_grade_id = '' if  not assessments_json['data'] else assessment_data['education_grade_id']
            
            assessments_period_data.append({f'{id}-{assessment_id}-{education_grade_id}' : '' if len(marks_and_name) == 0 else marks_and_name[0]['assessments_periods_ides']})
            assessments_period_data_text = '\\\\'.join([str(list(dictionary.items())[0][0]) + ',' + ','.join(str(i) for i in list(dictionary.items())[0][1]) for dictionary in assessments_period_data])
            
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
                new_ws.cell(row=row_number, column=14).value = f'=SUM(H{row_number},M{row_number})/2'
                # Set the font for the data rows
                for cell in new_ws[row_number]:
                    cell.font = data_font
    existing_wb.remove(existing_wb['Sheet1'])

    # Create a new sheet
    new_sheet = existing_wb.create_sheet("info_sheet")
    new_sheet.sheet_view.rightToLeft = True    
    # existing_ws.sheet_view.rightToLeft = True  
    
    # Access the new sheet by name
    info_sheet = existing_wb["info_sheet"]

    # Write data to the new sheet
    info_sheet["A1"] = school_name_id
    info_sheet["A2"] = modeeriah
    info_sheet["A3"] = hejri1
    info_sheet["A4"] = hejri2
    info_sheet["A5"] = melady1
    info_sheet["A6"] = melady2
    info_sheet["A7"] = baldah
    info_sheet["A8"] = teacher
    info_sheet["A9"] = assessments_period_data_text
    info_sheet["A10"] = str(period_id)

    # save the modified workbook
    existing_wb.save(f'{outdir}/{user_name}.xlsx')
    global secondery_students 
    secondery_students = []

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
        if file not in filenames and (file.endswith(".ods") or file.endswith(".pdf") or file.endswith(".bak") or file.endswith(".docx")or file.endswith(".xlsx") ):
            os.remove(os.path.join(dir_path, file))

def fill_official_marks_doc_wrapper_offline(lst, ods_name='send', outdir='./send_folder' ,ods_num=1 , do_not_delete_send_folder=False , templet_file = './templet_files/official_marks_doc_a3_two_face_white_cover.ods', color="#8cd6e6"):
    ods_file = f'{ods_name}{ods_num}.ods'
    copy_ods_file(templet_file , f'{outdir}/{ods_file}')
    fill_official_marks_a3_two_face_doc2_offline_version(students_data_lists=lst['file_data'], ods_file=f'{outdir}/{ods_file}')
    custom_shapes = lst['custom_shapes']
    
    fill_custom_shape(doc= f'{outdir}/{ods_file}' ,sheet_name= 'الغلاف الداخلي' , custom_shape_values= custom_shapes , outfile=f'{outdir}/modified.ods')
    fill_custom_shape(doc=f'{outdir}/modified.ods', sheet_name='الغلاف الازرق', custom_shape_values=custom_shapes, outfile=f"{outdir}/{custom_shapes['teacher']}.ods")
    os.system(f'soffice --headless --convert-to pdf:writer_pdf_Export --outdir {outdir} {outdir}/\"{custom_shapes["teacher"]}.ods\" ')
    add_margins(f"{outdir}/{custom_shapes['teacher']}.pdf", f"{outdir}/output_file.pdf",top_rec=30, bottom_rec=50, left_rec=68, right_rec=120, color_name=color)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/{custom_shapes['teacher']}.pdf",page=1 , top_rec=60, bottom_rec=80, left_rec=70, right_rec=120, color_name=color)
    split_A3_pages(f"{outdir}/output_file.pdf" , outdir)
    reorder_official_marks_to_A4(f"{outdir}/output.pdf" , f"{outdir}/reordered.pdf")

    add_margins(f"{outdir}/reordered.pdf", f"{outdir}/output_file.pdf",top_rec=60, bottom_rec=50, left_rec=68, right_rec=20, color_name=color)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/output_file1.pdf",page=1 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120, color_name=color)
    add_margins(f"{outdir}/output_file1.pdf", f"{outdir}/output_file2.pdf",page=50 , top_rec=100, bottom_rec=80, left_rec=70, right_rec=60, color_name=color)    
    add_margins(f"{outdir}/output_file2.pdf", f"{outdir}/{custom_shapes['teacher']}_A4.pdf",page=51 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120, color_name=color)  
    
    if not do_not_delete_send_folder :
        delete_files_except([f"{custom_shapes['teacher']}.pdf",f"{custom_shapes['teacher']}_A4.pdf",f'{custom_shapes["teacher"]}.ods'], outdir)
    
def fill_official_marks_doc_wrapper(usnername , password , ods_name='send', outdir='./send_folder' ,ods_num=1 , templet_file = './templet_files/official_marks_doc_a3_two_face_white_cover.ods', color="#8cd6e6"):
    ods_file = f'{ods_name}{ods_num}.ods'
    copy_ods_file(templet_file , f'{outdir}/{ods_file}')
    
    custom_shapes = fill_official_marks_a3_two_face_doc2(username= usnername, password= password , ods_file=f'{outdir}/{ods_file}')
    fill_custom_shape(doc= f'{outdir}/{ods_file}' ,sheet_name= 'الغلاف الداخلي' , custom_shape_values= custom_shapes , outfile=f'{outdir}/modified.ods')
    fill_custom_shape(doc=f'{outdir}/modified.ods', sheet_name='الغلاف الازرق', custom_shape_values=custom_shapes, outfile=f'{outdir}/final_'+ods_file)
    os.system(f'soffice --headless --convert-to pdf:writer_pdf_Export --outdir {outdir} {outdir}/final_{ods_file} ')
    add_margins(f"{outdir}/final_{ods_name}{ods_num}.pdf", f"{outdir}/output_file.pdf",top_rec=30, bottom_rec=50, left_rec=68, right_rec=120, color=color)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/{custom_shapes['teacher']}.pdf",page=1 , top_rec=60, bottom_rec=80, left_rec=70, right_rec=120, color=color)
    split_A3_pages(f"{outdir}/output_file.pdf" , outdir)
    reorder_official_marks_to_A4(f"{outdir}/output.pdf" , f"{outdir}/reordered.pdf")

    add_margins(f"{outdir}/reordered.pdf", f"{outdir}/output_file.pdf",top_rec=60, bottom_rec=50, left_rec=68, right_rec=20, color=color)
    add_margins(f"{outdir}/output_file.pdf", f"{outdir}/output_file1.pdf",page=1 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120, color=color)
    add_margins(f"{outdir}/output_file1.pdf", f"{outdir}/output_file2.pdf",page=50 , top_rec=100, bottom_rec=80, left_rec=70, right_rec=60, color=color)    
    add_margins(f"{outdir}/output_file2.pdf", f"{outdir}/{custom_shapes['teacher']}_A4.pdf",page=51 , top_rec=100, bottom_rec=80, left_rec=90, right_rec=120, color=color)  
    delete_files_except([f"{custom_shapes['teacher']}.pdf",f"{custom_shapes['teacher']}_A4.pdf",f'final_{ods_file}'], outdir)

def delete_file(file_path):
    """Delete a file"""
    os.remove(file_path)

def copy_ods_file(source_file_path, destination_folder):
    """Copy an ODS file to a destination folder"""
    shutil.copy(source_file_path, destination_folder)
    
def fill_official_marks_a3_two_face_doc2(username, password , ods_file ,session=None):
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
    classes_id_1 = sorted([[value for key , value in i['InstitutionSubjects'].items() if key == "id"][0] for i in get_teacher_classes1(auth,inst_id,user_id,period_id,session=session)['data']])
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
        classes_id_3.append([{'institution_class_id': class_info[0]['institution_class_id'] ,'sub_name': class_info[0]['institution_subject']['name'],'class_name': class_info[0]['institution_class']['name'] , 'subject_id': class_info[0]['institution_subject']['education_subject_id'] , 'education_grade_id':class_info[0]['institution_subject']['education_grade_id']}])

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
                                    ,inst_id
                                    ,classes_id_3[v][0]['education_grade_id'])
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
                    if mark_data['marks'] is not None:
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

def get_assessments_periods(auth ,term, assessment_id,session=None):
    '''
         استعلام عن تعريفات التقويمات في الفصل الدراسي 
        عوامل الدالة تعريفي السنة الدراسية و التوكن
        تعود بمعلومات عن تقيمات الصفوف في السنة الدراسية  
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentPeriods.json?_finder=academicTerm[academic_term:{term}]&assessment_id={assessment_id}&_limit=0"
    return make_request(url,auth,session=session)

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

def get_AcademicTerms(auth,assessment_id,session=None):
    '''
    دالة لاستدعاء اسم الفصل 
    و عواملها التوكن و رقم تقيم الصف 
    و تعود باسماء الفصول على شكل جيسن
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Assessment-AssessmentPeriods.json?_finder=uniqueAssessmentTerms&assessment_id={assessment_id}&_limit=0"
    return make_request(url,auth,session=session)        

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
            try :
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
            except ValueError : 
                pass
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

def make_request(url, auth ,session=None,timeout_seconds=60):
    headers = {"Authorization": auth, "ControllerAction": "Results"}
    controller_actions = ["Results", "SubjectStudents", "Dashboard", "Staff",'StudentAttendances','SgTree','Students']
    
    for controller_action in controller_actions:
        headers["ControllerAction"] = controller_action
        if session is None :
            response = requests.request("GET", 
                                        url,
                                        headers=headers,
                                        timeout=timeout_seconds,
                                        # verify=False
                                        )
        else : 
            response = session.get(url,
                                   headers=headers,
                                   timeout=timeout_seconds,
                                #    verify=False
                                   )
        if "403 Forbidden" not in response.text :
            return response.json()
        
    return ['Some Thing Wrong']

def get_auth(username , password ,proxies=None):
    ' دالة تسجيل الدخول للحصول على الرمز الخاص بالتوكن و يستخدم في header Authorization'
    url = "https://emis.moe.gov.jo/openemis-core/oauth/login"
    payload = {
        "username": username,
        "password": password
    }
    
    proxies = proxies if proxies else None
    
    response = requests.request("POST",
                                url, data=payload ,
                                proxies=proxies,
                                verify=False
                                # verify=False if proxies else True
                                )

    if response.json()['data']['message'] == 'Invalid login creadential':
        return False
    else: 
        return response.json()['data']['token']    
    
def inst_name(auth,session=None):
    '''
    استدعاء اسم المدرسة و الرقم الوطني و الرقم التعريفي 
        عوامل الدالة الرابط و التوكن
        تعود بالرقم التعريفي و الرقم الوطني و اسم المدرسة 
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Staff?_limit=1&_contain=Institutions&_fields=Institutions.code,Institutions.id,Institutions.name"
    return make_request(url,auth,session=session)   # institution

def inst_area(auth , inst_id = None ,session=None):
    '''
    استدعاء لواء المدرسة و المنطقة
    عوامل الدالة الرابط و التوكن
    تعود باسم البلدية و اسم المنطقة و اللواء 
    '''
    if inst_id is None:
        inst_id = inst_name(auth)['data'][0]['Institutions']['id']
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Institutions.json?id={inst_id}&_contain=AreaAdministratives,Areas&_fields=AreaAdministratives.name,Areas.name"
    return make_request(url,auth,session=session)

def user_info(auth,username,session=None):
    '''
        استدعاء معلومات عن المعلم او المستخدم 
        عوامل الدالة الرابط و التوكن و رقم المستخدم
        تعود برقم المستخدم الوطني و اسمه الرباعي  
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/User-Users?username={username}&is_staff=1&_fields=id,username,openemis_no,first_name,middle_name,third_name,last_name,preferred_name,email,date_of_birth,nationality_id,identity_type_id,identity_number,status&_limit=1"
    return make_request(url,auth,session=session)

def get_teacher_classes1(auth,ins_id,staff_id,academic_period,session=None):
    '''
        استدعاء معلومات صفوف المعلم 
        عوامل الدالة الرابط و التوكن و التعريفي للمدرسة و تعريفي الفترة و staffid 
        تعود الدالة بتعريفي اي صف مع المعلم و كود الصف
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStaff?institution_id={ins_id}&staff_id={staff_id}&academic_period_id={academic_period}&_contain=InstitutionSubjects&_limit=0&_fields=InstitutionSubjects.id,InstitutionSubjects.education_subject_id,InstitutionSubjects.name"
    return make_request(url,auth,session=session)

def get_teacher_classes2(auth,inst_sub_id,session=None):
    '''
    استدعاء معلومات تفصيلية عن الصفوف 
    عوامل الدالة الرابط و التوكن و رقم المستخدم
    تعود باسم الصف و تعريفي الصف و عدد الطلاب في الصف و اسم المادة التي يدرسها المعلم في الصف
    '''
    # url = "https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere=institution_subject_id:10513896,institution_subject_id:10513912,institution_subject_id:10513928,institution_subject_id:10513944"
    url = f"https://emis.moe.gov.jo/openemis-core/restful/Institution.InstitutionClassSubjects?status=1&_contain=InstitutionSubjects,InstitutionClasses&_limit=0&_orWhere=institution_subject_id:{inst_sub_id}"
    
    return make_request(url,auth,session=session)

def get_class_students(auth,academic_period_id,institution_subject_id,institution_class_id,institution_id,education_grade_id,session=None):
    '''
    استدعاء معلومات عن الطلاب في الصف
    عوامل الدالة هي الرابط و التوكن و تعريفي الفترة الاكاديمية و تعريفي مادة المؤسسة و تعريفي صف المؤسسة و تعريفي المؤسسة
    تعود بمعلومات تفصيلية عن كل طالب في الصف بما في ذلك اسمه الرباعي و التعريفي و مكان سكنه
    '''
    url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id,student_status_id,Users.id,Users.username,Users.openemis_no,Users.first_name,Users.middle_name,Users.third_name,Users.last_name,Users.address,Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_type_id,Users.identity_number,Users.external_reference,Users.status,Users.is_guardian&_limit=0&academic_period_id={academic_period_id}&institution_subject_id={institution_subject_id}&institution_class_id={institution_class_id}&institution_id={institution_id}&_contain=Users"
    data = make_request(url,auth,session=session)
    if not data['total']:
        try:
            alt_url = f"https://emis.moe.gov.jo/openemis-core/restful/v2/Institution.InstitutionSubjectStudents?_fields=student_id,student_status_id,Users.id,Users.username,Users.openemis_no,Users.first_name,Users.middle_name,Users.third_name,Users.last_name,Users.address,Users.address_area_id,Users.birthplace_area_id,Users.gender_id,Users.date_of_birth,Users.date_of_death,Users.nationality_id,Users.identity_type_id,Users.identity_number,Users.external_reference,Users.status,Users.is_guardian&_limit=0&_finder=StudentResults[institution_id:{institution_id};institution_class_id:{institution_class_id};assessment_id:{get_assessment_id_from_grade_id(auth,education_grade_id)};academic_period_id:{academic_period_id};institution_subject_id:{institution_subject_id};education_grade_id:{education_grade_id}]&_contain=Users"
            data = make_request(alt_url,auth,session=session)
            if not data['total']:
                raise IndexError
        except IndexError:
            global secondery_students 
            if not len(secondery_students):
                secondery_students =  get_school_students_ids(auth) 
            data = [i for i in secondery_students if i['institution_class_id'] == int(institution_class_id)]
            data = {'data': data , 'total': len(data)}
    return data

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

    response = requests.post(url,headers=headers,json=json_data)
    if response.status_code != 200:
        raise(Exception("couldn't enter the mark for some reason")) 
    else:
        print(marks , education_subject_id ,student_id , response.status_code)

def get_curr_period(auth,session=None):
    '''
    دالة  تستدعي معلومات السنة الحالية من الخادم
    التوكن 
    و تعود على المستخدم بمعلومات السنة الدراسية الحالية 
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/AcademicPeriod-AcademicPeriods?current=1&_fields=id,code,start_date,end_date,start_year,end_year,school_days"
    return make_request(url,auth,session=session)

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
    user = user_info(auth , username)['data']
    user_id = user[0]['id']
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
                #   التقويم الرابع و  الثالث و  الثاني و   الاول  
                # value[2]+ value[3]+ value[4]+value[5]
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

def sort_send_folder_into_two_folders(folder='./send_folder'):
    files = os.listdir(folder)
    pdf_folder = os.path.join(folder, 'PDFs')
    editable_folder = os.path.join(folder, 'قابل للتعديل')

    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(editable_folder, exist_ok=True)

    for file in files:
        if not file.endswith('.json'):
            file_path = os.path.join(folder, file)
            if file.endswith('.pdf'):
                shutil.move(file_path, pdf_folder)
            else:
                shutil.move(file_path, editable_folder)

def main():
    print('starting script')

    # '''9752045067/2761975
    # 9932039648/9932039648
    # 9832049975/9832049975
    # 9892023550/0772626275
    # 9902044251/9902044251
    # 9942011966/9942011966
    # 9872556471/9872556471
    # 1300902041/1300902041
    # 9832041371/123456
    # 9732044574/222222
    # 9932039564/9932039564
    # 9862045517/9862045517
    # 9852045060/9852045060
    # 9862055767/123456
    # 9892023326/9892023326
    # 2000223096/Besan@2001
    # 9902050711/9902050711
    # 9762045800/9762045800
    # 9942022052/Aa@9942022052
    # 9922052534/20182018
    # 9942036547/9942036547
    # 9842053654/654321
    # 9962040167/9962040167E$e
    # 9892055264/9892055264
    # 9922052664/9922052664
    # 1635857406/123456
    # 9832008276/ANMSOA
    # 9842053211/9842053211
    # 9722011390/9722011390
    # 9842048442/123456
    # 9921009580/9921009580
    # 9971055725/9971055725
    # 9991039132/9991039132Mm@
    # 9991014194/Zzaid#079079
    # 9961055140/Mtm#123456789
    # 9862053521
    # 0772323488/weam@137342
    # 9782051311/Aa@12345678
    # 9762051028
    # 9862049623/199435
    # 9772015488
    # 9692012484
    # 9781053164'''
    # passwords = '''9822041975/Aa@9822041975'''

    # تعمل في مؤسستين 
    # 9892050032/Manar@100 
    # فيها خلل غريب لا اعرف عنه 
    # الي هو الرابط عندي فيه مشكلة غريبة
    # 9892022099/9892022099 
    # برضو هذي بدها تعديل و شغل 
    # في صفوف بسحبهن api بس ما بظهرن على المنظومة (على الويب)
    # في الصف التوجيهي ما بدقر عند رقم الاسيسمنت ما بقدر يحوله
    # 2000223096/Besan@2001

    # File "/opt/programming/school_programms1/telegram_bot/utils1.py", line 4464, in create_e_side_marks_doc
    #     students = get_class_students(auth
    # ما بعرف سبب هذا الخطأ
    # 9942022052/Aa@9942022052

    #   File "/opt/programming/school_programms1/telegram_bot/utils1.py", line 5622, in main
    #     print(username , password)
    #     ^^^^^^^^^^^^^^^^^^^^^^^^^^^
    #   File "/opt/programming/school_programms1/telegram_bot/utils1.py", line 4444, in create_e_side_marks_doc
    #     print (classes_id_3[v][0]['institution_class_id'])
    #            ~~~~~~~~~~~~^^^
    # IndexError: list index out of range
    # 9922052534/20182018

    # There is error in 
    # 9962040167/9962040167E$e
    # There is error in 
    # 9842053211/9842053211
    # There is error in 
    # 9842048442/123456
    # There is error in 
    # 9862053521/9862053521
    # There is error in 
    # 9782051311/Aa@12345678
    # There is error in 
    # 9762051028/9762051028
    # There is error in 
    # 9692012484/9692012484
    # There is error in 
    # 9781053164/9781053164

    # bulk_e_side_note_marks(passwords)
    
    convert_to_marks_offline_from_send_folder(template='./templet_files/official_marks_doc_a3_two_face_white_cover.ods', color='#FFFFFF')
    
    # read_all_xlsx_in_folder()
    # fill_student_absent_doc_wrapper(9971055725,9971055725)


if __name__ == "__main__":
    main()