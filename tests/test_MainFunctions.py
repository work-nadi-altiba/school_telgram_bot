import pytest
import sys

try:
    import telegram_bot.utils1 as utils1
    import telegram_bot.bot as bot
except ModuleNotFoundError:
    sys.path.append('/home/kali/programming/school_programms1/telegram_bot')
    
from utils1 import create_e_side_marks_doc , compare_files , fill_official_marks_functions_wrapper_v2 , get_auth , fill_absent_document_wrapper_v2 , get_academic_periods , create_tables_wrapper , create_from_certs_template_wrapper
from bot import count_files , delete_send_folder

outdir='./tests/outdir' 

@pytest.mark.parametrize('file',[('9971055725=Aa@9971055725=15=0.xlsx'),('99310068300=99310068300@Mm=15=0.xlsx')])
def test_e_side_marks_with_marks(file):
    name= file.replace('.xlsx','').split('=')
    username = name[0]
    password = name[1]
    template= './telegram_bot/templet_files/e_side_marks.xlsx'
    outdir='./tests/outdir' 
    period_id = name[2]
    empty_marks = bool(int(name[3])) 
    create_e_side_marks_doc(username,password,template,outdir=outdir,period_id=period_id,empty_marks=empty_marks)
    files = count_files(outdir+'/*')

    diff = compare_files(f'./tests/sample_files/{file}' ,files[0] )
    assert len(diff) == 0
    delete_send_folder(outdir+'/*')

@pytest.mark.parametrize('file',[
    ('9971055725=Aa@9971055725=15=0=p1.ods'),
    (['99310068300=99310068300@Mm=15=0=p1=a3.ods' ,'99310068300=99310068300@Mm=15=0=p2=a3.ods']),
    ]
)
def test_official_marks(file):
    template='./telegram_bot/templet_files/official_marks_doc_a3_two_face_white_cover.ods'
    try:
        if isinstance(file, list):
            name= file[0].replace('.ods','').split('=')
            username = name[0]
            password = name[1]
            period_id = name[2]
            empty_marks = bool(int(name[3])) 
            fill_official_marks_functions_wrapper_v2(username , password , templet_file=template , outdir=outdir , period_id=period_id , empty_marks=empty_marks , convert_to_pdf=False)
            outdir_files = count_files(outdir+'/*')
            for file_part in file:
                part_num= file_part.replace('.ods','').split('=')[4].replace('p','')
                wanted_file = [i for i in outdir_files if f'جزء_{part_num}' in i][0]
                diff = compare_files(f'./tests/sample_files/{file_part}' , wanted_file )
                assert len(diff) == 0
            delete_send_folder(outdir+'/*')
        else:
            name= file.replace('.ods','').split('=')
            username = name[0]
            password = name[1]
            period_id = name[2]
            empty_marks = bool(int(name[3])) 
            fill_official_marks_functions_wrapper_v2(username , password , templet_file=template , outdir=outdir , period_id=period_id , empty_marks=empty_marks , convert_to_pdf=False)
            files = count_files(outdir+'/*')
            diff = compare_files(f'./tests/sample_files/{file}' ,files[0] )
            assert len(diff) == 0
            delete_send_folder(outdir+'/*')
    except:
        delete_send_folder(outdir+'/*')

@pytest.mark.parametrize('file',[
    ('99310068300=99310068300@Mm=15=1=absent_file.ods'),
    ]
)
def test_absent_document(file):
    name= file.replace('.ods','').split('=')
    username = name[0]
    password = name[1]
    period_id = name[2]
    get_student_absent=bool(int(name[3])) 
    auth  = get_auth(username,password)
    curr_period_data = get_academic_periods(auth , period_id)
    fill_absent_document_wrapper_v2(auth , username , ods_file='/home/kali/programming/school_programms1/telegram_bot/templet_files/emishub_st_abs_A3.ods', curr_period_data=curr_period_data , get_student_absent=get_student_absent,outdir=outdir+'/')
    files = count_files(outdir+'/*')
    wanted_file = [i for i in files if 'one_step_more.ods' not in i][0]
    diff = compare_files(f'./tests/sample_files/{file}' , wanted_file )
    assert len(diff) == 0
    delete_send_folder(outdir+'/*')

@pytest.mark.parametrize('file',[
    ('9991014194=Zzaid#079079=15=1=table.xlsx'),
    ]
)
def test_table_document(file):
    name= file.replace('.xlsx','').split('=')
    username = name[0]
    password = name[1]
    period_id = name[2]
    get_student_absent=bool(int(name[3])) 
    term= bool(int(name[4]))
    auth  = get_auth(username,password)
    curr_period_data = get_academic_periods(auth , period_id)
    create_tables_wrapper(username , password ,term2=term , just_teacher_class=True , curr_year = curr_period_data , template='/home/kali/programming/school_programms1/telegram_bot/templet_files/tamplete_table.xlsx',outdir=outdir+'/')
    files = count_files(outdir+'/*')

    diff = compare_files(f'./tests/sample_files/{file}' ,files[0] )
    assert len(diff) == 0
    delete_send_folder(outdir+'/*')

@pytest.mark.parametrize('file',[
    ('9991014194=Zzaid#079079=15=1=1=certs_word'),
    ]
)
def test_certs_word(file):
    name= file.split('=')
    username = name[0]
    password = name[1]
    period_id = name[2]
    get_student_absent=bool(int(name[3])) 
    term= bool(int(name[4]))
    auth  = get_auth(username,password)
    curr_period_data = get_academic_periods(auth , period_id)
    
    create_from_certs_template_wrapper(
                                        username ,
                                        password ,
                                        term2=term ,
                                        just_teacher_class=True ,
                                        curr_year = curr_period_data ,
                                        skip_art_sport=True ,
                                        template='/home/kali/programming/school_programms1/telegram_bot/templet_files/cartoon 1-10 FC_modified_windows.docx' ,
                                        outdir=outdir+'/',
                                    )
    
    files = count_files(outdir+'/*')

    diff = compare_files(f'./tests/sample_files/{file}' ,files[0] )
    for value in list(diff.values()):
        assert len(value) == 0
    delete_send_folder(outdir+'/*')

# test_official_marks('9971055725=Aa@9971055725=15=0=p1.ods')
# test_absent_document('99310068300=99310068300@Mm=15=1=absent_file.ods')
test_table_document('9991014194=Zzaid#079079=15=1=1=table.xlsx')
# test_certs_word('9991014194=Zzaid#079079=15=1=1=certs_word')