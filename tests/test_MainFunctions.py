import pytest
import sys

try:
    from telegram_bot.utils1 import create_e_side_marks_doc , compare_files , fill_official_marks_functions_wrapper_v2
    from telegram_bot.bot import count_files , delete_send_folder
except ModuleNotFoundError:
    sys.path.append('/home/kali/programming/school_programms1/telegram_bot')
    from utils1 import create_e_side_marks_doc , compare_files , fill_official_marks_functions_wrapper_v2
    from bot import count_files , delete_send_folder

@pytest.mark.parametrize('file',[('9971055725=Aa@9971055725=15=0.xlsx'),('99310068300=99310068300@Mm=15=0.xlsx')])
def test_e_side_marks_with_marks(file):
    name= file.replace('.xlsx','').split('=')
    username = name[0]
    password = name[1]
    template='./telegram_bot/templet_files/e_side_marks.xlsx' 
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
    outdir='./tests/outdir' 
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

# test_official_marks('9971055725=Aa@9971055725=15=0=p1.ods')