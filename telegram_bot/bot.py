# pip install python-telegram-bot
# pip install python-telegram-bot==13.7

from telegram.ext import Updater, CommandHandler, MessageHandler, Filters,ConversationHandler
from tqdm.contrib.telegram import tqdm, trange
from telegram import ReplyKeyboardMarkup
from telegram.ext import *
from telegram import Bot
from utils1 import *
from keys import production_bot , test_bot 
import argparse
import traceback
import io

print('Starting up bot...')

# # fill assess arbitrary marks conversation handler stats
INIT_F , RESPOND = range(2)
WAITING_FOR_RESPONSE = range(1)
CREDS, AVAILABLE_ASS ,WAITING_FOR_RESPONSE = range(3)
CREDS, FILE = range(2)
CREDS_2 ,ASK_QUESTION   = range(2)


help_text = '''
/marks_up_percentage كشف بنسبة ادخال علامات كل معلم في المدرسة
/students_basic_info كشف البيانات الاساسية للطلاب
/student_absent_document طباعة دفتر الحضور و الغياب الكتروني
/e_side_marks_note لطباعة كشف علامات جانبي الكتروني 
/side_marks_note  لطباعة كشف العلامات الجانبي العادي
/performance_side_marks_note لطباعة كشف العلامات و الاداء الجانبي 
/fill_assess_arbitrary لتسجيل العلامات العشوائية 
/empty_assess لمسح علامات الصف 
/check_marks اخذ عينة علامات طلاب في الصف
/official_marks لطباعة سجل العلامات الرسمي 
/certs لطباعة ملف الشهادات 
/tables لطباعة ملفات الجداول 
/links رابط بنك اسئلة المواد لاضافة اسئلة
/cancel لألغاء العملية
'''

# اكتب الكود تحت هذا التعليق
# write the function below here 

def test_qt(update, context):
    chat_id = update.message.chat.id
    for i in trange(100, token=f'{token}', chat_id=f'{chat_id}'):
        import time 
        time.sleep(.5)
        print(i)

def init_absent_doc(update, context):
    update.message.reply_text("هل تريد طباعة سجل الحضور و الغياب ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def send_absent_notebook_doc(update, context):
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        user = update.message.from_user
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        # update.message.reply_text("Thanks for sharing! You're a credentials user {} and password {}.".format(context.user_data['creds'][0], context.user_data['creds'][1] ) )
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            logger.info(f"username:{username} ----> password :{password}")
            update.message.reply_text("انتظر لحظة لو سمحت") 
            fill_student_absent_doc_wrapper(username, password)
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END

def check_if_send_folder_exist():
    folder_name = "send_folder"
    # Check if the folder exists
    if not os.path.exists(folder_name):
        # If it doesn't exist, create it
        os.makedirs(folder_name)
        print(f"Folder '{folder_name}' created.")

# https://forms.gle/1PMSeb75mQVJUEnw5
def links(update, context):
    update.message.reply_text("رابط بنك الاسئلة : \nhttps://forms.gle/1PMSeb75mQVJUEnw5") 
    return CREDS_2

def init_side_marks(update, context):
    update.message.reply_text("هل تريد كشف العلامات الجانبي العادي ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS_2

def send_side_marks_note_doc(update, context):
    term = update.message.text.replace('/','')
    username = context.user_data['creds'][0]
    password = context.user_data['creds'][1]
    # session = requests.Session()
        
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        update.message.reply_text("انتظر لحظة لو سمحت")     
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
            return CREDS_2
        else:
            logger.info(f"username:{username} ----> password :{password}")
            if term == 'term1':            
                side_marks_document_with_marks(username , password ,term=1 )
            elif term == "term2":
                side_marks_document_with_marks(username , password ,term=2 )
            else:
                func_text = '''اختر الفصل بالضغط عليه
                /term1 الفصل الاول 
                /term2 الفصل الثاني '''
                
                update.message.reply_text("لم تختر فصل ")
                
                update.message.reply_text(func_text) 
                
                            
            # side_marks_document(username, password)
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END

def init_check_five_names_marks(update, context):
    update.message.reply_text("هل تريد التحقق من علامات الطلاب ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS_2

def which_term(update, context):
    context.user_data['creds'] = update.message.text.split('/')
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        func_text = '''اختر الفصل بالضغط عليه
        /term1 الفصل الاول 
        /term2 الفصل الثاني '''
        update.message.reply_text(func_text) 
        return ASK_QUESTION

def init_marks_up_percentage(update, context):
    update.message.reply_text("هل تريد التحقق من علامات الطلاب ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS_2

def get_marks_up_percentage(update, context):
    term = update.message.text.replace('/','')
    username = context.user_data['creds'][0]
    password = context.user_data['creds'][1]
    auth = get_auth(username, password)
    session = requests.Session()
    update.message.reply_text("هذا الامر قد يستغرق بعض الوقت ارجوا منك الانتظار ")     
    if update.message.text == '/cancel':
        return cancel(update, context)
    elif update.message.text == re.findall(r'.*', update.message.text) and update.message.text != '/cancel' :
        update.message.reply_text("انتظر لحظة لو سمحت") 
    else:
        if term == 'term1':
            teachers_marks_upload_percentage_wrapper_version_2(auth=auth,first_term=True,session=session)
        elif term == "term2":
            teachers_marks_upload_percentage_wrapper_version_2(auth=auth,second_term=True,session=session)
        elif update.message.text == '/cancel':
            return cancel(update, context)        
        generate_pdf(f'./send_folder/نسب العلامات للاساسي.xlsx' , './send_folder' ,'نسب العلامات للاساسي')
        files = count_files()
        chat_id = update.message.chat.id
        send_files(bot, chat_id, files)
        delete_send_folder()
        return ConversationHandler.END

def print_check_five_names_marks(update, context):
    term = update.message.text.replace('/','')
    username = context.user_data['creds'][0]
    password = context.user_data['creds'][1]
    auth = get_auth(username, password)
    session = requests.Session()
    update.message.reply_text("انتظر لحظة لو سمحت")     
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        if term == 'term1':
            text = five_names_every_class_wrapper(auth,username,1,session)
        elif term == "term2":
            text = five_names_every_class_wrapper(auth,username,2,session)
        
        for t in text.split('----------------------------------------------------------------------')[:-1]:
            update.message.reply_text(t) 
        logger.info(f"username:{username} ----> password :{password}")
        return ConversationHandler.END

def upload_marks_bot_version(update, context):
    session = requests.Session()
    question = context.user_data['quetion']
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        print(username, password)
        auth = get_auth(username,password)
        if auth == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
            return CREDS_2
        else:
            logger.info(f"username:{username} ----> password :{password}")
            update.message.reply_text("انتظر لحظة لو سمحت") 

            file_id = context.user_data['file']
            file_name =context.user_data['file_name'] 
            file_extension = file_name.split('.')[-1].lower()
            username , password = context.user_data['creds'][0] , context.user_data['creds'][1]
            # Get the file object and read its content
            file_obj = context.bot.get_file(file_id)
            file_bytes = io.BytesIO(file_obj.download_as_bytearray())


            if question == 'check_upload' :
                if file_extension == 'xlsx':
                    excel_output = Read_E_Side_Note_Marks_xlsx(file_content=file_bytes)
                    # fill_official_marks_doc_wrapper_offline(Read_E_Side_Note_Marks_xlsx(file_content=file_bytes)) if question == 'document_marks' else None
                elif file_extension == 'ods':
                    excel_output = Read_E_Side_Note_Marks_ods(file_content=file_bytes)

                max_marks_for_classes_based_on_subject_id = create_max_of_dictionaries(excel_output , auth , session)
                above_max_marks = find_above_max_mark_for_assessments(excel_output , max_marks_for_classes_based_on_subject_id)
                if len(above_max_marks):
                    update.message.reply_text("هذه العلامات اكبر من العلامات القصوى و تحتاج الى تعديل")
                    for item in above_max_marks: 
                        # print(item['class_name'].replace('الصف ' , '') , item['subject_name'] , item['term&assess'] , item['name'] , item['mark'] , sep=' -- ')
                        joined_item = f"{item['class_name'].replace('الصف ' , '')} -- {item['subject_name']} -- {item['term&assess']} -- {item['name']} -- {item['mark'] }" 
                        update.message.reply_text(joined_item) 
                    return ConversationHandler.END 
                else:
                    update.message.reply_text("لا توجد علامات تحتاج الى تعديل")
                    return ConversationHandler.END 
            # TODO: handle empty editable_assessments list
            editable_assessments = get_editable_assessments(auth ,username ,session)
            data_to_enter_marks = get_required_data_to_enter_marks(auth ,username,session)
            string = assessments_commands_text(editable_assessments)
            update.message.reply_text(string)
            context.user_data['assessments'] = editable_assessments
            context.user_data['data_to_enter_marks'] = data_to_enter_marks
            
            # update.message.reply_text("سوف احاول ادخال العلامات بعد مسح اي علامة على المنظومة") 
            update.message.reply_text("سوف احاول ادخال العلامات") 

            # # فرغ كل التقومات من العلامات
            # assess_data = [i for i in editable_assessments]
            # enter_marks_arbitrary_controlled_version(username , password , data_to_enter_marks , assess_period_data=assess_data)

            # عبي التقويم من ملف سجل العلامات الجانبي
            if file_extension == 'xlsx':
                upload_marks_optimized(username,password,Read_E_Side_Note_Marks_xlsx(file_content=file_bytes))
                # fill_official_marks_doc_wrapper_offline(Read_E_Side_Note_Marks_xlsx(file_content=file_bytes)) if question == 'document_marks' else None
            elif file_extension == 'ods':
                upload_marks_optimized(username,password,Read_E_Side_Note_Marks_ods(file_content=file_bytes))
                # fill_official_marks_doc_wrapper_offline(Read_E_Side_Note_Marks_ods(file_content=file_bytes)) if question == 'document_marks' else None
                
            # if question == 'document_marks' :
            #     files = count_files()
            #     chat_id = update.message.chat.id
            #     context.user_data['chat_id'] = chat_id
            #     send_files(bot, chat_id, files)
            #     delete_send_folder()
        
        update.message.reply_text("تمام انتهينا")
        return ConversationHandler.END

# Define a function to handle incoming files
def init_empty_fill(update, context):
    update.message.reply_text("هل تريد مسح علامات صف ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def print_available_assessments_light_version(update, context):
    user = update.message.from_user
    context.user_data['creds'] = update.message.text.split('/')
    username = context.user_data['creds'][0]
    password = context.user_data['creds'][1]
    print(username, password)
    if get_auth(username, password) == False:
        update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
    else:
        logger.info(f"username:{username} ----> password :{password}")
        update.message.reply_text("انتظر لحظة لو سمحت") 
        auth = get_auth(username,password)
        # TODO: handle empty editable_assessments list
        editable_assessments = get_editable_assessments(auth ,username)
        string = assessments_commands_text(editable_assessments)
        update.message.reply_text(string)
        context.user_data['assessments'] = editable_assessments
        return  AVAILABLE_ASS

def print_available_assessments(update, context):
    user = update.message.from_user
    session = requests.Session()
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        print(username, password)
        auth = get_auth(username,password)
        if auth == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            logger.info(f"username:{username} ----> password :{password}")
            update.message.reply_text("انتظر لحظة لو سمحت") 
            # TODO: handle empty editable_assessments list
            editable_assessments = get_editable_assessments(auth ,username,session=session)
            data_to_enter_marks = get_required_data_to_enter_marks(auth ,username ,session=session)
            string = assessments_commands_text(editable_assessments)
            update.message.reply_text(string)
            context.user_data['assessments'] = editable_assessments
            context.user_data['data_to_enter_marks'] = data_to_enter_marks
            return  AVAILABLE_ASS

def fill_assess_empty(update, context):
    user = update.message.from_user
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        code = update.message.text.replace('/','')
        update.message.reply_text("انتظر لحظة لو سمحت")         
        editable_assessments = context.user_data['assessments'] 
        data_to_enter_marks = context.user_data['data_to_enter_marks']  
        username , password = context.user_data['creds'][0] , context.user_data['creds'][1]
        logger.info(f"username:{username} ----> password :{password}// {code}")
        if code == 'All_asses':
            assess_data = [i for i in editable_assessments]
            enter_marks_arbitrary_controlled_version(username , password , data_to_enter_marks , assess_period_data=assess_data)
        else:
            assess_data = [i for i in editable_assessments if i.get('code') == code][0]
            wanted_grades = [i for i in data_to_enter_marks if i.get('assessment_id') == assess_data['gradeId']]
            assess_data = [i for i in editable_assessments if i.get('code') == code][0]
            wanted_grades = [i for i in data_to_enter_marks if i.get('assessment_id') == assess_data['gradeId']]
            # enter_marks_arbitrary_controlled_version(username , password , required_data_list ,AssessId=None, assess_period_data=None ,range1='' , range2=''):            
            enter_marks_arbitrary_controlled_version(username,password,wanted_grades,assess_data['AssesId'] , [assess_data])
            update.message.reply_text("هل تريد مسح علامات صف اخر؟ نعم | لا",reply_markup=ReplyKeyboardMarkup([['نعم', 'لا']], one_time_keyboard=True))
            return WAITING_FOR_RESPONSE
        # End of conversation
        update.message.reply_text("تمام انتهينا")
        return ConversationHandler.END

def handle_response(update, context):
    response = update.message.text
    if response == 'نعم':
        return AVAILABLE_ASS
    else:
        update.message.reply_text("تمام انتهينا")
        return ConversationHandler.END

def receive_file(update, context ):
    '''
        dispatcher.add_handler(MessageHandler(Filters.document, receive_file))
    '''
    
    message = update.message
    if message.document:
        file_id = message.document.file_id
        file_name = message.document.file_name
        context.user_data['file'] = file_id
        context.user_data['file_name'] = file_name
        logger.info(f"{file_name}")
        if file_name.split('.')[-1].lower() == 'ods':
            file_obj = context.bot.get_file(file_id)
            file_bytes = io.BytesIO(file_obj.download_as_bytearray())
    
            if check_file_if_official_marks_file(file_content=file_bytes):
                convert_official_marks_doc(file_content=file_bytes)
                files = count_files()
                chat_id = update.message.chat.id
                context.user_data['chat_id'] = chat_id
                send_files(bot, chat_id, files)
                delete_send_folder()                
                update.message.reply_text("تمام انتهينا")                
                return ConversationHandler.END
            
        receive_file_massage = '''/document_marks  طباعة سجل العلامات و ادخال العلامات معا
/document طباعة سجل العلامات الرسمي من الملف فقط
/documentFirstThree طباعة سجل العلامات الرسمي من الملف فقط
/side_notes_first_term طباعة الكشف الجانبي الفصل الاول
/side_notes_second_term طباعة الكشف الجانبي الفصل الثاني 
/two_terms_side_note طباعة كشف علامات فصلين 
/check_upload التأكد من صحة العلامات و رفعها اذا لم توجد مشكلة
/marks ادخال العلامات من الملف فقط
'''
        update.message.reply_text(receive_file_massage)  
        return ASK_QUESTION
    else:
        update.message.reply_text('No document file received. Please send a document file.')
        return ConversationHandler.END

def handle_question(update, context):
    question = update.message.text.replace('/','')
    context.user_data['quetion'] = question
    file_id = context.user_data['file']
    file_name =context.user_data['file_name'] 
    file_extension = file_name.split('.')[-1].lower()
    
    # Get the file object and read its content
    file_obj = context.bot.get_file(file_id)
    file_bytes = io.BytesIO(file_obj.download_as_bytearray())
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        if question == 'document_marks':
            update.message.reply_text("انتظر لحظة لو سمحت \nو اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456")
            if file_extension == 'xlsx':           
                fill_official_marks_doc_wrapper_offline(Read_E_Side_Note_Marks_xlsx(file_content=file_bytes))
            elif file_extension == 'ods':
                fill_official_marks_doc_wrapper_offline(Read_E_Side_Note_Marks_ods(file_content=file_bytes))
            files = count_files()
            chat_id = update.message.chat.id
            context.user_data['chat_id'] = chat_id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return CREDS_2
        elif question == 'document':
            update.message.reply_text("انتظر لحظة لو سمحت")  
            if file_extension == 'xlsx':
                fill_official_marks_functions_wrapper_v2( e_side_notebook_data=Read_E_Side_Note_Marks_xlsx(file_content=file_bytes),divded_dfter_to_primary_and_secnedry=False)
            elif file_extension == 'ods':
                fill_official_marks_functions_wrapper_v2( e_side_notebook_data=Read_E_Side_Note_Marks_ods(file_content=file_bytes),divded_dfter_to_primary_and_secnedry=False)
            files = count_files()
            chat_id = update.message.chat.id
            context.user_data['chat_id'] = chat_id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END
        elif question == 'documentFirstThree':
            update.message.reply_text("انتظر لحظة لو سمحت")  
            if file_extension == 'xlsx':
                fill_official_marks_doc_wrapper_offline(Read_E_Side_Note_Marks_xlsx(file_content=file_bytes),templet_file='./templet_files/official_marks_document_from_grade_1-3_white_cover.ods')
            elif file_extension == 'ods':
                fill_official_marks_doc_wrapper_offline(Read_E_Side_Note_Marks_ods(file_content=file_bytes),templet_file='./templet_files/official_marks_document_from_grade_1-3_white_cover.ods')
            files = count_files()
            chat_id = update.message.chat.id
            context.user_data['chat_id'] = chat_id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END            
        elif question == 'side_notes_first_term':
            update.message.reply_text("انتظر لحظة لو سمحت")  
            if file_extension == 'xlsx':           
                side_marks_document_with_marks(term=1,classes_data=Read_E_Side_Note_Marks_xlsx(file_content=file_bytes))
            elif file_extension == 'ods':
                side_marks_document_with_marks(term=1,classes_data=Read_E_Side_Note_Marks_ods(file_content=file_bytes))
            files = count_files()
            chat_id = update.message.chat.id
            context.user_data['chat_id'] = chat_id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END  
        elif question == 'side_notes_second_term':
            update.message.reply_text("انتظر لحظة لو سمحت")  
            if file_extension == 'xlsx':           
                side_marks_document_with_marks(term=2,classes_data=Read_E_Side_Note_Marks_xlsx(file_content=file_bytes))
            elif file_extension == 'ods':
                side_marks_document_with_marks(term=2,classes_data=Read_E_Side_Note_Marks_ods(file_content=file_bytes))
            files = count_files()
            chat_id = update.message.chat.id
            context.user_data['chat_id'] = chat_id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END
        elif question == 'two_terms_side_note':
            update.message.reply_text("انتظر لحظة لو سمحت")
            if file_extension == 'xlsx':           
                create_two_terms_side_marks_doc(classes_data=Read_E_Side_Note_Marks_xlsx(file_content=file_bytes))
            elif file_extension == 'ods':
                create_two_terms_side_marks_doc(classes_data=Read_E_Side_Note_Marks_ods(file_content=file_bytes))
            files = count_files()
            chat_id = update.message.chat.id
            context.user_data['chat_id'] = chat_id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END
        elif question == 'check_upload':
            update.message.reply_text("اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
            return CREDS_2
        elif question == 'marks':
            update.message.reply_text("اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
            return CREDS_2
        else:
                update.message.reply_text('ادخال خاطيء')
                return ASK_QUESTION

def start(update, context):
    context.bot.send_message(chat_id=update.effective_chat.id, text=help_text)

def send_files(bot, chat_id, files , outdir='./send_folder',name="ملف مضغوظ"):
    if len(files) >= 4:
        create_zip(files,zip_name=name)
        delete_files_except(name , outdir)
        bot.send_document(chat_id=chat_id, document=open(f'{outdir}/{name}.rar', 'rb'), timeout=900)
    else:
        for file in files:
            bot.send_document(chat_id=chat_id, document=open(file, 'rb'), timeout=900)
        return False

# Lets us use the /help command1
def help_command(update, context):
    update.message.reply_text(help_text)

# Log errors
def error(update, context):
    username = context.user_data['creds'][0]
    password = context.user_data['creds'][1]
    update.message.reply_text(f"حصل هذا الخطا : {context.error} ")
    traceback.print_exc()
    logger.error(f"username:{username} ----> password :{password}"+'\n'+traceback.format_exc())
    print(f'Update {update} caused error {context.error}')
    return ConversationHandler.END

def cancel(update, context):
    user = update.message.from_user
    update.message.reply_text("تم ")
    return ConversationHandler.END

def get_user_creds(update, context):
    update.message.reply_text("اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return AVAILABLE_ASS

def check_user_creds(update, context):
    # code
    pass

def init_fill(update, context):
    update.message.reply_text("هل تريد تسجل علامات عشوائي ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def fill_assess_arbitrary(update, context):
    user = update.message.from_user
    user_range1,user_range2 = [""]*2
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        response = update.message.text.replace('/','').split(' ')
        if len(response) ==3 : 
            code ,user_range1 , user_range2 = response
        else:
            code = response[0]
        update.message.reply_text("انتظر لحظة لو سمحت")
        editable_assessments = context.user_data['assessments']
        data_to_enter_marks = context.user_data['data_to_enter_marks']
        username , password = context.user_data['creds'][0] , context.user_data['creds'][1]
        logger.info(f"username:{username} ----> password :{password}")
        if code == 'All_asses':
            assess_data = [i for i in editable_assessments]
            for assessment in assess_data:
                if not len(user_range1):
                    range1 , range2 = int(assessment['max_mark']*.75),assessment['max_mark']
                else:
                    range1 , range2 = user_range1 , user_range2
                wanted_grades = [i for i in data_to_enter_marks if i.get('assessment_id') == assessment['gradeId']]
                enter_marks_arbitrary_controlled_version(username,password ,wanted_grades ,AssessId=assessment['AssesId'] ,range1=int(range1),range2=int(range2),assess_period_data=editable_assessments)
        else:
            assess_data = [i for i in editable_assessments if i.get('code') == code][0]
            wanted_grades = [i for i in data_to_enter_marks if i.get('assessment_id') == assess_data['gradeId']]
            if not len(user_range1):
                user_range1 , user_range2 = assess_data['pass_mark'],assess_data['max_mark']
            enter_marks_arbitrary_controlled_version(username,password ,wanted_grades ,AssessId=assess_data['AssesId'] ,range1=int(user_range1),range2=int(user_range2),assess_period_data=[assess_data])
            # enter_marks_arbitrary_controlled_version(username , password , required_data_list ,AssessId=None, assess_period_data=None ,range1='' , range2=''):
            update.message.reply_text("هل تريد تعبئة علامات صف اخر؟ نعم | لا",reply_markup=ReplyKeyboardMarkup([['نعم', 'لا']], one_time_keyboard=True))
            return WAITING_FOR_RESPONSE            
        # End of conversation
        update.message.reply_text("تمام انتهينا")
        return ConversationHandler.END

def init_receive(update, context):
    update.message.reply_text("هل تريد تفريغ العلامات على المنظومة و تعبئة سجل العلامات الرسمي من السجل الالكتروني ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def check_creds(update, context):
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        user = update.message.from_user
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        # update.message.reply_text("Thanks for sharing! You're a credentials user {} and password {}.".format(context.user_data['creds'][0], context.user_data['creds'][1] ) )
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
            return CREDS                                    
        else:
            print('moving to next function')
            update.message.reply_text('0000ارسل ملف العلامات الجانبي الالكتروني؟')            
            return FILE

def init_performance_side_marks(update, context):
    update.message.reply_text("هل تريد كشف العلامات والاداء الجانبي ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def send_performance_side_marks_note_doc(update, context):
    user = update.message.from_user
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        update.message.reply_text("انتظر لحظة لو سمحت")         
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            logger.info(f"username:{username} ----> password :{password}")
            side_marks_document(username, password)
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END

def init_certs (update, context): 
    update.message.reply_text("هل تريد طباعة شهادات الطلاب ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS_2

def send_students_certs(update, context):
    user = update.message.from_user
    term = update.message.text.replace('/','')    
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        # update.message.reply_text("Thanks for sharing! You're a credentials user {} and password {}.".format(context.user_data['creds'][0], context.user_data['creds'][1] ) )
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            logger.info(f"username:{username} ----> password :{password}")
            update.message.reply_text("انتظر لحظة لو سمحت")
            if term == 'term1':            
                create_certs_wrapper(username, password)
                pass
            elif term == "term2":
                create_certs_wrapper(username, password , term2=True)
                pass
            else:
                func_text = '''اختر الفصل بالضغط عليه
                /term1 الفصل الاول 
                /term2 الفصل الثاني '''
                
                update.message.reply_text("لم تختر فصل ")
                
                update.message.reply_text(func_text)                                 
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END

def init_tables (update, context): 
    update.message.reply_text("هل تريد طباعة جداول علامات الطلاب ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def send_students_tables(update, context):
    term = update.message.text.replace('/','')    
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            logger.info(f"username:{username} ----> password :{password}")
            update.message.reply_text("انتظر لحظة لو سمحت")
            if term == 'term1':            
                create_tables_wrapper(username, password)
            elif term == "term2":
                create_tables_wrapper(username, password, term2=True)
            else:
                func_text = '''اختر الفصل بالضغط عليه
                /term1 الفصل الاول 
                /term2 الفصل الثاني '''
                
                update.message.reply_text("لم تختر فصل ")
                
                update.message.reply_text(func_text)                                 
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END        

def init_official_marks(update, context):
    update.message.reply_text("هل تريد سجل علامات رسمي ؟ \n قم باعطائي اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def send_official_marks_doc(update, context):
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        user = update.message.from_user
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        # update.message.reply_text("Thanks for sharing! You're a credentials user {} and password {}.".format(context.user_data['creds'][0], context.user_data['creds'][1] ) )
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            logger.info(f"username:{username} ----> password :{password}")
            update.message.reply_text("انتظر لحظة لو سمحت") 
            fill_official_marks_functions_wrapper_v2(username, password)
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END

def init_e_side_marks(update, context):
    update.message.reply_text("هل تريد كشف علامات جانبي الكتروني ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def send_e_side_marks_note_doc(update, context):
    user = update.message.from_user
    session = requests.Session()
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        update.message.reply_text("انتظر لحظة لو سمحت")         
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            logger.info(f"username:{username} ----> password :{password}")
            create_e_side_marks_doc(username, password , session =session)
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END

# Run the program
if __name__ == '__main__':
    setup_logging("access.log")
    parser = argparse.ArgumentParser(description='A script with command-line arguments.')
    # Define command-line arguments
    parser.add_argument('-t', '--test', action='store_true', help='Enable testing')
    
    # Parse the command-line arguments
    args = parser.parse_args()
    
    if args.test:
        token = test_bot
    else:
        token = production_bot
    
    check_if_send_folder_exist()
    updater = Updater(token, use_context=True)
    dp = updater.dispatcher

    bot = Bot(token=token)
    
    # Commands
    dp.add_handler(CommandHandler('help', help_command))
    dp.add_handler(CommandHandler('start', start))
    dp.add_handler(CommandHandler('links', links))
    dp.add_handler(CommandHandler('test_qt', test_qt))
    
    # Log all errors
    dp.add_error_handler(error)

    send_side_marks_note_doc_conv = ConversationHandler(
        entry_points=[CommandHandler('side_marks_note', init_side_marks)],
                                        states={
                                            CREDS_2 : [MessageHandler(Filters.text ,which_term)],
                                            ASK_QUESTION : [MessageHandler(Filters.text , send_side_marks_note_doc)]
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )
    
    fill_assess_arbitrary_marks_conv = ConversationHandler(
        entry_points=[CommandHandler('fill_assess_arbitrary', init_fill)],
                                        states={
                                            CREDS: [MessageHandler(Filters.text & ~Filters.command, print_available_assessments)],
                                            AVAILABLE_ASS: [MessageHandler(Filters.text , fill_assess_arbitrary)],
                                            WAITING_FOR_RESPONSE: [MessageHandler(Filters.text, handle_response)],
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)],
                                        # allow_reentry=True  # To allow restarting the conversation while it's in progress
                                                        ) 

    fill_assess_arbitrary_empty_marks_conv = ConversationHandler(
        entry_points=[CommandHandler('empty_assess', init_empty_fill)],
                                        states={
                                            CREDS: [MessageHandler(Filters.text & ~Filters.command, print_available_assessments)],
                                            AVAILABLE_ASS: [MessageHandler(Filters.text , fill_assess_empty)],
                                            WAITING_FOR_RESPONSE: [MessageHandler(Filters.text, handle_response)],
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)],
                                        allow_reentry=True  # To allow restarting the conversation while it's in progress
                                                        ) 

    send_performance_side_marks_note_doc_conv = ConversationHandler(
        entry_points=[CommandHandler('performance_side_marks_note', init_performance_side_marks)],
                                        states={
                                            CREDS : [MessageHandler(Filters.text , send_performance_side_marks_note_doc)]
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

    send_students_certs_conv = ConversationHandler(
                                        entry_points=[CommandHandler('certs', init_certs)],
                                        states={
                                            CREDS_2 : [MessageHandler(Filters.text ,which_term)],
                                            ASK_QUESTION : [MessageHandler(Filters.text , send_students_certs)]
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

    send_students_tables_conv = ConversationHandler(
                                        entry_points=[CommandHandler('tables', init_tables)],
                                        states={
                                            CREDS_2 : [MessageHandler(Filters.text ,which_term)],
                                            ASK_QUESTION : [MessageHandler(Filters.text , send_students_tables)]
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )
    
    send_official_marks_doc_conv = ConversationHandler(
                                        entry_points=[CommandHandler('official_marks', init_official_marks)],
                                        states={
                                            CREDS : [MessageHandler(Filters.text , send_official_marks_doc)]
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

    send_official_marks_doc_conv_offline = ConversationHandler(
                                        entry_points=[CommandHandler('official_marks_offline', init_receive)],
                                        states={
                                            CREDS : [MessageHandler(Filters.text , check_creds )],
                                            FILE : [MessageHandler(Filters.text ,receive_file)]
                                                },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

    send_students_absent_doc_conv = ConversationHandler(
                                    entry_points=[CommandHandler('student_absent_document', init_absent_doc)],
                                        states={
                                            CREDS : [MessageHandler(Filters.text ,send_absent_notebook_doc)],
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )
    
    send_e_side_marks_note_doc_conv = ConversationHandler(
                                    entry_points=[CommandHandler('e_side_marks_note', init_e_side_marks)],
                                        states={
                                            CREDS : [MessageHandler(Filters.text , send_e_side_marks_note_doc)]
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

    receive_file_handler_conv = ConversationHandler(
                                    entry_points=[MessageHandler(Filters.document, receive_file)],
                                        states={
                                            ASK_QUESTION: [MessageHandler(Filters.text, handle_question)], 
                                            CREDS_2: [MessageHandler(Filters.text & ~Filters.command, upload_marks_bot_version)],
                                            # CREDS_2: [MessageHandler(Filters.text & ~Filters.command, print_available_assessments)],
                                                },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

    check_five_names_marks_conv = ConversationHandler(
                                    entry_points=[CommandHandler('check_marks', init_check_five_names_marks)],
                                        states={
                                            CREDS_2 : [MessageHandler(Filters.text , which_term)],
                                            ASK_QUESTION : [MessageHandler(Filters.text , print_check_five_names_marks)],
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

    marks_up_percentage_conv = ConversationHandler(
                                    entry_points=[CommandHandler('marks_up_percentage', init_marks_up_percentage)],
                                        states={
                                            CREDS_2 : [MessageHandler(Filters.text , which_term)],
                                            ASK_QUESTION : [MessageHandler(Filters.text , get_marks_up_percentage)],
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )
    
    # Add the conversation handler to the dispatcher
    dp.add_handler(send_official_marks_doc_conv)
    dp.add_handler(send_e_side_marks_note_doc_conv)
    dp.add_handler(send_side_marks_note_doc_conv)
    dp.add_handler(fill_assess_arbitrary_marks_conv)
    dp.add_handler(fill_assess_arbitrary_empty_marks_conv)
    dp.add_handler(send_official_marks_doc_conv_offline)
    dp.add_handler(receive_file_handler_conv)
    dp.add_handler(send_students_certs_conv)
    dp.add_handler(send_students_tables_conv)
    dp.add_handler(check_five_names_marks_conv)
    dp.add_handler(marks_up_percentage_conv)
    dp.add_handler(send_performance_side_marks_note_doc_conv)
    dp.add_handler(send_students_absent_doc_conv)

    # Run the bot
    updater.start_polling(1.0)
    updater.idle()