# pip install python-telegram-bot
# pip install python-telegram-bot==13.7

from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, ConversationHandler
from telegram.ext import *
from telegram import Bot
from utils1 import *
import keys
import io

print('Starting up bot...')

# # fill assess arbitrary marks conversation handler stats
INIT_F , RESPOND = range(2)
# # send side marks document conversation handler stats
# INIT_SM , SEND_SN = range(1)
# # send students certs conversation handler stats
# INIT_C , SEND_SC = range(1)
# #send official marks document conversation handler stats
# INIT_OM , SEND_OM = range(1)
# CREDS = range(1)
CREDS, AVAILABLE_ASS  = range(2)
CREDS, FILE = range(2)
# TODO: make sure of every fallback function (cancle function)in the handler conversation 

# Define a function to handle incoming files

def receive_file(update, context ):
    '''
        dispatcher.add_handler(MessageHandler(Filters.document, receive_file))
    '''
    
    # Check if the message contains a document
    if not update.message.document:
        update.message.reply_text('Please send an Excel file.')
        return

    # Get the file object and read its content
    file_obj = context.bot.get_file(update.message.document.file_id)
    file_bytes = io.BytesIO(file_obj.download_as_bytearray())
    # file_content = file_bytes.read()
    fill_official_marks_doc_wrapper_offline(9971055725,9971055725,Read_E_Side_Note_Marks(file_content=file_bytes))
    files = count_files()
    chat_id = update.message.chat.id
    send_files(bot, chat_id, files)
    delete_send_folder()

    update.message.reply_text('تم بنجاح')

def start(update, context):
    context.bot.send_message(chat_id=update.effective_chat.id, text="/side_marks_note لطباعة ملف العلامات الجانبي \n /certs لطباعة ملف الشهادات \n /official_marks لطباعة ملف العلامات الرسمية \n /fill_assess_arbitrary لتسجيل العلامات العشوائية' \n /cancel لألغاء العملية")

def send_files(bot, chat_id, files):
    for file in files:
        bot.send_document(chat_id=chat_id, document=open(file, 'rb'))

# Lets us use the /help command
def help_command(update, context):
    update.message.reply_text('/side_marks_note لطباعة ملف العلامات الجانبي \n /certs لطباعة ملف الشهادات \n /official_marks لطباعة ملف العلامات الرسمية \n /fill_assess_arbitrary لتسجيل العلامات العشوائية')

# Log errors
def error(update, context):
    print(f'Update {update} caused error {context.error}')

def cancel(update, context):
    user = update.message.from_user
    update.message.reply_text("تم ")
    return ConversationHandler.END

def get_user_creds(update, context):
    # code
    pass

def check_user_creds(update, context):
    # code
    pass

def init_fill(update, context):
    update.message.reply_text("هل تريد تسجل علامات عشوائي ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def print_available_assessments(update, context):
    user = update.message.from_user
    context.user_data['creds'] = update.message.text.split('/')
    username = context.user_data['creds'][0]
    password = context.user_data['creds'][1]
    print(username, password)
    if get_auth(username, password) == False:
        update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
    else:
        update.message.reply_text("انتظر لحظة لو سمحت") 
        auth = get_auth(username,password)
        # TODO: handle empty editable_assessments list
        editable_assessments = get_editable_assessments(auth ,username)
        data_to_enter_marks = get_required_data_to_enter_marks(auth ,username)
        string = assessments_commands_text(editable_assessments)
        update.message.reply_text(string)
        context.user_data['assessments'] = editable_assessments
        context.user_data['data_to_enter_marks'] = data_to_enter_marks
        return  AVAILABLE_ASS
    
def fill_assess_arbitrary(update, context):
    user = update.message.from_user
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        code = update.message.text.replace('/','')
        editable_assessments = context.user_data['assessments'] 
        data_to_enter_marks = context.user_data['data_to_enter_marks']  
        username , password = context.user_data['creds'][0] , context.user_data['creds'][1]
        assess_data = [i for i in editable_assessments if i.get('code') == code][0]
        wanted_grades = [i for i in data_to_enter_marks if i.get('assessment_id') == assess_data['gradeId']]
        enter_marks_arbitrary_controlled_version(username,password,wanted_grades,assess_data['pass_mark'],assess_data['max_mark'])    
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
        
def init_side_marks(update, context):
    update.message.reply_text("بدك اعطيك كشف علامات جانبي ؟ \n اعطيني اسم المستخدم و كلمة السر من فضلك ؟ \n مثلا 9981058924/123456") 
    return CREDS

def send_side_marks_note_doc(update, context):
    user = update.message.from_user
    if update.message.text == '/cancel':
        return cancel(update, context)
    else:
        context.user_data['creds'] = update.message.text.split('/')
        username = context.user_data['creds'][0]
        password = context.user_data['creds'][1]
        # update.message.reply_text("Thanks for sharing! You're a credentials user {} and password {}.".format(context.user_data['creds'][0], context.user_data['creds'][1] ) )
        print(username, password)
        if get_auth(username, password) == False:
            update.message.reply_text("اسم المستخدم او كلمة السر خطأ") 
        else:
            side_marks_document(username, password)
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END

def init_certs (update, context): 
    # code
    pass

def send_students_certs(update, context):
    # code
    pass

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
            update.message.reply_text("انتظر لحظة لو سمحت") 
            fill_official_marks_doc_wrapper(username, password)
            files = count_files()
            chat_id = update.message.chat.id
            send_files(bot, chat_id, files)
            delete_send_folder()
            return ConversationHandler.END


# Run the program
if __name__ == '__main__':
    updater = Updater(keys.token, use_context=True)
    dp = updater.dispatcher

    bot = Bot(token=keys.token)
    
    # Commands
    dp.add_handler(CommandHandler('help', help_command))
    dp.add_handler(CommandHandler('start', start))

    # Log all errors
    dp.add_error_handler(error)

    fill_assess_arbitrary_marks_conv = ConversationHandler(
        entry_points=[CommandHandler('fill_assess_arbitrary', init_fill)],
                                        states={
                                            CREDS: [MessageHandler(Filters.text & ~Filters.command, print_available_assessments)],
                                            AVAILABLE_ASS: [MessageHandler(Filters.text , fill_assess_arbitrary)],
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)],
                                        allow_reentry=True  # To allow restarting the conversation while it's in progress
                                                        ) 

    send_side_marks_note_doc_conv = ConversationHandler(
        entry_points=[CommandHandler('side_marks_note', init_side_marks)],
                                        states={
                                            CREDS : [MessageHandler(Filters.text , send_side_marks_note_doc)]
                                        },
                                        fallbacks=[CommandHandler('cancel', cancel)]
                                                        )

#     send_students_certs_conv = ConversationHandler(
#     entry_points=[CommandHandler('شهادات', start)],
#     states={
#         NAME: [MessageHandler(Filters.text, name)],
#         AGE: [MessageHandler(Filters.regex('^(Less than 18|Between 18 and 30|More than 30)$'), age)],
#         GENDER: [MessageHandler(Filters.regex('^(Male|Female|Other)$'), gender)]
#     },
#     fallbacks=[CommandHandler('انهاء', cancel)]
# )

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
    # send_students_absent_doc_conv = ConversationHandler(
        
        

    # Add the conversation handler to the dispatcher
    dp.add_handler(send_side_marks_note_doc_conv)
    dp.add_handler(send_official_marks_doc_conv)
    dp.add_handler(fill_assess_arbitrary_marks_conv)
    dp.add_handler(send_official_marks_doc_conv_offline)
    dp.add_handler(MessageHandler(Filters.document, receive_file))

    # Run the bot
    updater.start_polling(1.0)
    updater.idle()
    
    
    
    