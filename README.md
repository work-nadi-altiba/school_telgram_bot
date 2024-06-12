# school_telgram_bot

بوت تلجرام يقوم بالاعمال الروتينية للمعلم كادخال العلامات و طباعة الكشف الجانبي و سجل العلامات الرسمي و دفتر الحضور و الغياب و طباعة الشهادات


A telegram robot that performs the routine work of the teacher, such as entering grades, printing the side list, the official grade record, the attendance and absence book, and printing certificates.

```sh
git clone https://github.com/Anas-jaf/school_telgram_bot.git
```

يجب ان تستخدم الكود على سيرفر لينكس و لتنزيل المتطلبات استعمل هذا الامر

you should use this code in linux machine/server and , to install the requirements use this command 

`sudo apt install python3-odf libcurl4-openssl-dev libssl-dev python3-docxtpl python3-fitz libreoffice ttf-mscorefonts-installer ; pip install docx2pdf num2words python-telegram-bot==13.7 hijri_converter fitz PyMuPDF ezodf PyPDF4 wfuzz pygments openpyxl webcolors PyPDF2~=2.0 requests[socks] python-decouple`

لتغير لغة الارقم من العربية الى الهندية CTLTextNumerals من صفر الى واحد
```
~/.config/libreoffice/<version>/user/registrymodifications.xcu
```

bot link 
https://t.me/sitToTellYou_bot
