
import requests

def login(username, password):
    ' دالة تسجيل الدخول للحصول على الرمز الخاص بالتوكن و يستخدم في header Authorization'
    url = "https://emis.moe.gov.jo/openemis-core/oauth/login"

    payload = {
        "username": username,
        "password": password
    }
    response = requests.request("POST", url, data=payload )

    # print(response.json()['data']['message'])
    if response.json()['data']['message'] == 'Invalid login creadential':
        print ('Invalid login creadential')
    else: 
        print ('Logged in successfully')
        print (response.json()['data']['token'])
        return response.json()['data']['token']
    
login()