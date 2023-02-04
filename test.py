import os
import subprocess
import pyshark
import requests
import json
from pygments import highlight
from pygments.lexers import JsonLexer
from pygments.formatters import TerminalFormatter

def curl_requests (url):
    cookies = {
        'System': 'Q2FrZQ%3D%3D.MWIxZDEyYzNhY2ZkNzJkNTg2NjVlYTQ2OGY0YmU5OWM1MGY0MzcwYTVhOTcyMGQwZDY5ZWNhOTQwM2Y2YWQzYo4YKTorv8yPGrgOor5UCzcQOmz%2F64ni0XSpSQolV9GvUvjax2jbT6oXmuDilod7iQ%3D%3D',
        'csrfToken': 'b59abd566347553e791bca4801792a6f24c0c419c1ccb0b177a72b313a0c0b98e4043d498655aef7501a91a188086542f0cffe71a330543ae42f05f992afb256',
        'TS01a6a8a1': '012cb8ec883f22c26d19ddc996f7754fd9dcc162e0e5e6c510c0e8d7c4396769598f78a0d5114b43c6b74be727d21a091b11ee46ad6b48a2a5a1223134e1cb995f9d4ff8e4bda2b983358723720546518d22e0d22e',
        '_ga': 'GA1.3.209686253.1674347858',
        'PHPSESSID': 'nk2mu8gpruj1356opsvutc74m2',
        'BIGipServeremis-SRV-Pool': '1946265792.20480.0000',
        'TS01149972': '012cb8ec883f8c5654b67965a13dd3c4a72de331dad2397b5c4a69eb9345c313ca02a3cd72416c6dce1c0e4a147d705776a9464640632db7735719df9105db768fdaff4aac0cec297efe721f800902beec4d84f0d5',
        '_gid': 'GA1.3.1364936227.1675380900',
        'TS01149972030': '01cd72cd0ab2db119af02ec1f1dcb909f72d01d7c019fc904e73d6c979712ac9e79b3eabed3bcb46ee49c1eb3f7a45f38e7bff9861',
        '_gat': '1',
    }

    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:102.0) Gecko/20100101 Firefox/102.0',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'en-US,en;q=0.5',
        'Referer': 'https://emis.moe.gov.jo/openemis-core/Institution/Institutions/eyJpZCI6MjYwMCwiNWMzYTA5YmYyMmUxMjQxMWI2YWY0OGRmZTBiODVjMmQ5ZDExODFjZDM5MWUwODk1NzRjOGNmM2NhMWU1ZTRhZCI6Im5rMm11OGdwcnVqMTM1Nm9wc3Z1dGM3NG0yIn0.MGE4YWRlNDExZDI0NGYzYjczYTUzNDJkYjRkYjM1ZGYyZTQ4ZDNjZjhhNjZmZDkyZDk3MmQxZTI4ZjgxNDk4ZA/StudentAttendances/index',
        'Connection': 'keep-alive',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'no-cors',
        'Sec-Fetch-Site': 'same-origin',
        'ControllerAction': 'StudentAttendances',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
    }

    response = requests.get(
        url,
        cookies=cookies,
        headers=headers,
    )
    
    print( response.text , sep='\n')
    
def filter_api_from_pcap():
    
    cap = pyshark.FileCapture(
        'superemis_apis/superemis.pcapng', override_prefs={'ssl.keylog_file': os.path.abspath('superemis_apis/sslkeylog.log')} ,display_filter='http')

    allPkts = [pkt for pkt in cap]
    counter = 0
    for packet in allPkts:
        try : 
            print(packet.http.request_uri)
            save_to_file(packet.http.request_uri+'\n')
        except:
            pass  
    
def save_to_file(text):
    with open('superemis_apis/requests2.txt', 'a+') as f :
        f.write(text)
        f.close()
        
def read_from_file(file):
    with open(file, 'r') as f :
        return f.read()
        f.close()
def curl_loop():
    for url in read_from_file('superemis_apis/requests.txt').split('\n'):
        curl_requests(url)
        print(url)
        input('Press Enter to continue...')

def curl_api_url(Authorization , url):
    headers = {"Authorization": Authorization}
    response = requests.request("GET", url, headers=headers)
    return(response.json())
        
def my_jq(data):
    # json_object = json.loads(data)
    json_str = json.dumps(data, indent=4, sort_keys=True)
    print(json_str)
    # print(highlight(json_str, JsonLexer(), TerminalFormatter())) 
    
def login(username=9971055725 , password=9971055725):
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
        # print ('Logged in successfully')
        # print (response.json()['data']['token'])
        return response.json()['data']['token']
               
def sort_api_help(url):
    '''
    دالة تساعدني في توثيق الروابط و كتابة اسماء الدوال و حفطها في ملف
    '''
    auth = login()
    print(curl_api_url(auth, url))
    my_jq(curl_api_url(auth, url))
    # subprocess.run(f'''{str(curl_api_url(auth, url))} | jq . ''')
    # my_jq(f'''{curl_api_url(auth, url)}''')

def main():
    sort_api_help('https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Staff?_limit=1&_contain=Institutions&_fields=Institutions.code,Institutions.id,Institutions.name')
    # login()
    # curl_loop()   
    # filter_api_from_pcap()
    # save_to_file(filter_api_from_pcap())

if __name__ == "__main__":
    main()