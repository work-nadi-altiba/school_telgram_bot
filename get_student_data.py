import csv
import re
from bs4 import BeautifulSoup
import requests
import json

def soup( html_doc , select_one : str= None ,select : str= None ,select_first : str= None):

    if select_one:
        soup = BeautifulSoup(html_doc, 'html.parser')
        return soup.select_one(select_one) 

    if select:
        soup = BeautifulSoup(html_doc, 'html.parser')
        return soup.select(select)

    if select_first:
        soup = BeautifulSoup(html_doc, 'html.parser')
        try :
            return soup.select(select_first)[0].text
        except : 
            return ''


def mreq2 (page ,cook):

    cookie = cook

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Referer': 'https://emis.moe.gov.jo/openemis-core/Institution/Institutions/eyJpZCI6MjYwMCwiNWMzYTA5YmYyMmUxMjQxMWI2YWY0OGRmZTBiODVjMmQ5ZDExODFjZDM5MWUwODk1NzRjOGNmM2NhMWU1ZTRhZCI6Impuam1sM29hcjRjN211NWJvcm1sM2tsdWcyIn0.MGQ1NGVkNjE0OGZlM2Y5ZGUzZTNlZjUzYTU3ODgzOTY3YjgyNTYzOWU1MzdlYjljOTEyZTM5MWFlMDUyODNlNA/Students/index?academic_period_id=13&status_id=1&education_grade_id=-1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.52',
        'sec-ch-ua': '"Chromium";v="106", "Microsoft Edge";v="106", "Not;A=Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    params = {
        'academic_period_id': '13',
        'status_id': '1',
        'education_grade_id': '280',
        'page': page,
    }

    response = requests.get('https://emis.moe.gov.jo/openemis-core/Institution/Institutions/Students/index', params=params, cookies=cookie, headers=headers)
    return response.text 

def html2text(html):        
    soup = BeautifulSoup(f'{html}' , features="html.parser")
    return soup.get_text()

def req (cook , link ):
    cookie = cook
    response = requests.get(link, cookies=cookie)
    return response.text

cookies = {
    'System': 'Q2FrZQ%3D%3D.NzA1MmFkZmIxMGQ5NDNjNTczNjRmZTk2MjA3ZGE0OWYyMzQyMGY5MjQ0OWE4ODY2Y2Y1MWNmMzVkMjU1ZmRlNcgPVMbanG88tt3VKpZEHy9y761CjPYS%2BHMomzF8jLloTwLRXYxiv3dRG79bNfviwQ%3D%3D',
    'csrfToken': '02bb6b45d45a4d0f3372d18bd2c4ca4241ddfb16251188af9bbdea12663d18a26fdce0bd1bc0b75b1aa3192ad31fbe329c1007bded2f3b11cbdae4a6ed97c541',
    '_ga': 'GA1.3.827323031.1669410400',
    '_gid': 'GA1.3.738795208.1670617527',
    'PHPSESSID': 'fanjda13qs0p7ig7a5taereqr7',
    'SRVNAME': 'S3|Y5UiP|Y5UTP',
    '_gat': '1',
}


dataRange= soup(mreq2(2 , cookies) , select='.pagination > li >a')
lis = [ html2text(i) for i in dataRange  if len(i) != 0 ]
data=[]

breakpoint()
# print(mreq2(1 , cookies))
'''for loop to get all the pages'''
for i in range(1,int(lis[-1])+1):
    print(i)
    res = mreq2(i , cookies)
    name = soup(res , select='table > tbody > tr > td:nth-child(3)')
    Class = soup(res , select='table > tbody > tr > td:nth-child(5)')
    link = soup(res , select='.dropdown-menu.action-dropdown>li:nth-child(2) > a')
    # breakpoint()
    for x in range(len(name)):
        res2 = req(cookies , f'https://emis.moe.gov.jo{html2text(link[x]["href"])}')
        # breakpoint()
        emis_id = soup(res2 , select_first='.panel-body>:nth-child(4)> .form-input')
        name1 = soup(res2 , select_first='.panel-body>:nth-child(5)> .form-input')
        name2 = soup(res2 , select_first='.panel-body>:nth-child(6)> .form-input')
        name3 = soup(res2 , select_first='.panel-body>:nth-child(7)> .form-input')
        name4 = soup(res2 , select_first='.panel-body>:nth-child(8)> .form-input')
        birth1 = soup(res2 , select_first='.panel-body>:nth-child(28)> .form-input')
        birth2 = soup(res2 , select_first='.panel-body>:nth-child(29)> .form-input')
        birth_date= soup(res2 , select_first='.panel-body>:nth-child(11)> .form-input')
        nationality = soup(res2 , select_first='.table>tbody>tr>td:nth-child(3)')
        gender = soup(res2 , select_first='.panel-body>:nth-child(10)> .form-input')
        resedent1 = soup(res2 , select_first='.panel-body>:nth-child(21)> .form-input')
        resedent2 = soup(res2 , select_first='.panel-body>:nth-child(22)> .form-input')
        resedent3 = soup(res2 , select_first='.panel-body>:nth-child(23)> .form-input')


        row = {"SN" : f"{html2text(name[x])}" , "SC" : f"{html2text(Class[x] )}" ,"emis_id " : f"{emis_id  }" , "name1" : f"{name1  }" , "name2 " : f"{name2  }","name3 " : f"{name3  }","name4 " : f"{name4  }","birth1 " : f"{birth1  }","birth2 " : f"{birth2  }","birth_date" : f"{birth_date}","nationality " : f"{nationality  }","gender " : f"{gender  }","resedent1 " : f"{resedent1  }","resedent2 " : f"{resedent2  }","resedent3 " : f"{resedent3  }" }

        data.append(row) 
        # row = {"SN" :'' , "SC" :'' ,"emis_id " : '' , "name1" : '' , "name2 " : '',"name3 " : '',"name4 " : '',"birth1 " : '',"birth2 " : '',"birth_date" : '',"nationality " : '',"gender " : '',"resedent1 " : '',"resedent2 " : '',"resedent3 " : '' } 
        # breakpoint()

def write_csv(data):
    with  open('data3.json' , 'w') as f:
        json.dump(data, f)

write_csv(data)
import winsound
duration = 1000  # milliseconds
freq = 440  # Hz
winsound.Beep(freq, duration)
winsound.Beep(freq, duration)
winsound.Beep(freq, duration)
winsound.Beep(freq, duration)

breakpoint()

