import csv
import re
from bs4 import BeautifulSoup
import requests
import json


def soup( html_doc , select_one : str= None ,select : str= None ):

    if select_one:
        soup = BeautifulSoup(html_doc, 'html.parser')
        return soup.select_one(select) 
    if select:
        soup = BeautifulSoup(html_doc, 'html.parser')
        return soup.select(select) 

def html2text(html):        
    soup = BeautifulSoup(f'{html}' , features="html.parser")
    return soup.get_text()

def req():
    pass

def mreq1 ():
    cookies = {
        'csrfToken': '30fada9209b3b40c8b65668cfef602334753e71c113e2da1e27e48b0e629792d2d50eb6373d53cec030cb725a3fb1a8d2ea78de9e23d8fa29e19cd0367b0090b',
        'System': 'Q2FrZQ%3D%3D.MjhkMzA2ZGMyNjAyZGVhYTBhZDA5MjBlNzBkODc5OGRhMmE3NzRhYWIwOTg2MmZjZDgzYmUwMTU4NDlmZDczNVk74hY5XJla4hdfxiLsHcK7FjIffkBQ3s0q3aX4caVItv2JMTNU81PtK8aYUmWFCA%3D%3D',
        '_ga': 'GA1.3.739379032.1666302696',
        '_gid': 'GA1.3.1191564847.1666302696',
        'PHPSESSID': '7hq698iojo8g2ae935ceq9mjb4',
        '_gat': '1',
        'SRVNAME': 'S2|Y1LzW|Y1LyT',
    }

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Language': 'en-US,en;q=0.9',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Content-Type': 'multipart/form-data; boundary=----WebKitFormBoundarygOBnFEABwmykO3UZ',
        # Requests sorts cookies= alphabetically
        # 'Cookie': 'csrfToken=30fada9209b3b40c8b65668cfef602334753e71c113e2da1e27e48b0e629792d2d50eb6373d53cec030cb725a3fb1a8d2ea78de9e23d8fa29e19cd0367b0090b; System=Q2FrZQ%3D%3D.MjhkMzA2ZGMyNjAyZGVhYTBhZDA5MjBlNzBkODc5OGRhMmE3NzRhYWIwOTg2MmZjZDgzYmUwMTU4NDlmZDczNVk74hY5XJla4hdfxiLsHcK7FjIffkBQ3s0q3aX4caVItv2JMTNU81PtK8aYUmWFCA%3D%3D; _ga=GA1.3.739379032.1666302696; _gid=GA1.3.1191564847.1666302696; PHPSESSID=7hq698iojo8g2ae935ceq9mjb4; _gat=1; SRVNAME=S2|Y1LzW|Y1LyT',
        'Origin': 'https://emis.moe.gov.jo',
        'Referer': 'https://emis.moe.gov.jo/openemis-core/Institution/Institutions/eyJpZCI6MjYwMCwiNWMzYTA5YmYyMmUxMjQxMWI2YWY0OGRmZTBiODVjMmQ5ZDExODFjZDM5MWUwODk1NzRjOGNmM2NhMWU1ZTRhZCI6IjdocTY5OGlvam84ZzJhZTkzNWNlcTltamI0In0.Y2FmN2VhMGU2NzJiNWVkMmE2OGFkN2VjNzg0YzBlMjJmZTc3NGE3OWFjNDdkYjdkMDhkOWZmMjk1NTRlMjc1Yg/Students/index',
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
        'education_grade_id': '-1',
    }

    data = '------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="_method"\r\n\r\nPOST\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="_csrfToken"\r\n\r\n30fada9209b3b40c8b65668cfef602334753e71c113e2da1e27e48b0e629792d2d50eb6373d53cec030cb725a3fb1a8d2ea78de9e23d8fa29e19cd0367b0090b\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][hasMany][first_name]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][hasMany][middle_name]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][hasMany][third_name]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][hasMany][last_name]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][hasMany][contact_number]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][hasMany][identity_type]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][hasMany][identity_number]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="AdvanceSearch[Students][isSearch]"\r\n\r\n\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="academic_period"\r\n\r\n13\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="education_grade"\r\n\r\n-1\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="student_status"\r\n\r\n1\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="Search[limit]"\r\n\r\n6\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="_Token[fields]"\r\n\r\n6c02c940b97ad894335399393267c28480079a53%3A\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ\r\nContent-Disposition: form-data; name="_Token[unlocked]"\r\n\r\nAdvanceSearch%7CSearch.searchField%7Creset\r\n------WebKitFormBoundarygOBnFEABwmykO3UZ--\r\n'

    response = requests.post('https://emis.moe.gov.jo/openemis-core/Institution/Institutions/eyJpZCI6MjYwMCwiNWMzYTA5YmYyMmUxMjQxMWI2YWY0OGRmZTBiODVjMmQ5ZDExODFjZDM5MWUwODk1NzRjOGNmM2NhMWU1ZTRhZCI6IjdocTY5OGlvam84ZzJhZTkzNWNlcTltamI0In0.Y2FmN2VhMGU2NzJiNWVkMmE2OGFkN2VjNzg0YzBlMjJmZTc3NGE3OWFjNDdkYjdkMDhkOWZmMjk1NTRlMjc1Yg/Students/index', params=params, cookies=cookies, headers=headers, data=data)
    return response.content

def mreq2 (page):
    cookies = {
        'csrfToken': '126fd302f955eae55a4546717a1d4247be1fd9e619140d5a640d57189e23028763cefdbbfcd4dd833f6b6d83aac4bccbe234c4e1265041eb574f7039f54e8811',
        'System': 'Q2FrZQ%3D%3D.ZTUyMTBhYjU2NWU3ZjUxYzBkNGIyMjJmYTM0N2ZhMmI0ZWRlZjNlMmUzNGZhMTNiZjU2ZDUyMGFhODNiMzI1YXjmAfRI1vcTg%2F%2Bm4b2K5ve9%2FYVemlI%2BXJvROZwD6AP9Zg3c08ZXWqvFs1Bp8IXzGQ%3D%3D',
        '_ga': 'GA1.3.739379032.1666302696',
        '_gid': 'GA1.3.1191564847.1666302696',
        'PHPSESSID': '9fbje2im45o60gtmkhnbkcgn11',
        'SRVNAME': 'S6|Y1RYW|Y1RQc',
    }

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        # Requests sorts cookies= alphabetically
        # 'Cookie': 'csrfToken=30fada9209b3b40c8b65668cfef602334753e71c113e2da1e27e48b0e629792d2d50eb6373d53cec030cb725a3fb1a8d2ea78de9e23d8fa29e19cd0367b0090b; System=Q2FrZQ%3D%3D.NmU5MzljNmZhMmMzZWZkMTAzYWRhZDFlMWI3YjE5MTNmOWYxZmFkNzE5ZDg4YjYzZTA2YTkzNzM1NDE4YmQ5NzANFnssWU0VPs6KEG5sQGERwivggJrSocxJzNrnLyRbcD2MyvmaA0cg77QtmSSCTQ%3D%3D; _ga=GA1.3.739379032.1666302696; _gid=GA1.3.1191564847.1666302696; PHPSESSID=jnjml3oar4c7mu5borml3klug2; SRVNAME=S2|Y1LZt|Y1LON',
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
        'education_grade_id': '-1',
        'page': page,
    }

    response = requests.get('https://emis.moe.gov.jo/openemis-core/Institution/Institutions/Students/index', params=params, cookies=cookies, headers=headers)
    return response.content    
# print(mreq1()) 
# for i in range (1 , 4):
#     mreq1(i)

def write2file(cont):
    with open("doc.html" , "wb+" ,encoding="utf8") as f:
        f.write(cont)

def write2file2(cont , file):
    with open(str(file) , "w+",encoding="utf8") as f:
        f.writelines(cont)

def readFile():
    with open("doc1.html" , 'r' ,encoding="utf8") as f:
        cont = f.readlines()
    cont =''.join(cont)
    return cont

# mark =soup(cont , '.counter')
# mark =html2text(mark)
# print(re.findall('\d+',mark ))

dataRange= soup(mreq2(2) , select='.pagination > li >a')
lis = [ html2text(i) for i in dataRange  if len(i) != 0 ]
data=[]

for i in lis:
    print(lis)
    res = mreq2(i)
    name = soup(res , select='table > tbody > tr > td:nth-child(3)')
    Class = soup(res , select='table > tbody > tr > td:nth-child(5)')
    # print(name , Class , sep="\n")
    # input('press anything')
    for x in range(len(name)):
        print( html2text(name[0]) , html2text(Class[0] ))    
        row = {"SN" : f"{html2text(name[x])}" , "SC" : f"{html2text(Class[x] )}"}
        data.append(row)        

write2file2(json.dumps(data) , 'hisSchoolStudents.json')

def readFile2(file):
    with open(file , 'r' ,encoding="utf8") as f:
        cont = f.readlines()
    cont =''.join(cont)
    return cont
x = readFile2('hisSchoolStudents.json')

x = json.loads(x)

f = csv.writer(open("hisSchoolStudents.csv", "w+" , newline='',encoding="utf8"))

# Write CSV Header, If you dont need that, remove this line
f.writerow(["اسم الطالب", "الصف و الشعبة"])

for x in x:
    f.writerow([x["SN"],x["SC"]])

    # write2file(mreq2(i))
    # print(i)
#     input('press any thing please')

# dic = {'اسم الطالب كامل': [name1 , name2 ,name3 ,fname ],
# 'تاريخ الولادة' :[ bd , bm , by],
# 'مكان الولادة': pob,
# 'الجنسية': nationality,
# 'الديانه': religion,
# 'العمر في بداية العام الدراسي' :[ d , m , y]  ,
# 'اسم ولي الامر ' :name2 ,
# 'عمله': work,
# 'رقم الهاتف' :phone }

# data = {}

# data[0] = {'a' :1}
# data[1] = {'b': 2}
# print(data)

# print(readFile())
# name = soup(f'{readFile()}' , select='table > tbody > tr > td:nth-child(3)')
# Class = soup(f'{readFile()}' , select='table > tbody > tr > td:nth-child(5)')
# print(name[0] , Class[0] , sep="\n")
# for i in range(len(name)):
#     print( html2text(name[i]) , html2text(Class[i] ))
# write2file2(str(item[0]))

# data = {}
# data[0] = [1 ,2 ]
# data[1] = [3 ,4 ]
# print(data)

''' 
loop for the pages 
inside the loop i will collect the data from each page
store the data in a dictionary 
convert dictionary to json 
write json to file to convert it later to csv file 
'''












