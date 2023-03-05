import requests
def get_institution_info(Authorization):
    '''
    دالة تستخدم للحصول على معلومات المدرسة 
    و تستعمل header Authorization
    و تعود بقيمة الرمز الخاص بالمدرسة و الاسم و الرقم الوطني للمنشأة
    '''
    url = "https://emis.moe.gov.jo/openemis-core/restful/v2/Institution-Staff"
    querystring = {"_limit":"1","_contain":"Institutions","_fields":"Institutions.code,Institutions.id,Institutions.name"}
    headers = {"Authorization": Authorization}
    response = requests.request("GET", url, headers=headers, params=querystring)
    print(response.text)
