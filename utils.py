from selenium import webdriver
from selenium.webdriver.common.by import By
import geckodriver_autoinstaller

def get_cookie(User , Pass):
    geckodriver_autoinstaller.install()
    
    driver = webdriver.Firefox()
    driver.get("https://emis.moe.gov.jo/openemis-core/")
    assert "Login" in driver.page_source
    
    username = driver.find_element(By.ID, 'username')
    password = driver.find_element(By.ID, "password")

    username.send_keys(str(User))
    password.send_keys(str(Pass))
    driver.find_element(By.NAME, "submit").click()
    
    if 'invalid username or password' in driver.page_source:
        print('Invalid Username or Password')
        driver.close()
    else: 
        print('Login Success')
        assert "Dashboard" in driver.page_source
        all_cookies= driver.get_cookies()
        cookies_dict = {}
        for cookie in all_cookies:
            cookies_dict[cookie['name']] = cookie['value']
        driver.close()
        # print(cookies_dict)
        return cookies_dict   

def get_teacherClassesInfo():
    pass

def main():
    # get_cookie( 9971055725 , 9971055725)
    pass

if __name__ == "__main__":
    main()