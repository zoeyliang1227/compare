import time
import yaml
import re

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

config = yaml.load(open('Testlink_config.yml'), Loader=yaml.Loader)

timeout = 20
Document_ID_list = []
Different_Locations = []

def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')                 # 瀏覽器不提供可視化頁面
    options.add_argument('-no-sandbox')  # 以最高權限運行
    options.add_argument('--start-maximized')  # 縮放縮放（全屏窗口）設置元素比較準確
    options.add_argument('--disable-gpu')  # 谷歌文檔說明需要加上這個屬性來規避bug
    options.add_argument('--window-size=1920,1080')  # 設置瀏覽器按鈕（窗口大小）
    options.add_argument('--incognito')  # 啟動無痕

    driver = webdriver.Chrome(options=options)
    url = config['url']

    # driver.implicitly_wait(10)
    # driver.get(url)
    # driver.delete_all_cookies() #清cookie

    # with open('cookies.yml', 'r', encoding='utf-8') as f:
    #     cookies = yaml.load(f, Loader=yaml.FullLoader)
    #     for c in cookies:
    #         driver.add_cookie(c)

    # print('Current cookies:', driver.get_cookies())
    driver.get(url)

    return driver


def find_from_testlink(GetFromPDF_list):
    # print(GetFromPDF_list)
    driver = get_driver()
    login(driver)

    driver.switch_to.frame('titlebar')
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/form/select'))).click()
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/form/select/option[9]'))).click()
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/a[2]'))).click()
    time.sleep(2)
    driver.switch_to.parent_frame()
    driver.switch_to.frame('mainframe')
    driver.switch_to.frame('treeframe')

    for a, Document_ID in enumerate(GetFromPDF_list):
        print(f'目前正在使用 {Document_ID}搜尋 Testlink, 還差 {len(GetFromPDF_list)-a} 個')
        time.sleep(2)
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.NAME, 'filter_doc_id'))).clear()
        if Document_ID and Document_ID != '':
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.NAME, 'filter_doc_id'))).send_keys(Document_ID)
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.NAME, 'filter_doc_id'))).send_keys(Keys.ENTER)

            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'expand_tree'))).click()
            if check_ID_text(driver) == True:
                ID_text = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ext-gen14"]/li/ul/li/ul/li/ul/li/ul/li/ul/li/div/a/span')))
                # print((Document_ID.strip()) ,'zzz', ID_text.text)
                if (Document_ID.strip()) in ID_text.text:
                    time.sleep(2)
                    ID_text.click()
                    driver.switch_to.parent_frame()
                    driver.switch_to.frame('workframe')
                    if check_text(driver) == True:
                        text = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/table[1]/tbody/tr[7]/td/fieldset'))).text

                        # print(text, 'zzz', Document_ID)
                        if 'Coverage' in text:
                            text = text.replace('Coverage', '')
                            # print(text.strip())
                            Document_ID_list.append(text.strip())
                    else:
                        Document_ID_list.append('')

                    driver.switch_to.parent_frame()
                    driver.switch_to.frame('treeframe')
                else:
                    Different_Locations.append('Document_ID')
                    print(f'{Document_ID}位置不同')
        else:
            Document_ID_list.append('')
    
    print(Different_Locations)        
    # print(len(Document_ID_list))  
    return Document_ID_list
 
def login(driver):
    print('Waiting for login...')
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.NAME, 'tl_login'))
    ).send_keys(config['username'])
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.NAME, 'tl_password'))
    ).send_keys(config['password'])
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.NAME, 'login_submit'))
    ).click()
    print('Waiting for the website to load...')
    # cookie2 = driver.get_cookies() #取得登入後cookie
    # with open('cookies.yml', 'w', encoding='utf-8') as f:
    #     yaml.dump(cookie2, stream=f, allow_unicode=True)


def check_prd(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="ext-gen14"]/li/ul/li')
        return True
    except:
        return False
    
def check_ID_text(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="ext-gen14"]/li/ul/li/ul/li/ul/li/ul/li/ul/li/div/a/span')
        return True
    except:
        return False
    
def check_click(driver, x):
    try:
        driver.find_element(By.ID, f'ext-gen{x}')
        return True
    except:
        return False

def check_text(driver):
    try:
        driver.find_element(
            By.XPATH, '/html/body/div/table[1]/tbody/tr[7]/td/fieldset/span/a'
        )
        return True
    except:
        return False
