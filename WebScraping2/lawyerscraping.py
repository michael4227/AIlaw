# from numpy import printoptions
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
# from pynput.keyboard import Key, Controller
from keyboard import press
# import pdfkit
import time
import xlsxwriter

path = '/usr/local/bin/chromedriver'
driver = webdriver.Chrome(path)
driver.get('https://apps.calbar.ca.gov/members/ls_search.aspx')
dropdown = driver.find_element_by_id('ctl00_PageContent_ddlSpecialty')
slcdropdown = Select(dropdown)
slcdropdown.select_by_value('04') #Family Law's html value
search_button = driver.find_element(by=By.XPATH, value='//*[@id="ctl00_PageContent_btnSubmit"]') # button's xpath copied from html
search_button.click()
time.sleep(3)



def get_lawyer_info():
    infolst = []
    lwyinfonm = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="moduleMemberDetail"]/div[2]/h3/b')))
    # print(lwyinfonm.text)
    lwyinfostatus = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="moduleMemberDetail"]/div[2]/div/p')))
    # print(lwyinfostatus.text)
    lwyinfoads = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="moduleMemberDetail"]/div[3]/p[1]')))
    lwyinfophone = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="moduleMemberDetail"]/div[3]/p[2]')))
    lwyinfoweb = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="moduleMemberDetail"]/div[3]/p[3]')))
    print(lwyinfoweb.text)
    infolst.append(lwyinfonm.text)
    infolst.append(lwyinfostatus.text)
    infolst.append(lwyinfoads.text)
    infolst.append(lwyinfophone.text)
    infolst.append(lwyinfoweb.text)
    # infolst.append(lwyinfoemail)
    # print(infolst)
    return infolst

filenum = 1000
def getpdf():
    global filenum
    try:
        driver.set_script_timeout(5)
        driver.execute_script('window.print()')
    except:
        time.sleep(2)   
        # keyboard = Controller()
        # keyboard.press(Key.enter)
        # time.sleep(1)
        # keyboard.press(Key.enter)
        press('enter')
        time.sleep(2)
        intfilenum = str(filenum)
        for i in intfilenum:
            press(f'{i}')
        time.sleep(1)
        press('enter')
        time.sleep(1)
    filenum += 1
    # print(intfilenum)
    return filenum
# print(filenum)

lwysinfo = []
tr = 0
for i in range(1035,1332):
    tr += 1
    lawyer = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,f'//*[@id="content-main"]/table/tbody/tr[{tr}]/td[1]/a')))
    lwynm = f'{str(lawyer.text)}'
    print(lwynm)
    lawyerclc = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,f'{lwynm}')))
    lawyerclc.click()
    lwysinfo.append(get_lawyer_info())
    time.sleep(1)
    # getpdf()
    driver.back()
    time.sleep(1)
print(lwysinfo)

# def doublereturn():
#     time.sleep(3)
#     keyboard = Controller()
#     keyboard.press(Key.enter)
#     time.sleep(1)
#     keyboard.press(Key.enter)



#print(driver.page_source)
# time.sleep(5)

# driver.close() #driver.quit()

def append():
    dcwb= xlsxwriter.Workbook('WS_FINAL3.xlsx')
    dcws = dcwb.add_worksheet('infosheet')
    row = 0
    col = 0
    for lwy in lwysinfo:
        for eachinfo in lwy:
            dcws.write(row,col,eachinfo)
            col += 1
        col = 0
        row += 1
    dcwb.close()
append()
