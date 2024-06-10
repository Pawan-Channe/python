
import os
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions


file_name = r"C:\Users\pawan\Desktop\Universe\Userinformation.xlsx"

if os.path.exists(path=file_name):
    workbook = openpyxl.load_workbook(filename=file_name)
    worksheet = workbook['Data']
else:
    print('Please Create an Excel')

user = '#### Your Username ####'
paswd = '#### Your Password ####'

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chrome_options)

driver.get(url='https://www.google.com/search?q=site:linkedin.com/in')

for _ in range(10):
    driver.find_element(By.TAG_NAME,'body').send_keys(Keys.END)
    time.sleep(3)

valid = []
all_href = driver.find_elements(By.XPATH,'//a[@jsname="UWckNb"]')
for i in all_href:
    href = i.get_attribute('href')
    valid.append(href)

count = 1
for x in valid:

    driver.get(x)
    time.sleep(5)

    heading = driver.find_element(By.TAG_NAME,'h1').text.strip()

    if heading == 'Join LinkedIn':
        driver.find_element(By.XPATH,'(//button[contains(text(),"Sign in")])[1]').click()
        time.sleep(2.5)
        driver.find_element(By.ID,"session_key").send_keys(user)
        driver.find_element(By.ID,"session_password").send_keys(paswd)
        driver.find_element(By.XPATH,'//button[contains(text(),"Sign in") and @type="submit"]').click()

    elif heading == 'Welcome to your professional community':
        driver.find_element(By.ID,"session_key").send_keys(user)
        driver.find_element(By.ID,"session_password").send_keys(paswd)
        driver.find_element(By.XPATH,'//button[contains(text(),"Sign in") and @type="submit"]').click()

    elif heading == 'Sign in':
        driver.find_element(By.ID,'username').send_keys(user)
        driver.find_element(By.ID,'password').send_keys(paswd)
        driver.find_element(By.XPATH,'//*[contains(text(),"Sign in") and @type="submit"]').click()
    else:
        try:
            driver.find_element(By.XPATH,'(//button[@aria-label="Ignorer"])[1]').click()
        except:
            pass

    if heading == 'Join LinkedIn' or heading == 'Welcome to your professional community' or heading == 'Sign in':
        WebDriverWait(driver, 500).until(expected_conditions.element_to_be_clickable((By.XPATH,'//*[@placeholder="Search"]')))
        if count == 1:
            driver.back()
    
    WebDriverWait(driver, 100).until(expected_conditions.visibility_of_element_located((By.ID,"profile-content")))

    profile_url = x
    name = driver.find_element(By.TAG_NAME,'h1').text
    username = x.split('/')[-1]

    try:
        currentdesignation = driver.find_element(By.XPATH,"//div[contains(@class, 'text-body-medium')]").text
    except:
        currentdesignation = 'NA'

    currentlyworkingas = 'Same as Current Designation'

    previousexperience = ''
    listofprexp = driver.find_elements(By.XPATH,'//span[text()="Experience"]/following::div[1]//div[contains(@class,"justify-space-between")]')
    listofprexp = list(set(listofprexp))
    for pre in listofprexp:
        previousexperience += pre.text

    educationalexperience = ''
    listofeduexp = driver.find_elements(By.XPATH,'//span[text()="Education"]/following::div[1]//div[contains(@class,"justify-space-between")]')
    listofeduexp = list(set(listofeduexp))
    for edu in listofeduexp:
        educationalexperience += edu.text

    worksheet.append([profile_url, name, username, currentdesignation, currentlyworkingas, previousexperience, educationalexperience])
    count += 1
    workbook.save(filename=file_name)
    
workbook.save(filename=file_name)
workbook.close()
driver.close()

print('All Done')
