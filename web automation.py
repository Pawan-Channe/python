
import os
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


file_name = r"C:\Users\pawan\Desktop\Universe\Userinformation.xlsx"

if os.path.exists(path=file_name):
    workbook = openpyxl.load_workbook(filename=file_name)
    worksheet = workbook['Data']
else:
    print('Please Create an Excel')

driver = webdriver.Chrome() # Add chromedriver.exe path here
driver.get(url='https://www.google.com/search?q=site:linkedin.com/in')

for _ in range(10):
    driver.find_element(By.TAG_NAME,'body').send_keys(Keys.END)
    time.sleep(3)

valid = []
all_href = driver.find_elements(By.XPATH,'//a[@jsname="UWckNb"]')
for i in all_href:
    href = i.get_attribute('href')
    valid.append(href)

for x in valid:
    driver.get(x)
    profile_url = x
    name = driver.find_element(By.XPATH,'//button[contains(@data-modal,"public_profile_top")]/h1').text
    username = x.split('/')[-1]
    try:
        currentdesignation = driver.find_element(By.XPATH,"//div[contains(@class, 'text-body-medium')]").text
    except:
        currentdesignation = 'NA'

    currentlyworkingas = 'Same as Current Designation'

    previousexperience = ''
    listofprexp = driver.find_elements(By.XPATH,'//span[text()="Experience"]/following::div[1]')
    for pre in listofprexp:
        previousexperience += pre.text

    educationalexperience = ''
    listofeduexp = driver.find_elements(By.XPATH,'//span[text()="Education"]/following::div[1]')
    for edu in listofeduexp:
        educationalexperience += edu.text

    worksheet.append([profile_url, name, username, currentdesignation, currentlyworkingas, previousexperience, educationalexperience])

workbook.save(filename=file_name)
workbook.close()
driver.close()
