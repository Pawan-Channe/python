
import openpyxl, time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By


driver = webdriver.Chrome()
url = 'https://pshpgeorgia.entrykeyid.com/as/authorization.oauth2?response_type=code&client_id=cnc-provider-mono&scope=openid%20profile&state=vB-9k6EmWmXmYoNL8utnyblilceaKf_QWlFZWfXoK1A%3D&redirect_uri=https://provider.pshpgeorgia.com/careconnect/login/oauth2/code/pingcloud&code_challenge_method=S256&nonce=AgsvKdzpKd7Q1yEm5l36yZJxIVDSWCnvAglTuh35X-Q&code_challenge=ReuqnyhQi7m4vDFFhf8dFDml9Joa8aoLsGL-5tSjsgw&app_origin=https://provider.pshpgeorgia.com/careconnect/login/oauth2/code/pingcloud&brand=pshpgeorgia'
driver.maximize_window()
driver.get(url)

username = 'naggarwal@nath-mds.com'
password = 'Secure@2023'

find_user = driver.find_element(By.ID,'identifierInput')

find_user.send_keys(username)
post_button = driver.find_element(By.XPATH,'//*[@id="postButton"]/a')
post_button.click()

find_pass = driver.find_element(By.ID,'password')
find_pass.send_keys(password)

sign_button = driver.find_element(By.ID,'signOnButtonSpan')
sign_button.click()

time.sleep(5)
plan_Type = driver.find_element(By.ID,'providerProfileName')
plan_Type.click()

select_Ambetter = Select(plan_Type)
select_Ambetter.select_by_value('3257403')

Go_button = driver.find_element(By.ID,'medicalDropdownSubmitID')
Go_button.click()

Eligiblity = driver.find_element(By.XPATH,'//*[@id="home-page"]/body/header/nav[1]/div/ul/li[1]/a')
Eligiblity.click()


load = openpyxl.load_workbook('H:\AR PRODUCTION REPORTS\Business Intelligence\ALL BOTS-NATH\AR\Client\OTG_862\OTG Authrization Checking New1\AuthCheck Input.xlsx')
open_sheet = load['Sheet1']

bot_output = []

for num in range(2,30):

    colrow_A = chr(65)+str(num)
    typeOfplan = open_sheet[colrow_A].value
    str_plan = str(typeOfplan)

    if str_plan == 'AMBPEAC ( AMBETTER PEACHSTATE HELTH PLAN MO)':

        colrow_J = chr(74)+str(num)
        collecDate = open_sheet[colrow_J].value
        date_str = str(collecDate)
        date_format = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        orignal_date = date_format.strftime("%m/%d/%Y")

        time.sleep(3)
        colrow_B = chr(66)+str(num)
        policyID = open_sheet[colrow_B].value
        policy_str = str(policyID)

        time.sleep(3)
        colrow_C = chr(67)+str(num)
        DoB = open_sheet[colrow_C].value
        dob_str = str(DoB)
        dob_format = datetime.strptime(dob_str, "%Y-%m-%d %H:%M:%S")
        orignal_dob = dob_format.strftime("%m/%d/%Y")



        time.sleep(3)
        if date_str == 'None' or policy_str == 'None' or dob_str == 'None':
            continue

        else:

            dateOfservice = driver.find_element(By.NAME,'dos')
            dateOfservice.clear()
            time.sleep(2)
            dateOfservice.send_keys(orignal_date)

            memberID = driver.find_element(By.NAME,'memberIdOrLastName')
            memberID.clear()
            time.sleep(2)
            memberID.send_keys(policy_str)

            DOB_input = driver.find_element(By.NAME,'dob')
            DOB_input.clear()
            time.sleep(2)
            DOB_input.send_keys(orignal_dob)


            checkEligiblity = driver.find_element(By.NAME,'check')
            checkEligiblity.click()

            time.sleep(5)

            Not_Found = driver.find_element(By.XPATH,'//*[@id="returned-elgibility"]/table/tbody/tr/td[1]/span').text

            if Not_Found == 'Not Found':
                continue

            else:

                time.sleep(3)
                viewDetails = driver.find_element(By.XPATH,'//*[@id="returned-elgibility"]/table/tbody/tr/td[3]/a')
                viewDetails.click()

                time.sleep(5)
                Authorizations = driver.find_element(By.XPATH,'//*[@id="memberdetails-page"]/body/section/div/div[2]/div/nav/ul/li/a[contains(text(),"Authorizations")]')
                Authorizations.click()
                time.sleep(3)

                tr_element = driver.find_element(By.XPATH,'//*[@id="auths-page"]/body/section/div/div[2]/div/div/div[1]/table/tbody/tr')
                time.sleep(3)

                td_elements = tr_element.find_elements(By.XPATH,"./*") 

                time.sleep(5)

                for td in td_elements:
                    td_text = td.text
                    bot_output.append(td_text)


                time.sleep(3)
                colrow_F = chr(70)+str(num)
                statusCell = open_sheet[colrow_F].value = bot_output[0]
                colrow_G = chr(71)+str(num)
                NBRcell = open_sheet[colrow_G].value = bot_output[1]

                time.sleep(5)
                bot_output.clear()

                driver.back()
                driver.back()

                continue
    else:
        continue

load.save('AuthCheck Input.xlsx')

driver.quit()


