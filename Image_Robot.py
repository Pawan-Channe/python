
import time
import re, os
import PyPDF2
import smtplib
import openpyxl
import pyautogui
import clipboard
import subprocess
import pandas as pd
from datetime import datetime as dt
from pywinauto.application import Application
from Rivercity_insured_bot import Rivercity_Insurance_ID_Fetch

Rivercity_Insurance_ID_Fetch()

url = 'https://eprg.wellmed.net/'

process = subprocess.Popen(["start", url], shell=True).wait(timeout=30)

while True:
    Login = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Login.PNG')
    if Login is not None:
        pyautogui.click(Login)
        break

time.sleep(10)
pyautogui.press('tab')
time.sleep(1)
pyautogui.typewrite('#Username')
time.sleep(1)
Continue = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Continue.PNG')
pyautogui.click(Continue)
time.sleep(5)
pyautogui.press('tab', presses=3)
time.sleep(1.5)
pyautogui.typewrite('#Password')
time.sleep(1.5)
pyautogui.press('enter')
time.sleep(5)
print('Successfully Login to Wellmed')

Agree = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Agree.PNG')
pyautogui.click(Agree)
time.sleep(2)
print('Successfully Clicked on Agree Button')

eligibility_pattern = r"Member ID:\s*((?:.|\n)*?)Name:"
insurance_pattern = r"Insurance:\s*(.*)"
plan_pattern = r"Plan:\s*((?:.|\n)*?)Plan Code:"
plan_code_pattern = r"Plan Code:\s*(.*)"
pcp_pattern = r"PCP:\s*(.*)"
clinic_pattern = r"Clinic:\s*((?:.|\n)*?)Start Date:"
start_date_pattern = r"Start Date:\s*(\d{2}/\d{2}/\d{4})"
benefits_pattern = r"Benefits:\s*(.*)"
date_verified_pattern = r"Date V erified:\s*(.*)"
spec_benefit_pattern = r"(?<=SPEC\s)(?:\$\$|\$)?(?:\d+|%)+" 

Rivercity_Excel_File_Path = r"H:\AR PRODUCTION REPORTS\Business Intelligence\Eligibility and Benefits BOT Output.xlsx"
WorkBook = openpyxl.load_workbook(Rivercity_Excel_File_Path)
WorkSheet = WorkBook['Data']
print('Rivercity Bot Output file Loaded Successfully')

folder_path = r"H:\AR PRODUCTION REPORTS\Business Intelligence\Eligibility BOT\PDF Downloads"
current_datetime = dt.now()
today_date = current_datetime.strftime("%m-%d-%Y")
today_folder_path = os.path.join(folder_path, today_date)
if not os.path.exists(today_folder_path):
    os.mkdir(today_folder_path)
    print(f"Created folder: {today_date}")

for row in WorkSheet.iter_rows(min_row=2):

    if row[0].value != None and row[11].value == None:
    
        Patient_Name = row[1].value
        split_strings = [substring.strip() for substring in Patient_Name.split(',')]

        Patient_PDF = os.path.join(today_folder_path, Patient_Name + '.pdf')

        Date_of_Birth = row[2].value
        Format_DOB = Date_of_Birth.strftime("%m/%d/%Y")

        Insured_ID = str(row[10].value)
        if not Insured_ID.endswith('01'):
            Added_Suffix = Insured_ID + '01'
        else:
            Added_Suffix = Insured_ID

        Eligibility = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Eligibility.PNG')
        pyautogui.moveTo(Eligibility)
        time.sleep(2)
        print('Successfully Clicked on Eligibility Button')

        MemberSearch = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Member Search.PNG')
        pyautogui.click(MemberSearch)
        time.sleep(3)
        print('Successfully Clicked on Member Search Button')
        
        if Insured_ID == 'Not Available':
            First_Name = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\First Name.PNG')
            pyautogui.click(First_Name)     
            print('Successfully Clicked on First Name Input Button')
            pyautogui.typewrite(split_strings[1])
            time.sleep(1.5)

            Last_Name = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Last Name.PNG')
            pyautogui.click(Last_Name)
            print('Successfully Clicked on Last Name Input Button')
            pyautogui.typewrite(split_strings[0])
            time.sleep(1.5)

            pyautogui.press('tab')
            pyautogui.typewrite(Format_DOB)
            time.sleep(1.5)
        else:
            Member_ID = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Member ID.PNG')
            pyautogui.click(Member_ID)
            print('Successfully Clicked on Member ID Input Button')
            pyautogui.typewrite(Added_Suffix)
            time.sleep(1.5)

        Search = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Search.PNG')
        pyautogui.click(Search)
        print('Successfully Clicked on Search Button')

        time.sleep(5.55555555555555555)

        Eligible = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Eligible.PNG')
        if Eligible is not None:
            pyautogui.moveTo(Eligible)
            pyautogui.moveRel(0, 30)
            pyautogui.click()
            time.sleep(3)

            Historical = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Historical.PNG')
            pyautogui.click(Historical)
            time.sleep(3)

            clipboard.copy('Not Available')
            Thru_End = pyautogui.locateOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Thru End.PNG')
            pyautogui.moveTo(Thru_End)
            pyautogui.moveRel(-35, 25)
            pyautogui.dragRel(70, 0)
            pyautogui.hotkey('ctrl', 'c')
            copied_date = clipboard.paste()
            time.sleep(1)

            Back = pyautogui.locateOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Back.PNG')
            pyautogui.click(Back)
            time.sleep(2)

            Print_Eligibility = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Print Eligibility.PNG')
            pyautogui.click(Print_Eligibility)
            time.sleep(3)
            print('Successfully Clicked on Print Eligibility Button')
            
            Print = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Print.PNG')
            pyautogui.click(Print)
            time.sleep(5)
            print('Successfully Clicked on Print Button')

            Save = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Save.PNG')
            pyautogui.click(Save)
            print('Successfully Clicked on Save Button')

            chrome_app = Application(backend='uia').connect(title_re="Print Member Eligibility - Google Chrome", timeout=60)
            window_switch = chrome_app.PrintMemberEligibilityGoogleChrome.child_window(title="Save As", control_type="Window")
            # chrome_app.PrintMemberEligibilityGoogleChrome.print_control_identifiers()

            pyautogui.typewrite(Patient_Name)
            time.sleep(1)

            Previous_Locations = chrome_app.PrintMemberEligibilityGoogleChrome.child_window(title="Previous Locations", control_type="Button")
            Previous_Locations.click_input()

            pyautogui.typewrite(today_folder_path)
            pyautogui.press('enter')
            time.sleep(1)

            Save_Button = chrome_app.PrintMemberEligibilityGoogleChrome.child_window(title="Save", auto_id="1", control_type="Button")
            Save_Button.click_input()
            time.sleep(1.5)
            
            try:
                Replace_Yes = chrome_app.PrintMemberEligibilityGoogleChrome.child_window(title="Yes", auto_id="CommandButton_6", control_type="Button")
                Replace_Yes.click_input()
                time.sleep(1.5)
            except:
                pass

            Close = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Close.PNG')
            pyautogui.click(Close)
            time.sleep(2)
            print('Successfully Clicked on Close Button')

            path = open(Patient_PDF, 'rb')
            pdfReader = PyPDF2.PdfReader(path)
            from_page = pdfReader.pages[0]
            text = from_page.extract_text()

            eligibility_match = re.search(eligibility_pattern, text)
            insurance_match = re.search(insurance_pattern, text)
            plan_match = re.search(plan_pattern, text)
            plan_code_match = re.search(plan_code_pattern, text)
            pcp_match = re.search(pcp_pattern, text)
            clinic_match = re.search(clinic_pattern, text)
            start_date_match = re.search(start_date_pattern, text)
            benefits_match = re.search(benefits_pattern, text)
            date_verified_match = re.search(date_verified_pattern, text)
            spec_benefit_match = re.search(spec_benefit_pattern, text)

            eligibility = eligibility_match.group(1) if eligibility_match else 'NA'
            insurance = insurance_match.group(1) if insurance_match else 'NA'
            plan = plan_match.group(1) if plan_match else 'NA'
            plan_code = plan_code_match.group(1) if plan_code_match else 'NA'
            pcp = pcp_match.group(1) if pcp_match else 'NA'
            clinic = clinic_match.group(1) if clinic_match else 'NA'
            start_date = start_date_match.group(1) if start_date_match else 'NA'
            benefits = benefits_match.group(1) if benefits_match else 'NA'
            date_verified = date_verified_match.group(1) if date_verified_match else 'NA'
            spec_benefit = spec_benefit_match.group(0) if spec_benefit_match else 'NA'

            Current_DateTime = dt.now()
            Current_DateTime = Current_DateTime.strftime("%m/%d/%Y %H:%M:%S")
            print('Successfully Got the Current Date and Time')

            if 'Not Eligible' in eligibility:
                row[11].value = 'No'
            else:
                row[11].value = 'Yes'
                row[12].value = insurance
                row[13].value = plan
                row[14].value = plan_code
                row[15].value = pcp
                row[16].value = clinic
                row[17].value = start_date
                row[18].value = benefits
                row[19].value = spec_benefit
                row[20].value = date_verified
                row[22].value = Current_DateTime
                row[23].value = copied_date
                print('Successfully Data Updated in Excel Sheet')

                if spec_benefit.endswith('%'):
                    spec_benefit = '$0'

                if copied_date == 'Not Available':
                    row[21].value = f"{date_verified}:-"\
                        f"Wellmed is active since - {start_date}"\
                        f" Facility INN, Copay {spec_benefit},"\
                        f" Referral not required"
                    print('Rivercity Notes Updated')
                else:
                    row[21].value = f"{date_verified}:-"\
                        f"Wellmed is active since - {start_date}"\
                        f" to {copied_date}"\
                        f" Facility INN, Copay {spec_benefit},"\
                        f" Referral not required"
                    print('Rivercity Notes Updated')

            WorkBook.save(Rivercity_Excel_File_Path)
        else:
            row[11].value = 'No'
            WorkBook.save(Rivercity_Excel_File_Path)

            pyautogui.hotkey('alt', 'left')
            time.sleep(3)
            
            Agree = pyautogui.locateCenterOnScreen(r'C:\Users\pchanne\Desktop\CodeVerse\Rivercity Gallery\Agree.PNG')
            pyautogui.click(Agree)
            time.sleep(2)

print('Successfully Clicked on Agree Button')
WorkBook.save(Rivercity_Excel_File_Path)
WorkBook.close()
pyautogui.hotkey('ctrl', 'w')
print('Rivercity Eligibility Bot Exicuted Successfully')

