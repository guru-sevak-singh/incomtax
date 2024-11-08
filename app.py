import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
from selenium.webdriver.common.keys import Keys
import os
import pandas as pd
import pyautogui
import random
import json
import subprocess
from read_pdf import read_pdf_15CA, read_pdf_15CB
from datetime import datetime
import sys
import xlwings as xw
import requests

global app_starting_time
app_starting_time = datetime.now()

def is_software_not_worked():
    current_time = datetime.now()
    time_gap = current_time - app_starting_time
    time_gap = time_gap.seconds
    if time_gap > 300:
        os.startfile('app.exe')
        driver.quit()
        sys.exit()
    else:
        return True
    
url = 'https://www.dopagentsoftware.com/uploads/dopfiles/tympass.json'

response = requests.get(url)

response = response.json()

file_status = response['Income Tax']

if not file_status:
    message = "This is Testing File For More Information Contact To Us\n Thankyou"
    command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""{message}"", 0, ""Information"":close")'
    subprocess.Popen(command)
    sys.exit()


try:
    os.remove('acknowledgement_number.json')
except:
    pass

wb = xw.Book('acknowledgement_numbers.xlsm')
sheet = wb.sheets.active

global download_location
download_location = sheet['F6'].value

if download_location == None:
    download_location = os.path.expanduser("~")+"\\Downloads"
    sheet['F6'].value = download_location



global after_download_location
after_download_location = sheet['F7'].value

if after_download_location == None:
    after_download_location = download_location
    sheet['F7'].value = after_download_location

if os.path.exists(after_download_location):
    print('folder is already there...')

else:
    os.mkdir(after_download_location)\

global form_type
form_type = sheet['F8'].value

if form_type == None:
    form_type = "15CA"

wb.save()

def ReadExcel():
    try:
        detail = pd.read_excel('acknowledgement_numbers.xlsm', dtype=str)
        all_ack_numbers = detail.values.tolist()
        return all_ack_numbers
    except:
            message = "There is Sompe Problem in acknowledgement_numbers.xlsm Please Once Check it And Restart The Application"
            command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""{message}"", 0, ""Information"":close")'
            subprocess.Popen(command)
            sys.exit()
            
try:
    with open('acknowledgement_number.json') as file:
        acknowledge_data = json.load(file)
    
except:
    acknowledge_data = {}

all_ack_numbers = ReadExcel()

for data in all_ack_numbers:
    if data[1] in acknowledge_data.keys():
        try:
            if data[2] == 'Done':
                all_ack_numbers[data] = True
            pass
        except:
            pass
    else:
        key = data[1]
        try:
            if data[2] == 'Done':
                value = True
            else:
                value = False
        except IndexError:
            value = False
        acknowledge_data[key] = value

json_object = json.dumps(acknowledge_data, indent=4)
with open('acknowledgement_number.json', 'w') as file:
    file.write(json_object)


def get_latest_file_in_folder(dir_path, file_type):
  
  files = [f for f in os.listdir(dir_path) if f.endswith(file_type)]

  file_info = [(f, os.stat(os.path.join(dir_path, f)).st_mtime) for f in files]

  file_info.sort(key=lambda x: x[1])

  latest_file = file_info[-1][0]
  
  return latest_file

def CheckForNewFile(old_pdf_file):

    for n in range(10):
        time.sleep(1)
        new_pdf_file = get_latest_file_in_folder(download_location, ".pdf")
        
        if new_pdf_file == old_pdf_file:
            continue
        else:
            break

    return new_pdf_file


def OpenBrowser():
    global driver
    driver = uc.Chrome()
    driver.maximize_window()
    driver.get("https://eportal.incometax.gov.in/iec/foservices/#/login")

def LoginUser(user_id, password_value): #panAdhaarUserId, 'decimal@2005'
    x = random.randint(0, 50)
    y = random.randint(300, 600)
    
    pyautogui.moveTo(x, y)
    time.sleep(0.5)
    pyautogui.click(x, y)

    while True:
        is_software_not_worked()
        current_url = driver.current_url
        
        if current_url == "https://eportal.incometax.gov.in/iec/foservices/#/login/password":
            break
        else:
            try:
                time.sleep(3)
                login_id = driver.find_element(By.ID, 'panAdhaarUserId')

                login_id.click()
                login_id.clear()
                login_id.send_keys(user_id)

                submit_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="maincontentid"]/app-login/div/app-login-page/div/div[2]/div[1]/div[2]/button'))
                )

                submit_button.click()
            except:
                pass

    check_box = WebDriverWait(driver, 3).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="passwordCheckBox"]/label/div')))

    check_box.click()

    x = random.randint(0, 50)
    y = random.randint(300, 600)
    
    time.sleep(0.5)
    pyautogui.moveTo(x, y)
    time.sleep(0.5)
    pyautogui.click(x, y)

    # pyautogui.scroll(-500)

    password = WebDriverWait(driver, 5).until(
    EC.presence_of_element_located((By.ID, 'loginPasswordField'))
    )
    
    password.click()

    p_value = password.get_attribute('value')
    if p_value != "":
        password.clear()

    time.sleep(2)
    password.send_keys(password_value)

    driver.execute_script('document.getElementsByClassName("large-button-primary width marTop26")[0].click()')

    while True:
        is_software_not_worked()
        current_url = driver.current_url

        if current_url == 'https://eportal.incometax.gov.in/iec/foservices/#/dashboard':
            break
        else:
            time.sleep(0.25)
            try:
                driver.execute_script("document.getElementsByClassName('defaultButton primaryButton primaryBtnMargin')[0].click()")
            except:
                try:
                    driver.execute_script('document.getElementsByClassName("large-button-primary width marTop26")[0].click()')
                except:
                    pass

            continue

def OpenTheMainPage():
    
    while True:
        is_software_not_worked()
        time.sleep(2)
        try:
            loader = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div/mat-dialog-container/app-loader/div/div[2]'))
            )
            if 'oading' in loader.text:
                time.sleep(2)
                continue
            else:
                break
        except:
            break

    while True:
        is_software_not_worked()
        current_url = driver.current_url
        if current_url == "https://eportal.incometax.gov.in/iec/foservices/#/dashboard/viewFiledForms":
            break
        else:
            try:
                div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="mat-expansion-panel-header-8"]')))

                div.click()

                link = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="mat-expansion-panel-header-8"]/span[1]/mat-panel-description/a')))

                link.click()

            except Exception as e:
                error = type(e).__name__
                print(error)

                time.sleep(0.25)
    

    while True:
        is_software_not_worked()
        time.sleep(1)
        try:
            loader  = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[2]/div/mat-dialog-container/app-loader/div/div[2]'))
            )
            if 'oading' in loader.text:
                time.sleep(2)
                continue
            else:
                break
        except:
            break
    
    print('going to select the form')
    for n in range(7):
        try:
            form_name = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/div/mat-card/div[2]/div[{n}]/mat-card/mat-card-header/div/mat-card-title/div[2]'))
            )
            form_name = form_name.text
            
            if form_type in form_name:
                link = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/div/mat-card/div[2]/div[{n}]/mat-card/mat-card-content/div[3]/div[2]/span'))
                )
                
                link.click()
                break
                


        except:
            time.sleep(1)
            if n == 4:
                print('Problem to Open First Page')
            continue

def AddFilter(ack_number):
    is_software_not_worked()
    try:
        print('Add Filter')

        x = random.randint(0, 50)
        y = random.randint(300, 600)
        
        pyautogui.moveTo(x, y)
        time.sleep(0.25)
        pyautogui.click(x, y)

        pyautogui.scroll(2000)

        if "15CA" in form_type:
            filter_box = driver.find_element(By.XPATH, '//*[@id="maincontentid"]/app-dashboard/app-view-filed-forms/app-filed-form-ass-year-det/div[1]/div[2]/div[3]/button')
        
        else:
            filter_box = driver.find_element(By.XPATH, '//*[@id="maincontentid"]/app-dashboard/app-view-filed-forms/app-filed-form-token-det/div[1]/div[4]/div[3]/button')


        filter_box.click()

        ackNo = driver.find_element(By.ID, 'ackNo')

        ackNo.clear()
        ackNo.send_keys(ack_number)

        if "15CA" in form_type:
            filter_button = driver.find_element(By.XPATH, '//*[@id="maincontentid"]/app-dashboard/app-view-filed-forms/app-filed-form-ass-year-det/div[1]/div[3]/div/div/mat-card/mat-card-footer/div/button[2]')
        else:
            filter_button = driver.find_element(By.XPATH, '//*[@id="maincontentid"]/app-dashboard/app-view-filed-forms/app-filed-form-token-det/div[1]/div[3]/div/div/mat-card/mat-card-footer/div/button[2]')


        while True:
            try:
                filter_button.click()
                time.sleep(1)
                loader_text = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[2]/div/mat-dialog-container/app-loader/div/div[2]')))

                if 'oading' in loader_text.text:
                    break

            except:
                continue


        while True:
            try:
                loader_text = driver.find_element(By.XPATH,'/html/body/div[2]/div[2]/div/mat-dialog-container/app-loader/div/div[2]')

                if 'oading' in loader_text.text:
                    time.sleep(1)
                    continue
                
            except:
                break


        if "15CA" in form_type:
            card_title = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/app-filed-form-ass-year-det/div[1]/div[6]/mat-card/mat-card-header/div[2]/div/mat-card-title/span[2]')))
        
            card_title = card_title.text
            card_title = card_title.replace(" ", "")
            
        else:
            card_title = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/app-filed-form-token-det/div[1]/div[6]/mat-card/mat-card-content/mat-card/mat-card-header/div[2]/div/mat-card-title'))
            )
        
            card_title = card_title.text
            card_title = card_title.replace(" ", "")
            card_title = card_title.split(":")[1]

        if card_title == ack_number:
            return


        else:
            # click on reset button
            if '15CA' in form_type:
                reset_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/app-filed-form-ass-year-det/div[1]/div[3]/div/div/mat-card/mat-card-footer/div/button[1]')))
            else:
                reset_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/app-filed-form-token-det/div[1]/div[3]/div/div/mat-card/mat-card-footer/div/button[1]'))
                )

            reset_button.click()
            time.sleep(0.5)

            AddFilter(ack_number)

    except Exception as e:

        try:
            session_box = driver.find_element('id', 'sessionTimerModal')
            if session_box:
                button = driver.find_element('id', 'okBtnNav2')
                button.click()

        except:
            pass

        error_type = type(e).__name__
        if error_type == 'NoSuchWindowException':
            message = "Web Browser is Closed !"
            command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""{message}"", 0, ""Information"":close")'
            subprocess.Popen(command)
            sys.exit()
            

        time.sleep(1)

        print('Problem in Adding The Filter')

        AddFilter(ack_number)

def DownloadFile(ack_number):
    while True:
        is_software_not_worked()
        try:
            print('going to Download File ')
            x = random.randint(0, 50)
            y = random.randint(300, 600)
            
            pyautogui.moveTo(x, y)
            time.sleep(0.5)
            pyautogui.click(x, y)

            pyautogui.scroll(1500)

            new_name = os.path.join(after_download_location, f"{ack_number}.pdf")
            
            if os.path.exists(new_name):
                os.remove(new_name)

            old_pdf_file = get_latest_file_in_folder(download_location, ".pdf")

            if '15CA' in form_type:
                download_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/app-filed-form-ass-year-det/div[1]/div[6]/mat-card/mat-card-content/div/div[4]/div/div[2]/button')
                ))
            else:
                download_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div[1]/div[4]/app-dashboard/app-view-filed-forms/app-filed-form-token-det/div[1]/div[6]/mat-card/mat-card-content/mat-card/mat-card-content/div/div[4]/div/div[1]/button'))
                )
            
            download_button.click()

            new_file_name = CheckForNewFile(old_pdf_file)
            
            for n in range(3):
                if new_file_name == old_pdf_file:
                    download_button.click()
                else:
                    break
            
            new_file_name = os.path.join(download_location, new_file_name)

            if '15CA' in form_type:
                acknowledge_number = read_pdf_15CA(new_file_name)
            else:
                acknowledge_number = read_pdf_15CB(new_file_name)

            if acknowledge_number == ack_number:
                os.rename(new_file_name, new_name)
                break
        

        except Exception as e:
            try:
                session_box = driver.find_element('id', 'sessionTimerModal')
                if session_box:
                    button = driver.find_element('id', 'okBtnNav2')
                    button.click()

            except:
                pass

            x = random.randint(0, 50)
            y = random.randint(300, 600)
            
            pyautogui.moveTo(x, y)
            time.sleep(0.5)
            pyautogui.click(x, y)

            print('Error in Downloading File')

            error_type = type(e).__name__
            if error_type == 'NoSuchWindowException':
                message = "Web Browser is Closed !"
                command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""{message}"", 0, ""Information"":close")'
                subprocess.Popen(command)
                sys.exit()

            continue
        

OpenBrowser()
with open('user_id.txt') as file:
    user_id = file.read()

with open('password.txt') as file:
    password_value = file.read()

LoginUser(user_id=user_id, password_value=password_value)


OpenTheMainPage()


n = 1
print('Software Start Working..')
for ack_number in acknowledge_data:
    n += 1
    
    if acknowledge_data[ack_number] == False:

        AddFilter(ack_number)
        DownloadFile(ack_number)
        acknowledge_data[ack_number] = True

        new_json = json.dumps(acknowledge_data, indent=4)
        
        with open('acknowledgement_number.json', 'w') as file:
            file.write(new_json)

        sheet.range(f"C{n}").value = 'Done'
        wb.save()
        app_starting_time = datetime.now()
    

message = "Your Work Is Completed Successfully! Without Any Problem"
command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""{message}"", 0, ""Information"":close")'
subprocess.Popen(command)


