import os
import glob
import time
import pandas as pd
from pathlib import Path
from PIL import Image
import pytesseract

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import gspread,json
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials

# Set up Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

# Get Downloads directory
downloads_dir = Path.home() / "Downloads"

# Chrome options for headless mode
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)

# Accounts
accounts = [
    {"name": "SID", "id": os.environ.get('SID_ID'), "password": os.environ.get('SID_PASSWORD')},
    {"name": "RAJAN", "id": os.environ.get('RAJAN_ID'), "password": os.environ.get('RAJAN_PASSWORD')},
    {"name": "RESHMA", "id": os.environ.get('RESHMA_ID'), "password": os.environ.get('RESHMA_PASSWORD')},
]

# Connect to Google Sheets
def connect_to_gsheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

    # Load credentials from environment variable
    service_account_info = json.loads(os.environ["GSHEET_CREDENTIALS_JSON"])
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    
    client = gspread.authorize(creds)
    return client

# Get latest XLS files
def get_latest_xls_files(num_files=3):
    xls_files = sorted(glob.glob(str(downloads_dir / '*.xls')), key=os.path.getmtime, reverse=True)
    return xls_files[:num_files]

# Login Function
def login(user_id, pwd):
    driver.find_element(By.NAME, 'partnerId1').send_keys(user_id)
    driver.find_element(By.NAME, 'password1').send_keys(pwd)
    captcha_image = driver.find_element(By.ID, 'imgCaptcha')
    captcha_image.screenshot('captcha.png')
    captcha_text = pytesseract.image_to_string(Image.open('captcha.png')).strip().replace(" ", "")
    driver.find_element(By.NAME, 'capcode').send_keys(captcha_text)
    driver.find_element(By.NAME, 'action').click()

# Authorize and download all files
def authorize_all():
    driver.get(os.environ.get('PARTNER_DESK'))
    for acc in accounts:
        login(acc['id'], acc['password'])
        time.sleep(5)
        if 'E-MF Account' not in driver.page_source:
            login(acc['id'], acc['password'])
            time.sleep(10)
        if 'popupCloseButton' in driver.page_source:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'popupCloseButton'))).click()
        # if 'popupClose' in driver.page_source:
        #     WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'popupClose'))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@onclick='javascript:getAccountDetail();']"))).click()
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != driver.current_window_handle:
                driver.switch_to.window(window_handle)
                break
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'export_xls'))).click()
        time.sleep(5)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        driver.get(os.environ.get('PARTNER_DESK'))
    driver.quit()

# Combine downloaded files and upload to GSheet
def combine_xls_files_to_minimal_output():
    files = get_latest_xls_files(3)
    combined_data = pd.DataFrame()

    for file in files:
        df = pd.read_excel(file, engine="xlrd", header=1)
        df = df[:-2]  # remove last two summary rows
        df = df.drop(index=0)  # remove unwanted row after header
        df = df[["Investor", "Date of Birth"]]
        combined_data = pd.concat([combined_data, df], ignore_index=True)

    combined_data.drop_duplicates(inplace=True)
    combined_data.reset_index(drop=True, inplace=True)

    # Upload to Google Sheets
    client = connect_to_gsheet()
    sheet = client.open("combined").worksheet("combined")
    sheet.clear()
    set_with_dataframe(sheet, combined_data)
    print("✅ Data successfully updated in Google Sheets.")

if __name__ == "__main__":
    authorize_all()
    combine_xls_files_to_minimal_output()
