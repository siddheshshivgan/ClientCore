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

# Set up Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

# Get Downloads directory
downloads_dir = Path.home() / "Downloads"

chrome_options = Options()
# chrome_options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

accounts = [
    {"name": "SID", "id": os.environ.get('SID_ID'), "password": os.environ.get('SID_PASSWORD')},
    {"name": "RAJAN", "id": os.environ.get('RAJAN_ID'), "password": os.environ.get('RAJAN_PASSWORD')},
    {"name": "RESHMA", "id": os.environ.get('RESHMA_ID'), "password": os.environ.get('RESHMA_PASSWORD')},
]

def get_latest_xls_files(num_files=3):
    xls_files = sorted(glob.glob(str(downloads_dir / '*.xls')), key=os.path.getmtime, reverse=True)
    return xls_files[:num_files]

def login(user_id, pwd):
    driver.find_element(By.NAME, 'partnerId1').send_keys(user_id)
    driver.find_element(By.NAME, 'password1').send_keys(pwd)
    captcha_image = driver.find_element(By.ID, 'imgCaptcha')
    captcha_image.screenshot('captcha.png')
    captcha_text = pytesseract.image_to_string(Image.open('captcha.png')).strip().replace(" ", "")
    driver.find_element(By.NAME, 'capcode').send_keys(captcha_text)
    driver.find_element(By.NAME, 'action').click()

def authorize_all():
    driver.get(os.environ.get('PARTNER_DESK'))
    for acc in accounts:
        login(acc['id'], acc['password'])
        time.sleep(5)
        if 'E-MF Account' not in driver.page_source:
            login(acc['id'], acc['password'])
            time.sleep(10)
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


# Combine data from latest files
def combine_xls_files_to_minimal_output():
    files = get_latest_xls_files(3)
    combined_data = pd.DataFrame()

    for file in files:
        df = pd.read_excel(file, engine="xlrd", header=1)
        df = df[:-2]  # remove last two summary rows
        df = df.drop(index=0)  # remove unwanted row after header
        df = df[["Investor", "Date of Birth"]]
        combined_data = pd.concat([combined_data, df], ignore_index=True)

    # Remove duplicates, reset index
    combined_data.drop_duplicates(inplace=True)
    combined_data.reset_index(drop=True, inplace=True)

    # Save final output
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    combined_file = output_dir / "combined.xlsx"
    combined_data.to_excel(combined_file, index=False)
    return combined_file

if __name__ == "__main__":
    authorize_all()
    combine_xls_files_to_minimal_output()
