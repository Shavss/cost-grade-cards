from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

import msoffcrypto
import io
import pandas as pd

import pathlib
import time

#import credentials
import password




PATH = "C:\Program Files (x86)\chromedriver.exe"
script_directory = pathlib.Path().absolute()

chrome_options = Options()
chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")
driver = webdriver.Chrome(PATH, options=chrome_options)
chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")
driver.maximize_window()
driver.get("https://davidchipperfield.rapport3.com/backoffice/settingsconsole.asp")
driver.implicitly_wait(4)

#We used cookies in order to log us in into the system so the below part is not necessary

#username = credentials.TEST_IO_USERNAME
#password = credentials.TEST_IO_PASSWORD
#username_input = driver.find_element(By.ID, 'i0116')
#username_input.send_keys(username)
#username_input.send_keys(Keys.RETURN)
#time.sleep(5)
#password_input = driver.find_element(By.ID, 'i0118')
#password_input.send_keys(password)
#password_input.send_keys(Keys.RETURN)
#time.sleep(10)
#next_button = driver.find_element(By.XPATH, "//input[@type='submit']")
#next_button.click()

decrypted = io.BytesIO()

with open('221021 Cost and Grade Rates.xlsx', 'rb') as f:
    file = msoffcrypto.OfficeFile(f)
    file.load_key(password=password.file_password)
    file.decrypt(decrypted)

df = pd.read_excel(decrypted, "Grade rates for Rapport Entry")
del (decrypted)

print(df)