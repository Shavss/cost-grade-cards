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
dic = df.to_dict()
keys = list(dic)

k = 1

key = keys[k]
del (decrypted)

def truncate(n, decimals=0):
    multiplier = 10 ** decimals
    return int(n * multiplier) / multiplier

class item:
    def __init__(self, key, input, xpath):
        self.key = key
        self.input = input
        self.xpath = xpath

    def dropdown(self):
        dropdown = driver.find_element(By.XPATH, self.xpath)
        dd = Select(dropdown)
        if self.input != " ":
            dd.select_by_visible_text(self.input)
        else:
            pass

    def textfill(self):
        textfield = driver.find_element(By.XPATH, self.xpath)
        if self.input != " ":
            textfield.clear()
            textfield.send_keys(self.input)
        pass

class field:

    def __init__(self, xpath):
        self.xpath = xpath
    def select(self):
        select = driver.find_element(By.XPATH, self.xpath)
        select.click()


WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it("iframe1"))

setting = field("//div[@title='Projects']")
setting.select()

setting2 = field("//div[@title='Test']")
setting2.select()

driver.switch_to.frame("iframe1")


#select Branch - DCAL
branch = item("DCAL", "DCAL" ,"//select[@id='BranchID']")
branch.dropdown()

driver.switch_to.frame("summaryframe")

j = 0
l = 0

# Filling out the Plan Cost Grade Rates per band
while k<11:
    field1 = field(f"//div[@id='DefRates_{l}_GradeCostRate']")
    field1.select()

    item1 = item(key, truncate(dic[key][j], 2) ,"//input[@id='Grid_-1_txtGeneric']")
    item1.textfill()

    k = k + 1
    key = keys[k]
    l = l + 1

# Filling out the Plan Charge Grade Rates per band TBC


