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
        dd.select_by_visible_text(self.input)

    def textfill(self):
        textfield = driver.find_element(By.XPATH, self.xpath)
        textfield.clear()
        textfield.send_keys(self.input)
        #textfield.send_keys(Keys.ENTER)


class field:

    def __init__(self, xpath):
        self.xpath = xpath
    def select(self):
        select = driver.find_element(By.XPATH, self.xpath)
        select.click()


WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it("iframe1"))

setting = field("//div[@title='Projects']")
setting.select()

setting1 = field("//div[@title='Test']")
setting1.select()

j = 0
x = 0

while x < len(df):
    driver.switch_to.frame("iframe1")

    # Select Branch - DCAL
    branch = item("DCAL", "DCAL" ,"//select[@id='BranchID']")
    branch.dropdown()

    driver.switch_to.frame("summaryframe")

    l = 0
    k = 1
    key = keys[k]

    # Filling out the Plan Cost Grade Rates per band
    while k<=11:
        field1 = field(f"//div[@id='DefRates_{l}_GradeCostRate']")
        field1.select()

        item1 = item(key, truncate(dic[key][j], 2) ,"//input[@id='Grid_-1_txtGeneric']")
        item1.textfill()

        k = k + 1
        key = keys[k]
        l = l + 1

    # Filling out the Plan Charge Grade Rates per band

    l = 0
    k = 1
    key = keys[k]

    while k<=11:
        field2 = field(f"//div[@id='DefRates_{l}_GradeChargeRate']")
        field2.select()

        multiplier = dic["Multiplier"][j]

        item2 = item(key, truncate(dic[key][j]*multiplier, 2), "//input[@id='Grid_-1_txtGeneric']")
        item2.textfill()

        k = k + 1
        key = keys[k]
        l = l + 1

    l = 0
    k = 1

    # Checking for NaN in Margin

    while k<=11:
        field3 = field(f"//div[@id='DefRates_{l}_GradeMargin']")
        field3.select()

        check = driver.find_element(By.XPATH, "//input[@id='Grid_-1_txtGeneric']")

        if check.get_attribute("value") == "NaN":
            item3 = item(key, 0, "//input[@id='Grid_-1_txtGeneric']")
            item3.textfill()
        else:
            pass

        k = k + 1
        l = l + 1

    driver.switch_to.default_content()
    driver.switch_to.frame("iframe1")
    driver.switch_to.frame("iframe1")

    save_button = field("//button[@id='btnSave']")
    save_button.select()

    driver.switch_to.default_content()
    driver.switch_to.frame("iframe1")

    setting2 = field("//div[@title='Grade Rate Bulk Update']")
    setting2.select()


    driver.switch_to.frame("iframe1")
    iframe3 = driver.find_element(By.XPATH, "//iframe[@src='/backoffice/settings/GradeBulkUpdate_projectselector.asp']")
    driver.switch_to.frame(iframe3)

    # Filtering out the projects#
    button3 = field("//i[@title='search']")
    button3.select()

    y = 0
    f1= 3
    f2 = 2

    while y < len(dic["Filters1"]): #FIX THIS

        window_before = driver.window_handles[0]
        window_after = driver.window_handles[1]
        driver.switch_to.window(window_after)

        item3 = item("Filters1", dic["Filters1"][y], f"//*[@id='filterselect_{f1}']")
        item3.dropdown()

        item4 = item("Filters2", dic["Filters2"][y], f"//*[@id='filtervalue_{f2}']")
        item4.dropdown()

        y = y + 1
        f1 = f1 + 1
        f2 = f2 + 1
    exit()

    button = field("//button[@class='ci-button ci-text-button ci-button-default']")
    button.select()
    driver.close()
    driver.switch_to.window(window_before)

    driver.switch_to.default_content()
    driver.switch_to.frame("iframe1")
    driver.switch_to.frame("iframe1")
    driver.switch_to.frame(iframe3)

    project = field("//option[text()='T9991B Test Fee Project 1']") #change to the variable with a project name
    project.select()

    button = field("//a[@href='javascript:movetoright()']")
    button.select()

    driver.switch_to.default_content()
    driver.switch_to.frame("iframe1")
    driver.switch_to.frame("iframe1")

    if j == 0:

        button1 = field("//input[@id='GetProjects']")
        button1.select()

        button2 = field("//input[@id='btnProcess']")
        button2.select()

    else:

        item5 = item("blank", "Create a new grade card for each project", "//select[@id='runMode']")
        item5.dropdown()

        dropdown1 = driver.find_element(By.XPATH, "// select[ @ id = 'startMonth']")
        dd1 = Select(dropdown1)
        dd1.select_by_value(dic["Band"][j])

        button1 = field("//input[@id='GetProjects']")
        button1.select()

        button2 = field("//input[@id='btnProcess']")
        button2.select()


    driver.switch_to.default_content()
    driver.switch_to.frame("iframe1")

    setting1.select()
    x = x + 1
    j = j + 1





