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


class Cost_Grade_Bot:

    def __init__(self):
        decrypted = io.BytesIO()

        with open('221021 Cost and Grade Rates.xlsx', 'rb') as f:
            file = msoffcrypto.OfficeFile(f)
            file.load_key(password=password.file_password)
            file.decrypt(decrypted)
        global key, df, dic, keys, counter

        df = pd.read_excel(decrypted, "Grade rates for Rapport Entry")
        dic = df.to_dict()
        keys = list(dic)
        del (decrypted)

        # Counting the lengths of "Filters1" column
        counter = 0
        for k in range(0, len(dic["Filters1"])):
            if pd.isna(dic["Filters1"][k]) == False:
                counter = counter + 1
            else:
                pass

        k = 1

        key = keys[k]

    def truncate(self, n, decimals=0):
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

    def start_cards(self, project):

        print("initiating bot...")

        PATH = "C:\Program Files (x86)\chromedriver.exe"
        script_directory = pathlib.Path().absolute()

        chrome_options = Options()
        chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")
        global driver
        driver = webdriver.Chrome(PATH, options=chrome_options)
        chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")
        driver.maximize_window()
        driver.get("https://davidchipperfield.rapport3.com/backoffice/settingsconsole.asp")
        driver.implicitly_wait(4)

        # def __init__(self):

        # We used cookies in order to log us in into the system so the below part is not necessary

        # username = credentials.TEST_IO_USERNAME
        # password = credentials.TEST_IO_PASSWORD
        # username_input = driver.find_element(By.ID, 'i0116')
        # username_input.send_keys(username)
        # username_input.send_keys(Keys.RETURN)
        # time.sleep(5)
        # password_input = driver.find_element(By.ID, 'i0118')
        # password_input.send_keys(password)
        # password_input.send_keys(Keys.RETURN)
        # time.sleep(10)
        # next_button = driver.find_element(By.XPATH, "//input[@type='submit']")
        # next_button.click()

        outter = Cost_Grade_Bot()

        WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it("iframe1"))

        setting = outter.field("//div[@title='Projects']")
        setting.select()

        setting1 = outter.field("//div[@title='Test']")
        setting1.select()

        j = 0
        x = 0

        while x < len(df):
            driver.switch_to.frame("iframe1")

            # Select Branch - DCAL
            branch = outter.item("DCAL", "DCAL" ,"//select[@id='BranchID']")
            branch.dropdown()

            driver.switch_to.frame("summaryframe")

            l = 0
            k = 1
            key = keys[k]

            # Filling out the Plan Cost Grade Rates per band
            while k<=11:
                field1 = outter.field(f"//div[@id='DefRates_{l}_GradeCostRate']")
                field1.select()

                item1 = outter.item(key, outter.truncate(dic[key][j], 2), "//input[@id='Grid_-1_txtGeneric']")
                item1.textfill()

                k = k + 1
                key = keys[k]
                l = l + 1

            # Filling out the Plan Charge Grade Rates per band

            l = 0
            k = 1
            key = keys[k]

            while k<=11:
                field2 = outter.field(f"//div[@id='DefRates_{l}_GradeChargeRate']")
                field2.select()

                multiplier = dic["Multiplier"][j]

                item2 = outter.item(key, dic[key][j]*multiplier, "//input[@id='Grid_-1_txtGeneric']")
                item2.textfill()

                k = k + 1
                key = keys[k]
                l = l + 1

            l = 0
            k = 1

            # Checking for NaN in Margin

            while k<=11:
                field3 = outter.field(f"//div[@id='DefRates_{l}_GradeMargin']")
                field3.select()

                check = driver.find_element(By.XPATH, "//input[@id='Grid_-1_txtGeneric']")

                if check.get_attribute("value") == "NaN":
                    item3 = outter.item(key, 0, "//input[@id='Grid_-1_txtGeneric']")
                    item3.textfill()
                else:
                    pass

                k = k + 1
                l = l + 1

            driver.switch_to.default_content()
            driver.switch_to.frame("iframe1")
            driver.switch_to.frame("iframe1")

            save_button = outter.field("//button[@id='btnSave']")
            save_button.select()

            driver.switch_to.default_content()
            driver.switch_to.frame("iframe1")

            setting2 = outter.field("//div[@title='Grade Rate Bulk Update']")
            setting2.select()


            driver.switch_to.frame("iframe1")
            iframe3 = driver.find_element(By.XPATH, "//iframe[@src='/backoffice/settings/GradeBulkUpdate_projectselector.asp']")
            driver.switch_to.frame(iframe3)


            driver.switch_to.default_content()
            driver.switch_to.frame("iframe1")
            driver.switch_to.frame("iframe1")
            driver.switch_to.frame(iframe3)

            #project_name = outter.field(f"//option[text()='{project}']")  # change to the variable with a project name
            #project_name.select()

            project_name = outter.field(f"//option[contains(text(), '{project}')]")  # change to the variable with a project name
            project_name.select()

            button = outter.field("//a[@href='javascript:movetoright()']")
            button.select()

            driver.switch_to.default_content()
            driver.switch_to.frame("iframe1")
            driver.switch_to.frame("iframe1")

            if j == 0:

                button1 = outter.field("//input[@id='GetProjects']")
                button1.select()

                button2 = outter.field("//input[@id='btnProcess']")
                button2.select()

            else:

                item5 = item("blank", "Create a new grade card for each project", "//select[@id='runMode']")
                item5.dropdown()

                try:
                    dropdown1 = driver.find_element(By.XPATH, "// select[ @ id = 'startMonth']")
                    dd1 = Select(dropdown1)
                    dd1.select_by_value(dic["Band"][j])

                    button1 = outter.field("//input[@id='GetProjects']")
                    button1.select()

                    button2 = outter.field("//input[@id='btnProcess']")
                    button2.select()
                except:
                    pass


            driver.switch_to.default_content()
            driver.switch_to.frame("iframe1")

            setting1.select()
            x = x + 1
            j = j + 1


