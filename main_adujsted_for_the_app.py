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

# Import the decryption password from password.py
import password

class CostGradeBot:
    """Automation bot for updating cost and grade rates."""

    def __init__(self):
        """Initialize the bot and load necessary data."""
        # Decrypt Excel file
        decrypted = io.BytesIO()
        with open('<ENCRYPTED_FILE_PATH>', 'rb') as f:
            file = msoffcrypto.OfficeFile(f)
            file.load_key(password=password.file_password)
            file.decrypt(decrypted)

        # Read Excel data
        self.df = pd.read_excel(decrypted, "<SHEET_NAME>")
        self.dic = self.df.to_dict()
        self.keys = list(self.dic)
        self.counter = sum(1 for val in self.dic["Filters1"].values() if not pd.isna(val))
        self.k = 1
        self.key = self.keys[self.k]

    def truncate(self, n, decimals=0):
        """Truncate a number to the specified decimal places."""
        multiplier = 10 ** decimals
        return int(n * multiplier) / multiplier

class Item:
    """Class representing an item in the form."""

    def __init__(self, key, input, xpath):
        self.key = key
        self.input = input
        self.xpath = xpath

    def dropdown(self):
        """Selects an option from a dropdown menu."""
        dropdown = driver.find_element(By.XPATH, self.xpath)
        dd = Select(dropdown)
        dd.select_by_visible_text(self.input)

    def textfill(self):
        """Fills a text field."""
        textfield = driver.find_element(By.XPATH, self.xpath)
        textfield.clear()
        textfield.send_keys(self.input)

class Field:
    """Class representing a field in the form."""

    def __init__(self, xpath):
        self.xpath = xpath

    def select(self):
        """Clicks the field."""
        select = driver.find_element(By.XPATH, self.xpath)
        select.click()

def start_cards(self, project):
    """Starts the process of updating cost and grade rates."""

    # Initialize the bot
    outter = CostGradeBot()

    # Set up Chrome WebDriver
    PATH = "<CHROME_DRIVER_PATH>"
    script_directory = pathlib.Path().absolute()
    chrome_options = Options()
    chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")

    global driver
    driver = webdriver.Chrome(PATH, options=chrome_options)
    chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")
    driver.maximize_window()

    driver.get("<RAPPORT3_URL>")
    driver.implicitly_wait(4)

    # Switching to an iframe named "iframe1"
    WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it("iframe1"))

    # Selecting the "Projects" option
    setting = outter.field("//div[@title='Projects']")
    setting.select()

    # Selecting the "Test" option
    setting1 = outter.field("//div[@title='Test']")
    setting1.select()

    # These values will depend on the period of which you want to start from
    j = 3
    x = 3

    # Loop through the DataFrame rows starting from index 3
    while x < len(outter.df):

        driver.switch_to.frame("iframe1")

        # Select Branch - DCAL
        branch = outter.item("DCAL", "DCAL" ,"//select[@id='BranchID']")
        branch.dropdown()

        driver.switch_to.frame("summaryframe")

        l = 0
        k = 1
        key = outter.keys[k]

        # Filling out the Plan Cost Grade Rates per band
        while k <= 12:
            field1 = outter.field(f"//div[@id='DefRates_{l}_GradeCostRate']")
            field1.select()

            item1 = outter.item(key, outter.truncate(outter.dic[key][j], 2), "//input[@id='Grid_-1_txtGeneric']")
            item1.textfill()

            k = k + 1
            key = outter.keys[k]
            l = l + 1

        # Filling out the Plan Charge Grade Rates per band
        l = 0
        k = 1
        key = outter.keys[k]

        while k <= 12:
            field2 = outter.field(f"//div[@id='DefRates_{l}_GradeChargeRate']")
            field2.select()

            profit = outter.dic["Profit"][j]
            item2 = outter.item(key, outter.dic[key][j] / profit, "//input[@id='Grid_-1_txtGeneric']")
            item2.textfill()

            k = k + 1
            key = outter.keys[k]
            l = l + 1

        l = 0
        k = 1

        # Checking for NaN in Margin
        while k <= 12:
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

        # Selecting and clicking the save button
        save_button = outter.field("//button[@id='btnSave']")
        save_button.select()

        driver.switch_to.default_content()
        driver.switch_to.frame("iframe1")

        # Selecting the "Grade Rate Bulk Update" option
        setting2 = outter.field("//div[@title='Grade Rate Bulk Update']")
        setting2.select()

        # Switching to nested iframes
        driver.switch_to.frame("iframe1")

        iframe3 = driver.find_element(By.XPATH, "//iframe[@src='/backoffice/settings/GradeBulkUpdate_projectselector.asp']")
        driver.switch_to.frame(iframe3)

        driver.switch_to.default_content()
        driver.switch_to.frame("iframe1")
        driver.switch_to.frame("iframe1")
        driver.switch_to.frame(iframe3)

        # Handling project names
        nr = len(project.split())

        if nr > 1:
            m = 0
            projects = project.split()

            while m < len(projects):
                project_name = outter.field(f"//option[contains(text(), '{projects[m]}')]")
                project_name.select()

                button = outter.field("//a[@href='javascript:movetoright()']")
                button.select()

                m = m + 1
        else:
            project_name = outter.field(f"//option[contains(text(), '{project}')]")
            project_name.select()

            button = outter.field("//a[@href='javascript:movetoright()']")
            button.select()

        driver.switch_to.default_content()
        driver.switch_to.frame("iframe1")
        driver.switch_to.frame("iframe1")

        if j == 0:
            # Clicking buttons
            button1 = outter.field("//input[@id='GetProjects']")
            button1.select()

            button2 = outter.field("//input[@id='btnProcess']")
            button2.select()

            time.sleep(15)
        else:
            # Handling dropdown and buttons
            item5 = outter.item("blank", "Create a new grade card for each project", "//select[@id='runMode']")
            item5.dropdown()

            try:
                dropdown1 = driver.find_element(By.XPATH, "//select[@id='startMonth']")
                dd1 = Select(dropdown1)
                dd1.select_by_value(outter.dic["Band"][j])

                button1 = outter.field("//input[@id='GetProjects']")
                button1.select()

                button2 = outter.field("//input[@id='btnProcess']")
                button2.select()

                time.sleep(15)
            except:
                pass

        driver.switch_to.default_content()
        driver.switch_to.frame("iframe1")

        # Selecting the "Test" option
        setting1.select()

        x = x + 1
        j = j + 1

        # Setting the rates to default ones (current)
        driver.switch_to.frame("iframe1")

        # Select Branch - DCAL
        branch = outter.item("DCAL", "DCAL", "//select[@id='BranchID']")
        branch.dropdown()

        driver.switch_to.frame("summaryframe")

        l = 0
        k = 1
        j = 6
        key = outter.keys[k]

        # Filling out the Plan Cost Grade Rates per band
        while k <= 11:
            field1 = outter.field(f"//div[@id='DefRates_{l}_GradeCostRate']")
            field1.select()

            item1 = outter.item(key, outter.truncate(outter.dic[key][j], 2), "//input[@id='Grid_-1_txtGeneric']")
            item1.textfill()

            k = k + 1
            key = outter.keys[k]
            l = l + 1

        # Filling out the Plan Charge Grade Rates per band
        l = 0
        k = 1
        key = outter.keys[k]

        while k <= 11:
            field2 = outter.field(f"//div[@id='DefRates_{l}_GradeChargeRate']")
            field2.select()

            multiplier = outter.dic["Multiplier"][j]

            item2 = outter.item(key, outter.dic[key][j] * multiplier, "//input[@id='Grid_-1_txtGeneric']")
            item2.textfill()

            k = k + 1
            key = outter.keys[k]
            l = l + 1

        driver.switch_to.default_content()
        driver.switch_to.frame("iframe1") 
        driver.switch_to.frame("iframe1") 
        
        save_button = outter.field("//button[@id='btnSave']") 
        save_button.select()
