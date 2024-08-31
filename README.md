An app that automates a process within rapport3 system. (Adding/updating porojects cost grades).

Cost and Grade Rate Automation Bot edited so it works with the GUI application documented after this one

(needs to be updated since some things changed?):

Overview:

This script automates the process of interacting with a web application using Selenium to update cost and grade rates for projects. It reads data from an encrypted Excel file, logs into the web application, and updates rates based on the data provided.

Requirements:

- Selenium: A web testing framework for browser automation.

- Msoffcrypto: A library to decrypt MS Office files.

- Pandas: A data manipulation library in Python.

- ChromeDriver: WebDriver for Chrome browser.

Usage:

1. Ensure you have the required libraries and ChromeDriver installed.

2. Provide the correct file paths and URLs in the appropriate places.

3. Update the 'password.py' file with the decryption password.

Create a .py file with a password variable so you can use it in the main file.

4. Adjust the XPath expressions to match the structure of the target website.
