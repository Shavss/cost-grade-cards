# Cost and Grade Rate Automation Bot with GUI Integration

## Overview

This repository contains an application that automates the process of adding or updating project cost grades within the rapport3 system. The core of this application is the `Cost_Grade_Bot` script, which interacts with a web application using Selenium to update cost and grade rates based on data from an encrypted Excel file. Additionally, a graphical user interface (GUI) has been integrated to provide a more user-friendly experience.

## Script Overview

The `Cost_Grade_Bot` script automates:

1. **Reading Data**: Extracts project information from an encrypted Excel file.
2. **Web Interaction**: Logs into the web application and navigates to the appropriate sections.
3. **Updating Rates**: Updates cost and grade rates based on the provided data.

### Script Requirements

- **Selenium**: A web testing framework for browser automation.
- **msoffcrypto**: A library to decrypt MS Office files.
- **Pandas**: A Python library for data manipulation.
- **ChromeDriver**: WebDriver for Chrome browser automation.

### Script Usage

1. Ensure the required libraries and ChromeDriver are installed.
2. Update file paths and URLs in the script to match your environment.
3. Add the decryption password in a `password.py` file, structured as a Python variable.
4. Adjust XPath expressions as necessary to match the structure of the target website.

## Application Overview

This application provides a GUI for the `Cost_Grade_Bot` script, allowing users to easily input project numbers and initiate the bot to update cost and grade rates.

### Application Requirements

- **customtkinter**: A library for creating modern, themed GUIs in Tkinter.
- **main_adjusted_for_the_app.py**: The customized `Cost_Grade_Bot` script integrated with the GUI.

### Application Usage

1. Install the `customtkinter` library.
2. Ensure the `main_adjusted_for_the_app.py` script is available in the same directory as the application.
3. Update the `project_entry` placeholder text if needed to match user instructions.
4. Launch the application, enter project details (separated by spaces), and run the bot.
