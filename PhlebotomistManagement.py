import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import OrderManagement as OM
import os
#import win32com.client as comclt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path
from selenium.webdriver.common import keys


FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
Url = str(RWDE.ReadData(FilePath, Sheet,3, 3))

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('disable-notifications')
driver = webdriver.Chrome(executable_path = str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver', options=chrome_options)
driver.maximize_window()
driver.get(Url)
#print(driver.title)

FilePath = str(Path().resolve()) + '\Excel Files\PhlebotomistManagement.xlsx'
Seconds = 1

#1. This is for SFDC Login

Sheet = 'Login Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)

for RowIndex in range(2, RowCount + 1):

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 2))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 3))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log in"]', 60)
    Element.click()

    Sheet1 = 'Phlebotomist page Data'
    RowCount1 = RWDE.RowCount(FilePath, Sheet1)
    for Rowindex1 in range(2, RowCount1 + 1):

        # Order Number
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Enter Order Number"]', 60)
        if (str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 2)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 2))
        else:
            Element.send_keys('')

        # Scan Order Button
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Scan Order"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # Scan BCK Id
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Scan BCK"]', 60)
        if (str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 3)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 3))
        else:
            Element.send_keys('')

        # Next Button
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Next"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # Tube1 TextBox
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div/div/div[1]/div[1]/lightning-input//input', 60)
        Tubes = Element.get_attribute('value')

        # Tube2 TextBox
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div[1]/lightning-input/div/input', 60)
        Tubes += ',' + Element.get_attribute('value')

        # Tube3 TextBox
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div/div/div[3]/div[1]/lightning-input//input', 60)
        Tubes += ',' + Element.get_attribute('value')

        # Tube4 TextBox
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[4]/div[1]/lightning-input/div/input', 60)
        Tubes += ',' + Element.get_attribute('value')

        # Submit Button
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Submit"]', 60)
        driver.execute_script('arguments[0].scrollIntoView(false);', Element)
        driver.execute_script('arguments[0].click();', Element)

        # Message
        #time.sleep(Seconds)
        #Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div/div/div/div/div/div/span', 60)
        #if (str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 5)) == Element.text):
        #    RWDE.WriteData(FilePath, Sheet1, Rowindex1, 6, Element.text)
        #    RWDE.WriteData(FilePath, Sheet1, Rowindex1, 7, 'Passed')
        #else:
        #    RWDE.WriteData(FilePath, Sheet1, Rowindex1, 6, Element.text)
        #    RWDE.WriteData(FilePath, Sheet1, Rowindex1, 7, 'Failed')

        # Go Back to Home Page Button
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Go Back to Home Page"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        time.sleep(Seconds)
        if (Rowindex1 == RowCount):
            # Account Icon
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div/button/div/div[1]/p', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Log Out Link
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Log Out"]', 60)
            driver.execute_script('arguments[0].click();', Element)










