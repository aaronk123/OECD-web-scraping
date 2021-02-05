import glob
import os
import time
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

import calendar
from datetime import date
from dateutil.relativedelta import relativedelta

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By


def OECD_scraper():
    # specifying the ChromeDrivers Location
    driver = webdriver.Chrome('driver/chromedriver.exe')

    # specifying the OECD indicators page we will scrape
    driver.get("https://stats.oecd.org/Index.aspx?DataSetCode=MEI_CLI")

    # navigating to the csv download button and clicking download
    element = driver.find_element(By.ID, "export-icon")
    actions = ActionChains(driver)
    actions.double_click(element).perform()
    driver.find_element(By.ID, "ui-menu-0-1").click()
    driver.switch_to.frame(1)
    element = driver.find_element(By.ID, "_ctl12_btnExportCSV")
    actions = ActionChains(driver)
    actions.double_click(element).perform()

    # waiting a few seconds to allow the page to load

    time.sleep(5)
    # asking our webdriver to quit
    driver.quit()


def get_download_path():
    # Returns the default downloads path for linux or windows
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


def get_CSV():
    # getting the path using our get_download_path() function
    path = get_download_path()

    # specifying we want to get a CSV file
    extension = 'csv'
    os.chdir(path)

    # grabbing a list of files with our specified extension
    files = glob.glob('*.{}'.format(extension))

    # using mtime to get the latest file by modification time to allow for compatibility on UNIX systems
    latest_file = max(files, key=os.path.getmtime)

    return latest_file


def map_oecd_data():
    csv_file = get_CSV()

    df = pd.read_csv(csv_file, parse_dates=['Time'], index_col=['Time'], encoding="ISO-8859-1")
    

    # Stating which countries are in the Euro Area for our dataframe later on
    euro_area= ['Austria', 'Belgium', 'Cyprus', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Ireland', 'Italy', 'Latvia', 'Lithuania',
               'Luxembourg', 'Malta', 'Netherlands', 'Portugal', 'Slovakia', 'Slovenia', 'Spain']

    # Creating our new dataframe composed of only countries in the Euro Zone
    euDf=df[df['Country'].isin(euro_area)]

    # Creating a dataframe consisting of only the United Kingdom
    ukDf=df[df['Country'] == 'United Kingdom']

    eu_percen_change = euDf['Value'].pct_change()*100
    uk_percen_change = ukDf['Value'].pct_change()*100

    print(euDf.head())
    print(eu_percen_change)


def main():
    # running the functions to obtain the newest leading indicators
    # OECD_scraper()
    # running the function to obtain the most recently downloaded csv file
    map_oecd_data()


main()
