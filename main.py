import glob
import os
import time

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By


def OECD_scraper():

    #specifying the ChromeDrivers Location
    driver = webdriver.Chrome('driver/chromedriver.exe')

    #specifying the OECD indicators page we will scrape
    driver.get("https://stats.oecd.org/Index.aspx?DataSetCode=MEI_CLI")

    #navigating to the csv download button and clicking download
    element = driver.find_element(By.ID, "export-icon")
    actions = ActionChains(driver)
    actions.double_click(element).perform()
    driver.find_element(By.ID, "ui-menu-0-1").click()
    driver.switch_to.frame(1)
    element = driver.find_element(By.ID, "_ctl12_btnExportCSV")
    actions = ActionChains(driver)
    actions.double_click(element).perform()

    #waiting a few seconds to allow the page to load

    time.sleep(5)
    #asking our webdriver to quit
    driver.quit()

def get_download_path():

    #Returns the default downloads path for linux or windows
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

    #getting the path using our get_download_path() function
    path=get_download_path()

    #specifying we want to get a CSV file
    extension = 'csv'
    os.chdir(path)

    #grabbing a list of files with our specified extension
    files = glob.glob('*.{}'.format(extension))

    #using mtime to get the latest file by modification time to allow for compatibility on UNIX systems
    latest_file = max(files, key=os.path.getmtime)
    print(latest_file)

def main():

    #running the functions to obtain the newest leading indicators
    OECD_scraper()
    #running the function to obtain the most recently downloaded csv file
    get_CSV()

main()

