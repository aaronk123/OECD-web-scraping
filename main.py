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
    euro_area = ['Austria', 'Belgium', 'Cyprus', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Ireland', 'Italy',
                 'Latvia', 'Lithuania', 'Luxembourg', 'Malta', 'Netherlands', 'Portugal', 'Slovakia', 'Slovenia', 'Spain']

    # Creating our new dataframe composed of only countries in the Euro Zone
    euDf = df[df['Country'].isin(euro_area)]
    euDf = euDf.reset_index()
    euDf = euDf[['Country', 'TIME', 'Value']]
    euDf = euDf.drop_duplicates(subset=['TIME', 'Country'], keep='last')

    euDf = euDf.pivot(index='TIME', columns='Country', values='Value')
    euDf.columns.name = ' '
    euDf['EU Mean'] = euDf.mean(axis=1)
    euDf = euDf.filter(['EU Mean'])

    # Creating a dataframe consisting of only the United Kingdom
    ukDf = df[df['Country'] == 'United Kingdom']
    ukDf = ukDf.reset_index()
    ukDf = ukDf[['Country', 'TIME', 'Value']]
    ukDf = ukDf.drop_duplicates(subset=['TIME', 'Country'], keep='last')

    ukDf = ukDf.pivot(index='TIME', columns='Country', values='Value')
    ukDf.columns.name = ' '
    ukDf['UK Mean'] = ukDf.mean(axis=1)
    ukDf = ukDf.filter(['UK Mean'])

    # Apply percentage change
    euDf['EU % CHNG'] = euDf['EU Mean'].pct_change() * 100
    ukDf['UK % CHNG'] = ukDf['UK Mean'].pct_change() * 100

    # Merge EU & UK results into one dataframe
    result_df = pd.merge(euDf, ukDf, how='inner', left_index=True, right_index=True)
    #print(result_df)

    # Uk line points
    x1 = result_df.reset_index()['TIME']
    y1 = result_df['UK % CHNG']

    # plotting the Uk line points
    plt.xticks(rotation=45)
    plt.plot(x1, y1, label="UK")

    # EuroZone line points
    x2 = x1
    y2 = result_df['EU % CHNG']

    # plotting the line 2 points
    plt.plot(x2, y2, label="Euro Area")
    plt.xlabel('x - axis')

    # Set the y axis label of the current axis.
    plt.ylabel('% change')

    plt.ylim([-15, 8])

    # Set a title of the current axes.
    plt.title('Monthly Change in Leading Economic Indicators')

    # show a legend on the plot
    plt.legend()

    # Display a figure.
    plt.show()


def main():
    # running the functions to obtain the newest leading indicators
    # OECD_scraper()
    # running the function to obtain the most recently downloaded csv file
    map_oecd_data()


main()
