import glob
import os
import sys
import time
from tkinter import ttk

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import addcharts

try:
    import Tkinter as tk
except ImportError:
    import tkinter as tk

try:
    import ttk

    py3 = False
except ImportError:
    import tkinter.ttk as ttk

    py3 = True

import project_support
import xlsxwriter
import openpyxl
from openpyxl import Workbook,load_workbook

import calendar
from datetime import date
from dateutil.relativedelta import relativedelta

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select


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


def second_oecd_scraper():
    driver = webdriver.Chrome('driver/chromedriver.exe')
    driver.get("https://stats.oecd.org/Index.aspx?DataSetCode=MEI_CLI")
    driver.set_window_size(945, 876)

    element_to_hover_over = driver.find_element_by_id("customize-icon")

    hover = ActionChains(driver).move_to_element(element_to_hover_over)
    hover.perform()

    driver.find_element_by_link_text('Time & Frequency [24]').click()

    time.sleep(10)

    # Find the outer iframe and switch to it
    iframe = driver.find_element_by_id('DialogFrame')
    driver.switch_to.frame(iframe)

    # Find the inner iframe and switch to it
    iframe = driver.find_element_by_id('TimeSelector')
    driver.switch_to.frame(iframe)

    # Select number of years wanted
    # driver.execute_script("document.getElementById('cboRelativeAnnual').getElementsByTagName('option')[10].selected = 'selected'")

    # driver.find_element_by_id('lbtnViewData').click()
    # driver.execute_script("document.getElementById('lbtnViewData').click()")


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


def diffusion_index():
    csv_file = get_CSV()

    col_names = ['index', 'Time', 'Country', 'TIME', 'Value', 'Binary']
    diff_index_df = pd.DataFrame(columns=col_names)

    df = pd.read_csv(csv_file, low_memory=False, encoding="ISO-8859-1")

    df = df[['Subject', 'Time', 'Country', 'TIME', 'Value']]
    df = df.loc[df['Subject'] == 'Amplitude adjusted (CLI)']

    df.set_index('Time')

    # getting a list of all countries in our dataframe
    unique_countries = np.unique(df['Country'].astype(str))

    # iterating through each country's data
    for country in unique_countries:

        # creating a specific dataframe for each country
        uniq_country_df = df[df['Country'] == country]
        uniq_country_df = uniq_country_df.reset_index()

        # iterating through each row in a specific country's dataframe
        for index, row in uniq_country_df.iterrows():
            if index != 0:

                if uniq_country_df.loc[index, 'Value'] > uniq_country_df.loc[index - 1, 'Value']:
                    uniq_country_df.loc[index, 'Binary'] = 1

                elif uniq_country_df.loc[index, 'Value'] < uniq_country_df.loc[index - 1, 'Value']:
                    uniq_country_df.loc[index, 'Binary'] = 0

        # removing the first index as it does not have a binary value
        uniq_country_df = uniq_country_df.iloc[1:]
        uniq_country_df.reset_index(drop=True).rename_axis(index=None, columns=None)

        # appending to our final dataframe
        diff_index_df = diff_index_df.append(uniq_country_df, ignore_index=True)

    diff_index_df.drop_duplicates(subset=['TIME', 'Country'])

    p = diff_index_df.groupby(['TIME']).agg({'Binary': 'sum'})

    global_num = p.sum()

    p.sort_values(by='TIME')
    p.reset_index(level=0, inplace=True)

    p.reset_index()

    p.plot(kind='line', x='TIME', y='Binary', legend=None)

    plt.title('Leading Economic Indicator Diffusion Index\n'
              'If LEI reading for the economy improves score =1, if not score = 0')

    plt.savefig('diffusion index.png')

    writer = pd.ExcelWriter('Diffusion index.xlsx', engine='xlsxwriter')
    global_num.to_excel(writer, sheet_name='Sheet1')

    worksheet = writer.sheets['Sheet1']
    worksheet.insert_image('C2', 'diffusion index.png')
    writer.save()


def annual_changes():
    csv_file = get_CSV()

    df = pd.read_csv(csv_file, low_memory=False, encoding="ISO-8859-1")

    df = df[['Subject', 'Time', 'Country', 'TIME', 'Value']]
    df = df.loc[df['Subject'] == 'Amplitude adjusted (CLI)']

    df.set_index('Time')

    ########################
    # Graphing the Euro Zone
    ########################
    euro_area = ['Austria', 'Belgium', 'Cyprus', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Ireland', 'Italy',
                 'Latvia', 'Lithuania', 'Luxembourg', 'Malta', 'Netherlands', 'Portugal', 'Slovakia', 'Slovenia', 'Spain']

    # Creating our new dataframe composed of only countries in the Euro Zone
    eu_Df = df[df['Country'].isin(euro_area)]
    eu_Df = eu_Df.reset_index()
    eu_Df = eu_Df[['Country', 'TIME', 'Value']]
    eu_Df = eu_Df.drop_duplicates(subset=['TIME', 'Country'], keep='last')
    eu_Df = eu_Df.pivot(index='TIME', columns='Country', values='Value')

    eu_Df.columns.name = ' '
    eu_Df['EU Mean'] = eu_Df.mean(axis=1)
    eu_Df = eu_Df.filter(['EU Mean'])

    eu_Df['EU % Change'] = eu_Df['EU Mean'].pct_change() * 100
    eu_Df = eu_Df.reset_index('TIME')
    eu_Df.reset_index(level=0, inplace=True)
    eu_Df = eu_Df[['TIME', 'EU Mean', 'EU % Change']]

    # Create a workbook
    workbook = xlsxwriter.Workbook('Annual Changes.xlsx')

    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter('Annual Changes.xlsx', engine='xlsxwriter')

    # Write DataFrame to workbook sheet
    eu_Df.to_excel(writer, sheet_name='Euro-Zone', startcol=0, index=False)

    # Set worksheet
    workbook = writer.book
    worksheet = writer.sheets['Euro-Zone']

    # Create a chart object.
    chart = workbook.add_chart({'type': 'line'})

    # Configure the series of the chart from the dataframe data.
    chart.add_series({
        'categories': ['Euro-Zone', 1, 0, len(eu_Df) + 2, 0],
        'values': ['Euro-Zone', 1, 2, len(eu_Df) + 2, 2],
    })

    # Insert the chart into the worksheet.
    worksheet.insert_chart('E2', chart)

    #############################
    # Graphing singular countries
    ##############################

    desired_countries = ["Ireland", "United Kingdom", "United States", "Japan", "China (People's Republic of)"]

    # Creating our dataframe of the other countries we wish to use
    df = df[df['Country'].isin(desired_countries)]

    for country in desired_countries:
        # creating a specific dataframe for each country
        uniq_country_df = df[df['Country'] == country]
        uniq_country_df = uniq_country_df.reset_index()

        uniq_country_df['Value'] = uniq_country_df['Value'].pct_change() * 100

        country_name = uniq_country_df.iloc[-1]['Country']

        uniq_country_df.reset_index(level=0, inplace=True)

        uniq_country_df.plot(kind='line', x='TIME', y='Value', legend=None).axhline(y=0, color='Grey', linestyle='--')

        # Add DataFrame to workbook
        uniq_country_df.to_excel(writer, sheet_name=country_name, startcol=0, index=False)

        # Set worksheet
        workbook = writer.book
        worksheet = writer.sheets[country_name]

        # Create a chart object.
        chart = None
        chart = workbook.add_chart({'type': 'line'})

        # Configure the series of the chart from the dataframe data.
        chart.add_series({
            'categories': [country_name, 1, 5, len(uniq_country_df) + 2, 5],
            'values': [country_name, 1, 6, len(uniq_country_df) + 2, 6],

        })
        # Insert the chart into the worksheet.
        worksheet.insert_chart('I2', chart)

    # Save the writer
    writer.save()


def vp_start_gui():
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = tk.Tk()
    top = Toplevel1(root)
    project_support.init(root, top)
    root.mainloop()


w = None


def create_Toplevel1(rt, *args, **kwargs):
    '''Starting point when module is imported by another module.
       Correct form of call: 'create_Toplevel1(root, *args, **kwargs)' .'''
    global w, w_win, root
    # rt = root
    root = rt
    w = tk.Toplevel(root)
    top = Toplevel1(w)
    project_support.init(w, top, *args, **kwargs)
    return (w, top)


def destroy_Toplevel1():
    global w
    w.destroy()
    w = None


class Toplevel1:
    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9'  # X11 color: 'gray85'
        _ana1color = '#d9d9d9'  # X11 color: 'gray85'
        _ana2color = '#ececec'  # Closest X11 color: 'gray92'
        self.style = ttk.Style()
        if sys.platform == "win32":
            self.style.theme_use('winnative')
        self.style.configure('.', background=_bgcolor)
        self.style.configure('.', foreground=_fgcolor)
        self.style.configure('.', font="TkDefaultFont")
        self.style.map('.', background=
        [('selected', _compcolor), ('active', _ana2color)])

        top.geometry("555x405+625+196")
        top.minsize(120, 1)
        top.maxsize(3604, 1061)
        top.resizable(1, 1)
        top.title("OECD Graph Generator")
        top.configure(background="#d9d9d9")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="black")

        self.launch_OECD_Btn = tk.Button(top)
        self.launch_OECD_Btn.place(relx=0.414, rely=0.121, height=24, width=77)
        self.launch_OECD_Btn.configure(activebackground="#ececec")
        self.launch_OECD_Btn.configure(activeforeground="#000000")
        self.launch_OECD_Btn.configure(background="#d9d9d9")
        self.launch_OECD_Btn.configure(command=second_oecd_scraper)
        self.launch_OECD_Btn.configure(disabledforeground="#a3a3a3")
        self.launch_OECD_Btn.configure(foreground="#000000")
        self.launch_OECD_Btn.configure(highlightbackground="#d9d9d9")
        self.launch_OECD_Btn.configure(highlightcolor="black")
        self.launch_OECD_Btn.configure(pady="0")
        self.launch_OECD_Btn.configure(text='''Launch OECD''')

        self.menubar = tk.Menu(top, font="TkMenuFont", bg=_bgcolor, fg=_fgcolor)
        top.configure(menu=self.menubar)

        self.TSeparator1 = ttk.Separator(top)
        self.TSeparator1.place(relx=0.0, rely=0.244, relwidth=0.991)

        self.TSeparator2 = ttk.Separator(top)
        self.TSeparator2.place(relx=0.0, rely=0.415, relwidth=0.991)

        self.diff_Index_Btn = tk.Button(top)
        self.diff_Index_Btn.place(relx=0.234, rely=0.296, height=24, width=87)
        self.diff_Index_Btn.configure(activebackground="#ececec")
        self.diff_Index_Btn.configure(activeforeground="#000000")
        self.diff_Index_Btn.configure(background="#d9d9d9")
        self.diff_Index_Btn.config(command=diffusion_index)
        self.diff_Index_Btn.configure(disabledforeground="#a3a3a3")
        self.diff_Index_Btn.configure(foreground="#000000")
        self.diff_Index_Btn.configure(highlightbackground="#d9d9d9")
        self.diff_Index_Btn.configure(highlightcolor="black")
        self.diff_Index_Btn.configure(pady="0")
        self.diff_Index_Btn.configure(text='''Diffusion Index''')

        self.ann_Change_Btn = tk.Button(top)
        self.ann_Change_Btn.place(relx=0.577, rely=0.296, height=24, width=97)
        self.ann_Change_Btn.configure(activebackground="#ececec")
        self.ann_Change_Btn.configure(activeforeground="#000000")
        self.ann_Change_Btn.configure(background="#d9d9d9")
        self.ann_Change_Btn.configure(command=annual_changes)
        self.ann_Change_Btn.configure(disabledforeground="#a3a3a3")
        self.ann_Change_Btn.configure(foreground="#000000")
        self.ann_Change_Btn.configure(highlightbackground="#d9d9d9")
        self.ann_Change_Btn.configure(highlightcolor="black")
        self.ann_Change_Btn.configure(pady="0")
        self.ann_Change_Btn.configure(text='''Annual Change''')


def main():
    # running the functions to obtain the newest leading indicators
    # OECD_scraper()
    # running the function to obtain the most recently downloaded csv file
    # map_oecd_data()
    # second_oecd_scraper()
    vp_start_gui()
    # annual_changes()
    #diffusion_index()


main()
