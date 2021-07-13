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

from urllib.request import Request, urlopen
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

def get_query():

    URL = 'https://stats.oecd.org/SDMX-JSON/data/MEI_CLI/LOLITOAA.AUT+BEL+EST+FIN+FRA+DEU+GRC+IRL+ITA+LTU+LVA+NLD+PRT+SVK+SVN+ESP+JPN+USA+GBR+MLT+CHN.M/OECD?startTime=1994-01&contentType=csv'

    req = Request(URL, headers={'User-Agent': 'XYZ/3.0'})
    response = urlopen(req, timeout=20)

    file = open("OECD.csv","wb")
    file.write(response.read())
    file.close() 

    df = pd.read_csv('OECD.csv')
    df = df[['Country','Time','Value']]
    df  =df.pivot(index='Time', columns='Country', values='Value') 


def diffusion_index():
    
    csv_file = 'OECD.csv'

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
    csv_file = 'OECD.csv'

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
        self.launch_OECD_Btn.configure(command=get_query)
        self.launch_OECD_Btn.configure(disabledforeground="#a3a3a3")
        self.launch_OECD_Btn.configure(foreground="#000000")
        self.launch_OECD_Btn.configure(highlightbackground="#d9d9d9")
        self.launch_OECD_Btn.configure(highlightcolor="black")
        self.launch_OECD_Btn.configure(pady="0")
        self.launch_OECD_Btn.configure(text='''Query OECD''')

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
    
    vp_start_gui() 

main()
