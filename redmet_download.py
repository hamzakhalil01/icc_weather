#!/usr/bin/python

####################################################################################################
## 
## ICC Bot to automatically download excel file weather data 
## from https://redmet.icc.org.gt/redmet/comparativas; 
## 
## redmet_download.py : The main python script
## 
## An output file will contain data of Estaciones, Variables within a defined time range and Agrupar
## 
## Usage
## =====
## 
## python redmet_download.py <arguements>
## 
## Arguments
## ==========
##
## -------------------------------------------------------------------------------------------------
## CLI arguments
## -------------------------------------------------------------------------------------------------
## :param '--auto'              `boolean`             : True = Loop every 15 minutes and collected latest 15 minutes of data
## :param '--path'              `string`              : Where to save the data (defaults current folder)
## :param '--csv'               `boolean`             : If True, then the excel file will be converted to CSV (UTF-8)
## :param '--user_name'         `string`              : User name for loggin in to website
## :param '--user_pass'         `string`              : User password for loggin in to website
## :param '--estaciones_list'   `List of strings`     : List of Estaciones to collect
## :param '--variable_list'     `List of strings`     : List of Variables to collect
## :param '--start_date'        `string`              : Start date and time to collect (only if auto is False)
## :param '--end_date'          `string`              : End date and time to collect (only if auto is False)
## :param '--agrupar'           `string`              : Agrupar
## :param '--mergefarms'        `boolean`             : Merge the current, daily, 7-day and 31-day weather to the farm locations
## :param '--farms_file'        `string`              : If merging, name of the file with the farm information
## 
##
## Estaciones and Variable options can be found in the file List of Stations and Variables as of 10 May 2022.txt which comes zipped with this file.
##
## Examples
## ==========
##
##  One time download : 
##
##  python redmet_download.py --user_name "<YOUR USERNAME>" --user_pass "<YOUR PASSWORD>" --estaciones_list "San Nicolas" "Bonanza" "Amazonas" --variable_list "Temperatura (°C) " "Radiacion (w/m²) " "Humedad Relativa (%) " --start_date "01/04/2022 00:00" --end_date "02/04/2022 00:00" --agrupar "Cada 15 minutos" 
##
##
##
## Download and save every 15 minutes
##
## python redmet_download.py --user_name "<YOUR_USERNAME>" --user_pass "<YOUR PASSWORD>" --estaciones_list "San Nicolas" "Bonanza" "Amazonas" --variable_list "Temperatura (°C) " "Radiacion (w/m²) " "Humedad Relativa (%) " --auto True --csv True
##
## Download and Merge with Farm Data
##
## python redmet_download.py --csv True --user_name "<YOUR_USERNAME>" --user_pass "<YOUR PASSWORD>" --estaciones_list "San Nicolas" "Bonanza" "Amazonas" --variable_list "Temperatura (°C) " "Radiacion (w/m²) " "Humedad Relativa (%) " --start_date "01/04/2022 00:00" --end_date "02/04/2022 00:00" --agrupar "Cada 15 minutos" --mergefarms True## 
##
## Dependencies
## ============
##
## Tested on the following;
##
## pandas==1.4.0
## pytz==2021.3
## webdriver-manager==3.5.4
## selenium==4.1.3
## xlrd==2.0.1
##
####################################################################################################
## Date: 21 MAY 22
## Version: 1.1.0
## Status: Development
####################################################################################################


import time
from datetime import datetime, timedelta
import pytz

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import os
import glob
import xlrd
import argparse
import pandas as pd
import numpy as np
from os.path import abspath, expanduser


class RedmetDownload:
    init_url = "https://redmet.icc.org.gt/"
    schedule = "0 * * * *"
    count_files = 0
    count_repeat = 0
    first_load = True

    def __init__(self, auto, path, csv, user_name, user_pass, resort_list, variable_list, start_date,
                 end_date, agrupar, mergefarms, farms_file):

        print("Loading Driver")
        options = webdriver.ChromeOptions()
        options.add_argument("start-maximized")
        options.add_argument("--window-size=1520,900")
        options.add_argument("--disable-notifications")
        options.add_argument("--incognito")
        options.add_argument("--headless")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-web-security")
        options.add_argument("log-level=3")
        options.add_argument("--no-sandbox")
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        self.driver.set_page_load_timeout(240)
        self.driver.implicitly_wait(240)
        print("Driver Loaded")
        self.auto = auto
        self.path = path
        self.csv = csv
        self.user_name = user_name
        self.user_pass = user_pass
        self.resort_list = resort_list
        self.variable_list = variable_list
        self.start_date = start_date
        self.end_date = end_date
        self.agrupar = agrupar
        self.mergefarms = mergefarms
        self.farms_file = farms_file
        print("Variables Loaded")

    def download_file(self):
        while True:

#            try:
            print ("Downloading file")
            if self.count_repeat >= 3:
                raise ValueError("Script doesn't download!")
            self.check_files(True)
            self.login_seth_path()
            self.fill_table_variables_download()
            self.check_files(False, True)
            self.convert_to_csv()
            self.merge_farms()
                    
#            except:
#                if (not self.auto):
#                    print ("Error processing download, it doesn't look like your issue so please try again.")
#                else:
#                    print ("Error processing download, it doesn't look like your issue, we will try again in 15 minutes")

            if (not self.auto):
                print ("Finished")
                break
            print ("Sleeping 15 minutes till next check")
            time.sleep(300)
            print ("10 minutes till next check")
            time.sleep(300)
            print ("5 minutes till next check")
            time.sleep(300)
                

    def login_seth_path(self):
        if self.first_load:
            if self.path == "current directory":
                params = {"behavior": "allow", "downloadPath": os.getcwd()}
            else:
                params = {"behavior": "allow", "downloadPath": abspath(expanduser(self.path))}
            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", params)
        
        self.driver.get(self.init_url)

        if self.first_load:
            self.first_load = False
            
            input_email = self.driver.find_element(By.XPATH, '//input[@placeholder="Email"]').send_keys(self.user_name)
            input_pass = self.driver.find_element(By.XPATH, '//input[@type="password"]').send_keys(self.user_pass)
            load_login = self.driver.find_element(By.XPATH, '//button[@type="submit"]').click()
        
        click_mapas = self.driver.find_element(By.XPATH, '//span[@class="info-box-icon bg-green"]').click()
        time.sleep(5)

    def fill_table_variables_download(self):
        self.driver.find_element(By.XPATH, '//body').send_keys(Keys.ESCAPE)
        
        for resort in self.resort_list:
            enter_resort_text = self.driver.find_element(By.XPATH, '//input[@id="fincas-selectized"]').send_keys(resort)
            enter_resort = self.driver.find_element(By.XPATH, '//input[@id="fincas-selectized"]').send_keys(Keys.ENTER)
            time.sleep(1)
        
        for variable in self.variable_list:
            enter_veriable_text = self.driver.find_element(By.XPATH, '//input[@id="valores-selectized"]').send_keys(variable)
            enter_veriable = self.driver.find_element(By.XPATH, '//input[@id="valores-selectized"]').send_keys(Keys.ENTER)
            time.sleep(1)
            
        # Calculate the Dates to use (if auto)
        
        if (self.auto):
            tz = pytz.timezone("America/Guatemala")
            dt = datetime.now(tz)
            end_date_time = datetime(dt.year, dt.month, dt.day, dt.hour,15*(dt.minute // 15))
#            self.start_date = (end_date_time  - timedelta(hours=0, minutes=15)).strftime("%d/%m/%Y %H:%M")
            self.start_date = (end_date_time  - timedelta(days=31)).replace(hour=0, minute=0).strftime("%d/%m/%Y %H:%M")
            self.end_date = end_date_time.strftime("%d/%m/%Y %H:%M")
        
        delete_start_date = self.driver.find_element(By.XPATH, '//input[@id="txFechaIni"]').send_keys(Keys.CONTROL + "a")
        delete_start_date = self.driver.find_element(By.XPATH, '//input[@id="txFechaIni"]').send_keys(Keys.DELETE)
        enter_start_date = self.driver.find_element(By.XPATH, '//input[@id="txFechaIni"]').send_keys(self.start_date)

        clear_end_date = self.driver.find_element(By.XPATH, '//input[@id="txFechaFin"]').send_keys(Keys.CONTROL + "a")
        clear_end_date = self.driver.find_element(By.XPATH, '//input[@id="txFechaFin"]').send_keys(Keys.DELETE)
        enter_end_date = self.driver.find_element(By.XPATH, '//input[@id="txFechaFin"]').send_keys(self.end_date)

        enter_agrupar = self.driver.find_element(By.XPATH, f'//select[@id="agrupar"]/option[text()="{self.agrupar}"]').click()

        download_excel = self.driver.find_element(By.XPATH, '//button[@id="btnExcel2"]').click()
        
        # Monitor till we see a new file or timeout
        if self.path != "current directory":
            os.chdir(os.path.expanduser(self.path))

        # Wait up to 240 seconds for a file to be created
        timeout = 24
        for i in range(timeout):
            expected_count = self.count_files + 1
            current_count = len(glob.glob('*'))
            if expected_count == current_count:
                
                # Give time for the file to be written
                time.sleep(15)
                break
            time.sleep(10)


    def convert_to_csv(self):
        self.change_name_file()
        
        if self.csv:
            print ("Converting to CSV")
            newest = self.return_last_file()
            workbook = xlrd.open_workbook_xls(newest, ignore_workbook_corruption=True)
            tabs = pd.ExcelFile(workbook).sheet_names
            data_xls = pd.read_excel(workbook, tabs[0],
                                     skiprows=4, skipfooter=4, 
                                     index_col=None)
            csv_file = newest.replace(".xls", ".csv")
#            os.remove(newest)
            data_xls.to_csv(csv_file, encoding='utf-8', date_format='%Y-%m-%d %H:%M')

    def return_last_file(self):
        is_file = glob.glob("*.xls")
        files = sorted(is_file, key=os.path.getmtime)
        newest = files[-1]
        return newest

    def change_name_file(self):
        newest = self.return_last_file()
        current_time = datetime.now().strftime('%Y-%m-%d-%H-%M')
        excel_file = f"IfcmarketsSugar{current_time}.xls"
        os.rename(newest, excel_file)

    def check_files(self, init_check=False, final_check=False):
        if self.path != "current directory":
            os.chdir(os.path.expanduser(self.path))
            
        if init_check:
            self.count_files = len(glob.glob('*'))

        if final_check:
            expected_count = self.count_files + 1
            current_count = len(glob.glob('*'))
            if expected_count != current_count:
                self.count_repeat += 1
                self.download_file()
            else:
                self.count_files = 0
                self.count_repeat = 0

    def stats(self, data_xls, startdate, enddate, my_agg, prefix):
        df = data_xls[(data_xls['Fecha'] >= startdate) & (data_xls['Fecha'] <= enddate)].groupby('Estacion', as_index=False).agg(my_agg)
        df.columns = ['_'.join(col) for col in df.columns]
        df = df.add_prefix(prefix)
        df = df.rename(columns = {prefix + 'Estacion_':'Nearest_station'})
        return df

    def merge_farms(self):
    
        if self.mergefarms:

            print ("Merging aggregate to farms")

            # Read the data
            newest = self.return_last_file()
            workbook = xlrd.open_workbook_xls(newest, ignore_workbook_corruption=True)
            tabs = pd.ExcelFile(workbook).sheet_names
            custom_date_parser = lambda x: datetime.strptime(x, "%Y-%m-%d %H:%M") if len(x) == 16 else datetime.strptime(x, "%d/%m/%y %H:%M")
            data_xls = pd.read_excel(workbook, tabs[0],
                                     skiprows=4, skipfooter=4,
                                     index_col=None, 
                                     parse_dates=['Fecha'], date_parser=custom_date_parser)
            #data_xls['Fecha'] = pd.to_datetime(data_xls['Fecha'], dayfirst=True, format='%Y-%m-%d %H:%M')

            farms = pd.read_csv(self.farms_file, encoding='cp1252')

            # Get rid of data that is known to be bad
            #
            # Tempature less than -10 degC or greater than 50 degC
            data_xls.temperatura=data_xls.temperatura.where(data_xls.temperatura.between(-10,50)) 

            # Get all aggregate headers together
            my_agg = {}
            for variable in data_xls.columns:
                if ((variable != 'Estacion') & (variable !='Fecha')):
                    my_agg[variable] = ['mean', 'min', 'max']

            # Now statistics
            df_current = data_xls[data_xls.groupby('Estacion').Fecha.transform('max') == data_xls['Fecha']]
            df_current = df_current.add_prefix('Now_')
            df_current = df_current.rename(columns = {'Now_Estacion':'Nearest_station'})
            df_current.to_csv('weather_current.csv', encoding='utf-8')
            farms = farms.merge(df_current, how="left", on="Nearest_station")

            # Daily time statistics
            if (self.auto):
                today_now = data_xls.Fecha.max()
            else:
                today_now = datetime.strptime(self.end_date, "%d/%m/%Y %H:%M")
            today_enddate = today_now.strftime('%m/%d/%Y %H:%M')
            today_startdate = today_now.replace(hour=0, minute=0).strftime('%m/%d/%Y %H:%M')
            today_data = self.stats(data_xls, today_startdate, today_enddate, my_agg, 'Today_')
            today_data.to_csv('weather_today.csv', encoding='utf-8')
            farms = farms.merge(today_data, how="left", on="Nearest_station")

            print('Cross check daily filter: ', len(today_data))
            
            # Yesterday statistics
            yesterday_date = today_now - timedelta(days=1)
            yesterday_enddate = yesterday_date.replace(hour=23, minute=59).strftime('%m/%d/%Y %H:%M')
            yesterday_startdate = yesterday_date.replace(hour=0, minute=0).strftime('%m/%d/%Y %H:%M')
            yesterday_data = self.stats(data_xls, yesterday_startdate, yesterday_enddate, my_agg, 'Yesterday_')
            yesterday_data.to_csv('weather_yesterday.csv', encoding='utf-8')
            farms = farms.merge(yesterday_data, how="left", on="Nearest_station")

            # Last 7 days statistics
            sevendays_date = yesterday_date - timedelta(days=7)
            sevendays_enddate = yesterday_date.replace(hour=23, minute=59).strftime('%m/%d/%Y %H:%M')
            sevendays_startdate = sevendays_date.replace(hour=0, minute=0).strftime('%m/%d/%Y %H:%M')
            sevendays_data = self.stats(data_xls, sevendays_startdate, sevendays_enddate, my_agg, '7Days_')
            sevendays_data.to_csv('weather_sevendays.csv', encoding='utf-8')
            farms = farms.merge(sevendays_data, how="left", on="Nearest_station")

            print('Cross check 7 Day filter: ', len(sevendays_data))

            # Last 31 days statistics
            thirtyonedays_date = yesterday_date - timedelta(days=31)
            thirtyonedays_enddate = yesterday_date.replace(hour=23, minute=59).strftime('%m/%d/%Y %H:%M')
            thirtyonedays_startdate = thirtyonedays_date.replace(hour=0, minute=0).strftime('%m/%d/%Y %H:%M')
            thirtyonedays_data = self.stats(data_xls, thirtyonedays_startdate, thirtyonedays_enddate, my_agg, '31Days_')
            thirtyonedays_data.to_csv('weather_thirtyonedays.csv', encoding='utf-8')
            farms = farms.merge(thirtyonedays_data, how="left", on="Nearest_station")

            print('Cross check 31 day filter: ', len(thirtyonedays_data))

            # Only one row per farm/paddock
            farms = farms.groupby(['Farm Name', 'Paddock'], as_index=False).first()
            
            # Save Data
            farms.to_csv(self.farms_file.replace('.csv', '_with_weather.csv'), encoding='utf-8')

            
parser = argparse.ArgumentParser(description='Fill arguments!')
parser.add_argument('--auto', help="Auto download on 15 minute loop",
                    default=False)
parser.add_argument('--path', help="Path to download",
                    default="current directory")
parser.add_argument('--csv', help="Convert to csv",
                    default=False)
parser.add_argument('--user_name', help="User name for login!")
parser.add_argument('--user_pass', help="Pass for login!")
parser.add_argument('--estaciones_list', nargs='+', help="Estaciones in first field",
                    default=["San Nicolas", "Bonanza", "Amazonas"])
parser.add_argument('--variable_list', nargs='+', help="Variable for second field",
                    default=["Temperatura (°C) ", "Radiacion (w/m²) ", "Humedad Relativa (%) "])
parser.add_argument('--start_date', help="Start date",
                    default="01/04/2022 00:00")
parser.add_argument('--end_date', help="End_date",
                    default="01/04/2022 02:00")
parser.add_argument('--agrupar', help="Agrupar",
                    default="Cada 15 minutos")
parser.add_argument('--mergefarms', help="Merge the current, daily, 7-day and 31-day weather to the farm locations",
                    default=False)
parser.add_argument('--farms_file', help="If merging, name of the file with the farm information",
                    default="Weather Stations+Farms.csv")

args = parser.parse_args()

if __name__ == "__main__":
    RedmetDownload(args.auto, args.path, args.csv, args.user_name, args.user_pass, args.estaciones_list,
                   args.variable_list, args.start_date, args.end_date,
                   args.agrupar, args.mergefarms, args.farms_file).download_file()
