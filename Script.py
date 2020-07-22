from lxml import html
import requests
import csv
import argparse
import json
import time
import os
from bs4 import BeautifulSoup
from datetime import datetime
import logging
from selenium.webdriver.firefox.options import Options
from selenium import webdriver
import pandas as pd
import shutil
import dropbox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import date

# google Api Stuff

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('creditkarma.json', scope)
client = gspread.authorize(creds)

websiteUrl = "https://www.creditkarma.com/auth/logon"
EquifaxUrl = "https://www.creditkarma.com/credit-health/equifax/factors"
WebpageUrl = "https://www.creditkarma.com/credit-health/transunion/factors"
LogoutUrl = "https://www.creditkarma.com/logout/lockdown"
TransunionPdf = "https://www.creditkarma.com/myfinances/creditreport/transunion/view/print#overview"
EquifaxPdf = "https://www.creditkarma.com/myfinances/creditreport/equifax/view/print#overview"
LOGNAME = 'log_CreditKarma.txt'

PdfFolderName = 'Pdf'
PdfFolderPath = os.path.join(os.getcwd(), PdfFolderName)

logging.basicConfig(format='%(message)s', level=logging.INFO)

fileHandler = logging.FileHandler(LOGNAME, mode='w')
consoleHandler = logging.StreamHandler()
rootLogger = logging.getLogger()
rootLogger.addHandler(fileHandler)

# username="Ashleigh.cori@me.com"
# password="Strive1415"

DataBaseFile = os.path.join(os.getcwd(), "SampleDataSheet.xlsx")
DataFile = os.path.join(os.getcwd(), "Data.csv")
DataFileXlsx = os.path.join(os.getcwd(), "DataXlsx.xlsx")
DataBase = []


class TransferData:
    def __init__(self, access_token):
        self.access_token = access_token

    def upload_file(self, file_from, file_to):
        """upload a file to Dropbox using API v2
        """
        dbx = dropbox.Dropbox(self.access_token)

        with open(file_from, 'rb') as f:
            dbx.files_upload(f.read(), file_to, autorename=True)

        dropboxlink = dbx.files_get_temporary_link(file_to)
        return dropboxlink


def SendfileToDropBox(FileLocation, DropLocation):
    access_token = 'gT2-0SxwABsAAAAAAACJwPUqG2KQISVr7tpXyLaNcKQ2ljyvJ5FFdAwdX2XSIx7e'
    transferData = TransferData(access_token)
    # API v2
    dropboxlink = transferData.upload_file(FileLocation, DropLocation)
    return dropboxlink


def RenameFile(Initial_path, fileName):
    # print(os.path.getctime)
    # print(Initial_path)
    # print([os.path.join(Initial_path,f)  for f in os.listdir(Initial_path)])
    filename = max([os.path.join(Initial_path, f) for f in os.listdir(Initial_path)], key=os.path.getctime)
    # print("file: ",filename)
    shutil.move(filename, os.path.join(Initial_path, fileName))


def foxinit():
    options = Options()
    options.headless = True
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.cache.disk.enable", False)
    profile.set_preference("browser.cache.memory.enable", False)
    profile.set_preference("browser.cache.offline.enable", False)
    profile.set_preference("network.http.use-cache", False)
    driver = webdriver.Firefox(profile, options=options)
    return driver


def chromeInit():
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")
    settings = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
             'savefile.default_directory': PdfFolderPath, 'download.directory_upgrade': True}

    chrome_options.add_experimental_option('prefs', prefs)

    chrome_options.add_argument('--kiosk-printing')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--profile-directory=Default')
    chrome_options.add_argument("--disable-plugins-discovery");
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking");
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    path = os.path.join(os.getcwd(), "chromedriver")
    driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=path)
    return driver


def loadData():
    sheet = ''
    try:
        sheet = client.open("Creditkarma").get_worksheet(1)
    except Exception as e:
        print(e)
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    return df


def login(driver, username, password):
    success = True
    try:
        driver.get(websiteUrl)
        time.sleep(4)

        driver.find_element_by_id("username").send_keys(username)
        time.sleep(2)
        driver.find_element_by_id("password").send_keys(password)
        time.sleep(3)
        driver.find_element_by_class_name("logonBtn").click()
        time.sleep(5)

        content = driver.find_element_by_id("log-on-form-section").text
        if "The email or password you entered is incorrect" in content:
            success = False
    except Exception as e:
        print()

    return success


def ScrapeUserData(driver, username, data):
    driver.get(WebpageUrl)
    time.sleep(5)
    # TransUnion
    try:
        score = driver.find_element_by_class_name("credit-health-score-dial").find_elements_by_tag_name("text")[-2].text
        data.append(score)
    except Exception as e:
        print(e)
        data.append("TransunionNotfound")
    try:
        tiles = driver.find_elements_by_class_name("factor-tile-content")
    except Exception as e:
        print(e)

    for tile in tiles:
        try:
            data.append(tile.find_element_by_class_name("f2").text)
        except Exception as e:
            print(e)
            data.append("")
    # Select print window.

    # element=driver.find_elements_by_class_name("credit-health-tab")[1]
    # driver.execute_script("arguments[0].scrollIntoView();", element)
    # time.sleep(2)

    # element.click()
    # time.sleep(3)
    # Select print and download :
    today = date.today()

    # dd/mm/YY
    d1 = today.strftime("%d_%m_%Y")
    driver.get(TransunionPdf)
    time.sleep(5)
    if "error" in driver.current_url:
        data.append("Error Downloading Credit Report.")
    else:
        driver.execute_script('window.print();')
        time.sleep(2)
        # # print("Renaming Start : ")
        username = username.split("@")[0]
        filepdfPath = os.path.join(PdfFolderPath, "{}_transuinion.pdf".format(username))
        RenameFile(PdfFolderPath, filepdfPath)
        time.sleep(2)
        # To dropbox
        try:
            DropBoxLocation = "/CreditKarma/{}_transuinion_{}.pdf".format(username, d1)
            dropboxlink = SendfileToDropBox(filepdfPath, DropBoxLocation)
            data.append(dropboxlink.link)
        except Exception as e:
            print(e)
            data.append("")

    print("Equifax Starts ")
    # Equifax
    equifaxfound = False
    try:
        driver.get(EquifaxUrl)
        time.sleep(5)
        if driver.current_url == WebpageUrl:
            data.append("Equifax Not found")
        else:
            try:
                scorefax = \
                driver.find_element_by_class_name("credit-health-score-dial").find_elements_by_tag_name("text")[-2].text
                data.append(scorefax)
            except Exception as e:
                print(e)
                data.append("Equifax NotFound")
            try:
                tilesfax = driver.find_elements_by_class_name("factor-tile-content")
            except Exception as e:
                print(e)

            for tile in tilesfax:
                try:
                    data.append(tile.find_element_by_class_name("f2").text)
                except Exception as e:
                    print(e)
                    data.append("")

            driver.get(EquifaxPdf)
            time.sleep(3)
            if "error" in driver.current_url:
                data.append("Error Downloading Equifax PDF")
            else:
                driver.execute_script('window.print();')
                time.sleep(2)
                # # print("Renaming Start : ")
                username = username.split("@")[0]
                filepdfPath = os.path.join(PdfFolderPath, "{}_equifax.pdf".format(username))
                RenameFile(PdfFolderPath, filepdfPath)
                time.sleep(2)

                #  to dropbox
                try:
                    DropBoxLocation = "/CreditKarma/{}_equifax_{}.pdf".format(username, d1)
                    dropboxlink = SendfileToDropBox(filepdfPath, DropBoxLocation)
                    data.append(dropboxlink.link)
                except Exception as e:
                    print(e)
                    data.append("")
    except Exception as e:
        print("")

    return data


def csvtoExcel():
    df_csv = pd.read_csv(DataFile)
    df_csv.to_excel(DataFileXlsx, index=False)


def Run(driver, df):
    df_values = list(df.values)
    start_row = 0
    sheet = ''
    try:
        sheet = client.open('Creditkarma').sheet1
        start_row = len(sheet.get_all_values()) + 1
    except Exception as e:
        print(e)
        start_row = 2
    print(sheet, start_row)
    with open(DataFile, "a", ) as file:
        csvwriter = csv.writer(file)
        for i in range(len(df_values)):
            loginSuccess = False
            username = df_values[i][0]
            password = df_values[i][1]
            print(username, password)
            try:
                loginSuccess = login(driver, username, password)
                time.sleep(3)
            except Exception as e:
                print()
                loginSuccess = False
            data = [username, password, loginSuccess]
            # print(loginSuccess)
            if loginSuccess:
                data = ScrapeUserData(driver, username, data)
                # print(data)
                csvwriter.writerow(data)
                sheet.insert_row(data, start_row)
                driver.get(LogoutUrl)
                start_row += 1
            else:
                csvwriter.writerow([username, password, "Login unsuccessful"])
                sheet.insert_row([username, password, "Login unsuccessful"], start_row)
                start_row += 1


if __name__ == '__main__':
    logging.info(" Script Starts : ")
    driver = chromeInit()
    df = loadData()
    try:
        Run(driver, df)
    # csvtoExcel()
    except Exception as e:
        print(e)
        # csvtoExcel()
