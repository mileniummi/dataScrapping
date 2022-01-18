import os
import time
import json
import shutil
import pandas as pd
from selenium import webdriver
from unidecode import unidecode
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

municipalities_with_quarters = {}


def latest_download_file():
    path = r'C:\Users\milen\Downloads'
    os.chdir(path)
    files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
    newest = files[-1]

    return newest


url = "https://www.chip.gov.co/schip_rt/index.jsf"
ser = Service('C:\\Users\\milen\\PycharmProjects\\UpWorkDataMining\\chromedriver\\chromedriver_win32'
              '\\chromedriver.exe')
op = webdriver.ChromeOptions()
browser = webdriver.Chrome(service=ser, options=op)

# loading municipalities
xls = pd.ExcelFile(r'C:\Users\milen\PycharmProjects\UpWorkDataMining\Municipalities.xlsx')
sheetX = xls.parse(0)
municipalities = sheetX['Municipio'].values
municipality_codes = sheetX['Municipality Code'].values
try:
    browser.get(url)
    # go to consultas informe
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'j_idt17:j_idt19:j_idt30'))).click()
    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.ID, 'j_idt17:j_idt19:j_idt30:InformacionEnviada:icn'))).click()


    def pre_fill_form(municipality, code):
        WebDriverWait(browser, 20).until(
            EC.presence_of_element_located((By.ID, 'frm1:SelBoxEntidadCiudadano_input'))).send_keys(
            f'{code} - {unidecode(municipality)}')
        entidad_list = WebDriverWait(browser, 20).until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="frm1:SelBoxEntidadCiudadano_div"]//div//*')))

        for entidad in entidad_list:
            if entidad.text.endswith(unidecode(municipality)):
                entidad.click()
                break

        time.sleep(1)
        Select(WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.ID, 'frm1:SelBoxCategoria')))).select_by_value('K21')
        time.sleep(1)


    def fill_form(municipality, code):
        if not os.path.exists(fr'C:\Users\milen\PycharmProjects\UpWorkDataMining\data\{municipality}'):
            os.mkdir(fr'C:\Users\milen\PycharmProjects\UpWorkDataMining\data\{municipality}')
        # fill the from untill quarter
        pre_fill_form(municipality, code)

        # get all quarters available
        select = Select(browser.find_element(By.ID, 'frm1:SelBoxPeriodo'))
        i = 0
        quarters = []
        for quarter in select.options:
            if not i == 0 and int(quarter.get_attribute('value').split('|')[1]) < 2019:
                quarters.append(quarter.get_attribute('value'))
            i = 1
        municipalities_with_quarters[municipality] = quarters

        first = True
        for quarter in quarters:
            if not first:
                pre_fill_form(municipality, code)
            time.sleep(1)
            select = Select(WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.ID, 'frm1:SelBoxPeriodo'))))
            select.select_by_value(quarter)
            quarter_text = select.first_selected_option.text.split(' ')
            quarter_text.pop(1)
            quarter_text = f'{quarter_text[3]}_{quarter_text[0]}_{quarter_text[1]}'
            time.sleep(1)
            Select(WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.ID, 'frm1:SelBoxForma')))).select_by_value('24')
            WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.ID, 'frm1:BtnConsular'))).click()
            WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'frm1:_t224'))).click()
            fileends = "crdownload"
            while "crdownload" == fileends:
                time.sleep(1)
                newest_file = latest_download_file()
                if "crdownload" in newest_file:
                    fileends = "crdownload"
                else:
                    fileends = "none"
            # saving file
            shutil.move(fr'C:\Users\milen\Downloads\{newest_file}',
                        fr'C:\Users\milen\PycharmProjects\UpWorkDataMining\data\{municipality}\FGI_{code}_{quarter_text}.xls')

            WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.ID, 'frm1:j_idt171'))).click()
            first = False


    for i in range(1, 20):
        fill_form(municipalities[i], municipality_codes[i])
        json_object = json.dumps(municipalities_with_quarters, indent=4)
        with open("municipalities_quarters_info.json", "a") as outfile:
            outfile.write(json_object)

except Exception as ex:
    print(ex)

finally:
    browser.close()
    browser.quit()
