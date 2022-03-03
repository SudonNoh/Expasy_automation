from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date
import pandas as pd


class SeleniumControl:
    def __init__(self, version):
        self.chrome_service = Service(ChromeDriverManager(version=version).install())
        self.options = webdriver.ChromeOptions()
        self.options.add_argument('windows-size=800,1000')
        self.driver = webdriver.Chrome(service=self.chrome_service, options=self.options)

    def site_enter(self, url):
        self.driver.get(url=url)
        self.driver.implicitly_wait(5)

    def input_seq(self, seq):
        self.search = self.driver.find_element(By.XPATH, '//*[@id="sib_body"]/form/textarea')
        self.search.clear()
        self.search.send_keys(seq)
        self.driver.find_element(By.XPATH, '//*[@id="sib_body"]/form/p[1]/input[2]').submit()

    def get_body(self):
        body = self.driver.find_element(By.XPATH, '//*[@id="sib_body"]')
        return body.text

    def time_sleep(self, sec):
        self.driver.implicitly_wait(sec)

    def site_back(self):
        self.driver.back()

    def site_close(self):
        self.driver.quit()


class ExcelControl:

    def excel_read(self, url, sheet_name):
        excel_data = pd.read_excel(url, sheet_name=sheet_name)
        data = excel_data.drop(excel_data.index[0:3])
        data = data[[data.columns[8]]].dropna(axis=0)
        seq = [j for i in data.values.tolist() for j in i]
        # sequence list로 추출
        return seq
    
    def make_excel_file(self, data_list, url, sheet_name):
        string = ''
        for i in data_list:
            string += i+'\n\n'
            
        data_list = string.split('\n')
        df = pd.DataFrame(data_list)
        
        wb = Workbook()
        # wb.create_sheet(title=sheet_name)
        ws = wb['Sheet']
        ws.title = sheet_name
        
        for r in dataframe_to_rows(df, index=False, header=False):
            ws.append(r)
        
        new_url = url[:-5] + '_' + date.today().strftime('%y%m%d') + '.xlsx'
        
        wb.save(filename='D:/Expasy/excel/file.xlsx')
        
        