from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date
import pandas as pd
import os

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
        
        # data list를 불러와서 하나의 string으로 추출
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
            
        new_url = self.make_new_url(url)
        
        wb.save(filename=new_url)
        
    def excel_read_2(self, url, sheet_name):
        
        excel_data = pd.read_excel(url, sheet_name=sheet_name)
        data = excel_data.drop(excel_data.index[0:3])
        data1 = data[[data.columns[8]]].dropna(axis=0)
        data2 = data[[data.columns[7]]].dropna(axis=0)
        seq = [j for i in data1.values.tolist() for j in i]
        seq_id = [j for i in data2.values.tolist() for j in i]
        # sequence list로 추출
        return seq, seq_id
        
    def make_excel_file_2(self, data_list, url, sheet_name, seq_id):
        
        # MW, PI, ABS
        data_frame_list = []
        for i, j in zip(data_list, seq_id):
            data = []
            data.append(j)
            data.append(round(float(self.string_slice(i, 'Molecular weight:'))/1000, 2))
            data.append(self.string_slice(i, 'Theoretical pI:'))
            data.append(self.string_slice(i, 'Abs 0.1% (=1 g/l)'))
            data_frame_list.append(data)
            
        df = pd.DataFrame(data_frame_list, columns=['ID', 'MW', 'PI', 'Abs'])
        
        new_url = self.make_new_url(url)
        
        df.to_excel(new_url, sheet_name=sheet_name, index=False)
        
        # 이 다음에 make_excel_file을 실행시키는 식으로 ?
        # 만든 excel 파일을 열어서 다시 붙여넣는 식으로 ?
        
    def string_slice(self, data, string):
        try:
            str_len = len(string)
            start_num = data.find(string)+str_len
            if 'Abs' in string:
                export_string = data[start_num:data.find(',', start_num)]
            else:
                export_string = data[start_num:data.find('\n', start_num)]
        except:
            return ' '
        return export_string.replace(' ', '')
        
    def make_new_url(self, url):
        
        today = date.today().strftime('%y%m%d')
        print('today : ', today)
        
        folder = url[:url.rfind("/")]
        file_list = os.listdir(folder)
        
        count = 1
        if url[url.rfind("/")+1:] in file_list:
            print('file list !!:', file_list)
            for i in file_list:
                print('i !! : ', i)
                if today in i:
                    count += 1
                    print('count !!:', count)
        else:
            count = 1
        
        print('마지막 Count : ', count)
        new_url = url[:-5] + '_' + today + '(' + str(count) + ')' + '.xlsx'
        
        return new_url