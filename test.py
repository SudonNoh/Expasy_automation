from expasy import SeleniumControl
from expasy import ExcelControl

import pandas as pd

file_route = 'D:/Expasy/excel/TEST.xlsx'
sheet = 'seq'
site_route = 'https://web.expasy.org/protparam'

ec = ExcelControl()

data = ec.excel_read(url=file_route, sheet_name=sheet)

sc = SeleniumControl()

sc.site_enter(site_route)
sc.time_sleep(3)

expasy_data = []
for i in data:
    temp_data = []
    sc.input_seq(i[1])
    sc.time_sleep(5)
    data_text = sc.get_body()
    temp_data = [i[0], data_text]
    expasy_data.append(temp_data)
    print(expasy_data)
    sc.site_back()
    sc.time_sleep(5)

sc.site_close()

df = pd.DataFrame(expasy_data)
df.to_excel('D:/Expasy/excel/TEST2.xlsx')