from expasy import SeleniumControl
from expasy import ExcelControl

import pandas as pd

file_route = 'D:/Expasy/excel/TEST.xlsx'
sheet = 'seq'
site_route = 'https://web.expasy.org/protparam'

ec = ExcelControl()

data = pd.read_excel(file_route, sheet_name=sheet)
data = data.values.tolist()

sc = SeleniumControl()

sc.site_enter(site_route)
sc.time_sleep(3)

expasy_data = []
for i in data:
    sc.input_seq(i[1])
    sc.time_sleep(5)
    data_text = sc.get_body()
    expasy_data.append(data_text)
    print(expasy_data)
    sc.site_back()
    sc.time_sleep(5)

sc.site_close()

string = ''
for i in expasy_data:
    string += i + '\n\n'
    
string2 = string.split(sep='\n')
print(string2)

df = pd.DataFrame(string2)

with pd.ExcelWriter('D:/Expasy/excel/TEST.xlsx', mode='w', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='update', index=False, header=False)
