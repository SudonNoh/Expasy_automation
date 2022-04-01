import pandas as pd
from expasy.expasy import SeleniumControl, ExcelControl
import time

data = pd.read_excel('C:/Users/SD NOH/Desktop/excel/SID_original.xlsx', sheet_name='Sheet1')
# adjusted_data = data.drop(data.index[5:])
# [[id, adjusted_seq, seq, structure, dimer], [...]]
adjusted_data = data.dropna()
data_list = adjusted_data.values.tolist()

selenium = SeleniumControl(version="98.0.4758.102")
excel_ctrl = ExcelControl()
selenium.site_enter(url='https://web.expasy.org/protparam')

for idx, val_list in enumerate(data_list):
    selenium.input_seq(val_list[1])
    result = selenium.get_body()
    data_list[idx].append(result)

    mw = round(float(excel_ctrl.string_slice(result, 'Molecular weight:'))/1000, 2)
    pi = excel_ctrl.string_slice(result, 'Theoretical pI:')
    abs_ = excel_ctrl.string_slice(result, 'Abs 0.1% (=1 g/l)')

    data_list[idx].append(mw)
    data_list[idx].append(pi)
    data_list[idx].append(abs_)

    # data_list = [[id, adjusted_seq, seq, structure, dimer, result, mw, pi, abs], [...]]
    selenium.site_back()
    print('\n',
          'ID : ', data_list[idx][0], '\n'
          'adjusted_seq : ', data_list[idx][1], '\n'
          'seq : ', data_list[idx][2], '\n'
          'structure : ', data_list[idx][3], '\n'
          'mw : ', data_list[idx][6], '\n'
          'pi : ', data_list[idx][7], '\n'
          'abs : ', data_list[idx][8], '\n')

    if (idx+1) % 15 == 0:
        time.sleep(10)
        print(idx+1, " 진행 중 입니다.")

    if (idx+1) % 500 == 0:
        df = pd.DataFrame(data_list,
                          columns=['id', 'adjusted_seq', 'seq', 'structure', 'dimer', 'result', 'mw', 'pi', 'abs'])
        df.to_excel('C:/Users/SD NOH/Desktop/excel/SID.xlsx')
        time.sleep(30)

selenium.site_close()

df = pd.DataFrame(data_list,
                  columns=[
                      'id',
                      'adjusted_seq',
                      'seq',
                      'structure',
                      'dimer',
                      'result',
                      'mw',
                      'pi',
                      'abs'
                  ]
                  )
df.to_excel('C:/Users/SD NOH/Desktop/excel/SID.xlsx')
