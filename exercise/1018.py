# import pandas as pd
#
# df = pd.DataFrame({'col1': [1, 2, 3], 'col2': [4, 5, 6]})
#
# def my_func(x):
#     return x['col1'] + x['col2']
#
# df['new_col'] = df.apply(my_func, axis=1)
#
# print(df['new_col'])

# import shutil
#
# shutil.copyfile(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\New FA Register 2023.09 - Final.xlsx', r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\New FA Register 2023.09 - Final2.xlsx')

import os
path_template = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202309\系统导入模板\资产2023_10_25_13_47_33.xlsx'

print(os.path.dirname(path_template))