import xlwings as xw
import pandas as pd
import os

file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\Fixed Asset List\SZ\Fixed Asset List'
file_name_newList = 'SZ New FA Register.xlsx'#target
file_name_oldList = 'SZ Fixed Asset List - Ops @0331.xlsx'#source
file_path_newList = os.path.join(file_path, file_name_newList)
file_name_oldList = os.path.join(file_path, file_name_oldList)

app = xw.App(visible=True, add_book=False)
workbook = app.books.open(file_path_newList)
worksheet = workbook.sheets['Sheet1']
value_newList = worksheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
workbook.close()
workbook = app.books.open(file_name_oldList)
worksheet = workbook.sheets['for concat']
value_oldList = worksheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
workbook.close()
value_concat = value_newList
for i in range(len(value_oldList)):
    qty = value_oldList.loc[i]['16.数量']
    df = value_oldList[i:i+1]
    for j in range(int(qty)):
        value_concat = pd.concat([value_concat, df], join='inner')
for x in value_concat.columns:
    value_newList[f'{x}']=value_concat[f'{x}']

workbook = app.books.open(file_path_newList)
worksheet = workbook.sheets['Sheet1']
worksheet['A2'].options(index=False,header=False).value = value_newList
workbook.save(file_path_newList)




