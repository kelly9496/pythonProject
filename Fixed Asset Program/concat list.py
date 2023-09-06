import xlwings as xw
import pandas as pd
import os

file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\HK List'
file_name_newList = 'HK New FA Register.xlsx'#target
file_name_oldList = 'Source.xlsx'#source
file_path_newList = os.path.join(file_path, file_name_newList)
file_name_oldList = os.path.join(file_path, file_name_oldList)


value_newList = pd.read_excel(file_path_newList, sheet_name='Sheet1',  header=0)
print(value_newList)
value_oldList = pd.read_excel(file_name_oldList, sheet_name='for concat',  header=0)
print(value_oldList)
value_concat = pd.DataFrame(columns=value_newList.columns)
for i in range(len(value_oldList)):
    qty = value_oldList.loc[i]['數量']
    df = value_oldList[i:i+1]
    for j in range(int(qty)):
        value_concat = pd.concat([value_concat, df], join='inner')
value_newList = pd.concat([value_newList, value_concat], join='outer')
print(value_newList)

app = xw.App(visible=True, add_book=False)
workbook = app.books.open(file_path_newList)
worksheet = workbook.sheets['Sheet1']
worksheet['A2'].options(index=False,header=False).value = value_newList
# workbook.save(file_path_newList)
workbook.save(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\HK List\IT4.xlsx')
workbook.close()
app.quit()



