import pandas as pd
import re

df=pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\HK List\IT4.xlsx', sheet_name='Sheet1', header=0)
df['IT Serial Number'] = df['IT Serial Number'].astype('string')
print(df['IT Serial Number'])

for serial_number, df_bySerialNo in df.groupby('IT Serial Number'):
    if len(df_bySerialNo) == 1:
        continue
    print(serial_number)
    if '\n' in serial_number:
        list_serial_number = serial_number.splitlines()
        number = len(list_serial_number)
    # else:
    #     re_serialNo = re.compile(r'"(.+) Processing')
    #     match_staffName = re_staffName.search(line)
    #     if match_staffName:
    #         staffName = match_staffName.group(1)
    #         list_staffName.append(staffName)
    #     list_serial_number = serial_number.split('\r')
    if list_serial_number:
        number = len(list_serial_number)
        qty = len(df_bySerialNo)
        if number == qty:
            for index, row in df_bySerialNo.iterrows():
                number = number-1
                df.loc[index, 'IT Serial Number'] = list_serial_number[number]

df.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\HK List\IT5.xlsx')


