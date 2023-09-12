import pandas as pd

df = pd.DataFrame({'Name': ['John', 'Mike', 'Lisa'],
                   'Age': [25, 30, 35],
                   'Salary': [50000, 70000, 90000]})

writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

header_format = writer.book.add_format({
    'bold': True,
    'font_color': 'white',
    'bg_color': 'blue',
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'
})

df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, header=False)

worksheet = writer.sheets['Sheet1']
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)