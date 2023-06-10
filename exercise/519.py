# import os
# file_path = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\GL'
# file_list = os.listdir(file_path)
# file_name = 'SH GL.xlsx'
# print(os.path.join(file_path, file_name))
#
#
# print('apple'.startswith('p', 1, 3))

# import xlwings as xw
# file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202304\PRC New FA Register.xlsx'
# sheet_name = 'Sheet1'
# app = xw.App(visible=True, add_book=False)
# workbook = app.books.open(file_path)
# worksheet = workbook.sheets[sheet_name]
# value = worksheet.range('A1').expand('table').value
# data = dict()
# for i in range(len(value)-1):
#      asset_type = value[i+1][5]
#      if asset_type not in data:
#          data[asset_type] = []
#      data[asset_type].append(value[i+1])
# for key, value in data.items():
#     new_workbook = xw.books.add()
#     new_worksheet = new_workbook.sheets.add(key)
#     new_worksheet['A1'].value = worksheet['A1'].expand('right').value
#     new_worksheet['A2'].value = value
#     new_workbook.save(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202304\\{}.xlsx'.format(key))
# app.quit()

##将多个工作表拆分为多个工作簿
# import xlwings as xw
# file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202212\SHZ Fixed Assets Register 2022.12.xlsx'
# app = xw.App(visible=True, add_book=False)
# print("Begin")
# workbook = app.books.open(file_path)
# print("open successfully")
# try:
#     for i in workbook.sheets:
#         if i.name == 'BneWorkBookProperties' or i.name == 'BneLog':
#             continue
#         print(i.name)
#         workbook_split = app.books.add()
#         worksheet_split = workbook_split.sheets[0]
#         i.api.Copy(Before=worksheet_split.api)
#         workbook_split.save(rf'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202212\ss\SHZ Fixed Assets Register 2022.12_{i.name}.xlsx')
# finally:
#     print("done")
#     workbook.close()
#     app.quit()

##批量合并多个工作簿中的同名工作表
# import xlwings as xw
#
# file_path_A = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202304\for test\A.xlsx'
# file_path_B = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202304\for test\B.xlsx'
# app = xw.App(visible=True, add_book=False)
# workbook_A = xw.books.open(file_path_A)
# worksheet_A = workbook_A.sheets[0]
# workbook_B = xw.books.open(file_path_B)
# worksheet_B = workbook_B.sheets[0]
# worksheet_A['A1'].api.EntireRow.Copy(Destination=worksheet_B['A1'].api)
# row_num = worksheet_B['A1'].current_region.last_cell.row
# worksheet_A['A1'].current_region.offset(1,0).api.Copy(Destination=worksheet_B['A{}'.format(row_num+1)].api)
# workbook_B.save()
# workbook_B.close()
# workbook_A.close()


##精确调整多个工作簿的行高和列宽
import os
import xlwings as xw

file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202304\for test'
file_name = os.listdir(file_path)
app = xw.App(visible=True, add_book=False)
for i in file_name:
    if i.startswith('~$'):
        continue
    file_paths = os.path.join(file_path,i)
    workbook = app.books.open(file_paths)
    for j in workbook.sheets:
        value = j.range('A1').expand('table')
        value.column_width = 20
        value.row_height = 20
        j['A1'].expand('right').api.Font.name = '宋体'
        j['']
    workbook.save()
    workbook.close()
app.quit()
