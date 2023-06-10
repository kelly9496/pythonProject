import xlwings as xw
app = xw.App(visible=True, add_book=False)
file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202305\FA Data.xlsx'
workbook = app.books.open(file_path)
worksheet = workbook.sheets['Sheet1']

# content = worksheet['A1'].expand('table').value #数值类型是一个嵌套的list
# for index, val in enumerate(content):
#     if val[2] == 'Beijing':
#         val[2] = 'BJ'

values = worksheet['A1'].expand()
number = values.shape[1] #values.shape[]
worksheet.range(number+1,1).value = [[5,6,7],[1,2,3]]
print(values)
print(number)



