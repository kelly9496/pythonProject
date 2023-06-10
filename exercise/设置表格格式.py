import xlwings as xw
app = xw.App(visible=True, add_book=False)
file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202305\FA Data.xlsx'
workbook = app.books.open(file_path)
worksheet = workbook.sheets['Sheet1']
heading = worksheet.range('A1').expand('right')
print(heading)
heading.font.name = '微软雅黑'
heading.font.size = 12
heading.font.bold = True
heading.font.color = (0,0,0)
heading.color = (193,205,205)
heading.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
heading.api.HorizontalAlignment= xw.constants.HAlign.xlHAlignCenter
content = worksheet.range('A2').expand('table')
print(content)
content.font.name = '微软雅黑'
content.font.size = 12
content.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
content.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
for cell in worksheet['A1'].expand('table'):
    for b in range(7,12):
        cell.api.Borders(b).LineStyle = 1
        cell.api.Borders(b).Weight = 2
workbook.save()
