import xlwings as xw
file_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202212\SHZ Fixed Assets Register 2022.12.xlsx'
app = xw.App(visible=True, add_book=False)
print("Begin")
workbook = app.books.open(file_path)
print("open successfully")
try:
    for i in workbook.sheets:
        if i.name == 'BneWorkBookProperties':
            continue
        print(i.name)
        workbook_split = app.books.add()
        worksheet_split = workbook_split.sheets[0]
        i.copy(before=worksheet_split)
        workbook_split.save(rf'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202212\ss\SHZ Fixed Assets Register 2022.12_{i.name}.xlsx')
finally:
    print("done")
    workbook.close()
    app.quit()