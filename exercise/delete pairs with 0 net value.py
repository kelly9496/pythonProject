
#查找重复值
# import pandas as pd
# import xlwings as xw
# app = xw.App(visible=True, add_book=False)
# file_path = r'C:\Users\he kelly\Downloads\GL Dump query Jun.16 _ with RU and Currency Conversion 2022.1-2023.5 HK USD.xlsx'
# book = app.books.open(file_path)
# sheet = book.sheets['FA List']
# value = sheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
# invoice_list = list(value['Supplier InvoiceNo'])
# duplicate = [n for n in invoice_list if invoice_list.count(n)>1]
# output_sheet = book.sheets.add('Output')
# output_sheet['A1'].value = duplicate
# duplicate_list = value.loc[value['Supplier InvoiceNo'].isin(duplicate)]
# result = book.sheets.add('result')
# result['A1'].options(index=False, header=True).value=duplicate_list
# book.save()
# book.close()
# app.quit()

#merge vlookup&dataframe列表内删去加总等0
# import pandas as pd
# import xlwings as xw
# app = xw.App(visible=True, add_book=False)
# file_path_GL = r'C:\Users\he kelly\Downloads\GL Dump query Jun.16 _ with RU and Currency Conversion 2022.1-2023.5 HK USD.xlsx'
# file_path_List = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202304\HK Accrual Tracking - Apr.xlsx'
# book_GL = app.books.open(file_path_GL)
# value_GL = book_GL.sheets['FA List'].range('A1:AG223').options(pd.DataFrame, header=1, index=False).value
# value_GL['Amount Avg Rate'] = value_GL['Amount Avg Rate'].astype('int')
# value_GL_bygroup = value_GL.groupby('Invoice Number')
# record_index=[]
# for i, j in value_GL_bygroup:
#     for index_a, row_a in j.iterrows():
# #        print(index, row, row['Amount Avg Rate'])
#          if index_a in record_index:
#              continue
#          for index_b, row_b in j.iterrows():
#              if index_b in record_index:
#                  continue
#              if row_a['Amount Avg Rate'] + row_b['Amount Avg Rate'] == 0:
#                  record_index.append(index_a)
#                  record_index.append(index_b)
# value_GL_filtered = value_GL.drop(index=record_index)
# book_List = app.books.open(file_path_List)
# sheet_List = book_List.sheets['Accrual List']
# value_List = sheet_List.range('I1:N37').options(pd.DataFrame, header=1, index=False).value
# value_List = value_List.merge(value_GL_filtered[list(value_List.columns)], how='left', on='Invoice Number')
# sheet_Result = book_List.sheets.add('Result')
# sheet_Result['A1'].options(index=False, header=True).value = value_List
# book_List.save()
# book_GL.close()

# 表内删去加总等0
# import pandas as pd
# import xlwings as xw
# app = xw.App(visible=True, add_book=False)
# file_path_GL = r'C:\Users\he kelly\Downloads\GL Dump query Jun.16 _ with RU and Currency Conversion 2022.1-2023.5 HK USD.xlsx'
# book_GL = app.books.open(file_path_GL)
# value_GL = book_GL.sheets['FA List'].range('A1:AG223').options(pd.DataFrame, header=1, index=False).value
# value_GL_leasehold = value_GL[value_GL['Account Cd'] == 163200]
# record_index=[]
# value_GL_leasehold_grouped = value_GL_leasehold.groupby('Invoice Number')
# for i, j in value_GL_leasehold_grouped:
#     for index_a, row_a in j.iterrows():
# #        print(index, row, row['Amount Avg Rate'])
#          if index_a in record_index:
#              continue
#          for index_b, row_b in j.iterrows():
#              if index_b in record_index:
#                  continue
#              if row_a['Amount Avg Rate'] + row_b['Amount Avg Rate'] == 0:
#                  record_index.append(index_a)
#                  record_index.append(index_b)
# value_GL_leasehold = value_GL_leasehold.drop(index=record_index)
# sheet_result = book_GL.sheets.add('result')
# sheet_result['A1'].options(header=1, index=False).value = value_GL_leasehold
# book_GL.save()
