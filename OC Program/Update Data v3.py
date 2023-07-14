import pandas
import pandas as pd
import xlwings as xw

# #从更新的大表姐中提取dataframe
file_path_source = r'C:\Users\he kelly\Downloads\GL Dump query Jun.16 _ with RU and Currency Conversion TW 7.6.xlsx'
df_source = pandas.read_excel(file_path_source, sheet_name="Drill", header=1)
# df_source = df_source[df_source['Set of Books Name'].str.contains('Taiwan', na=False, regex=False, case=False)]
print(df_source)
#
# #提取原文件中的dataframe
file_path_target = r'C:\Users\he kelly\Desktop\OC\2023\TW.xlsx'
df_target = pandas.read_excel(file_path_target, sheet_name="Data", header=1)
#
# #
# #
#
# #截取需要更新的数据部分
entryID_source = df_source['JE Header Id'].to_list()
entryID_target = df_target['JE Header Id'].to_list()
entryID_difference = list(set(entryID_source).difference(set(entryID_target)))
df_source_filtered = df_source[df_source['JE Header Id'].isin(entryID_difference)]
df_target_template = pd.DataFrame(columns=list(df_target.columns))
df_source_filtered = pandas.concat([df_target_template, df_source_filtered])
print(df_source_filtered)
# #
# # df_source_filtered.to_excel(r'C:\Users\he kelly\Desktop\OC\2023\TW1.xlsx', sheet_name='New Data', index=False, header=False)
#
app = xw.App(visible=True, add_book=False)
book = app.books.open(file_path_target)
sheet = book.sheets["Data"]
row_num = sheet['A3'].current_region.last_cell.row
print(row_num)
sheet['A{}'.format(row_num+1)].options(index=False, header=False).value=df_source_filtered
book.save()


#
# #Check
# sumNewData =dfNewData.loc[dfNewData['Period Name']=='APR-23','Amount Func Cur'].sum()
# sumData = dfData.loc[dfData['Period Name']=='APR-23','Amount Func Cur'].sum()+dfNewData_filtered.loc[dfNewData_filtered['Period Name']=='APR-23', 'Amount Func Cur'].sum()
# if round(sumNewData, 2) == round(sumData, 2):
#     print("Checked")
# else:
#     print("Error")

