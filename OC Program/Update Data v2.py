import openpyxl
import pandas
from openpyxl.utils.dataframe import dataframe_to_rows


#从更新的大表姐中提取dataframe
file_path_newData = r'C:\Users\he kelly\Downloads\GL Dump query Jun.16 _ with RU and Currency Conversion TW 0504.xlsx'
dfNewData = pandas.read_excel(file_path_newData, sheet_name="Drill", header=None)
dfNewData.columns = dfNewData.loc[1]
dfNewData.drop(index=[0, 1], inplace=True)
dfNewData.dropna(how='all', axis=1, inplace=True)



#提取原文件中的dataframe
file_path_Data = r'C:\Users\he kelly\Desktop\OC\2023\TW1.xlsx'
dfData = pandas.read_excel(file_path_Data, sheet_name="Data")
dfData.columns = dfData.loc[0]
dfData.drop(index=0, inplace=True)
dfData.dropna(how='all', axis=1, inplace=True)

#对齐两个dataframe的列数
for i in dfData.columns:
    if i in dfNewData.columns:
        pass
    else:
        dfData.drop(f'{i}', axis=1, inplace=True)


#截取需要更新的数据部分
mergeRows = list(dfData.columns)
del mergeRows[-2]
intersected_df = pandas.merge(dfNewData, dfData, how='inner', on=mergeRows)
intersected_df.drop("Amount Avg Rate_y", axis=1, inplace=True)
intersected_df.rename(columns={"Amount Avg Rate_x":"Amount Avg Rate"}, inplace=True)
dfNewData_filtered = pandas.concat([dfNewData, intersected_df]).drop_duplicates(keep=False)

#将需更新部分贴进excel
wbData = openpyxl.load_workbook(file_path_Data)
wsData = wbData["Data"]
for r in dataframe_to_rows(dfNewData_filtered, index=False, header=False):
    wsData.append(r)
wbData.save(r'C:\Users\he kelly\Desktop\OC\2023\TW2.xlsx')


#Check
sumNewData =dfNewData.loc[dfNewData['Period Name']=='APR-23','Amount Func Cur'].sum()
sumData = dfData.loc[dfData['Period Name']=='APR-23','Amount Func Cur'].sum()+dfNewData_filtered.loc[dfNewData_filtered['Period Name']=='APR-23', 'Amount Func Cur'].sum()
if round(sumNewData, 2) == round(sumData, 2):
    print("Checked")
else:
    print("Error")