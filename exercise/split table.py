import pandas as pd

file_path = r'C:\Users\he kelly\Desktop\TW\TW Audit\2023\Tax audit\BCG-TW FY22 Audit Sample_v2.xlsx'
data = pd.read_excel(file_path, sheet_name='Tax Audit Sample 111年')
pro_data = data.groupby('Allocation')
print(pro_data)
for i, j in pro_data:
    new_file_path = 'C:\\Users\\he kelly\\Desktop\\TW\\TW Audit\\2023\\Tax audit\\PWC PBC - ' + i + '.xlsx'
    j.to_excel(new_file_path,sheet_name=F"PWC PBC - {i}",index=False)

# index value type originIndex
# 0 1000.0 transfer 10
valueList = [1000.0]
indexList = [0]
typeList = ['transfer']
originIndexList = [10]

list=[1,2,3]
filterList = []
for item in list:
    if item == 1:
        filterList.append()
class filterItem (item) :
    value=item.value
    type=item.type
    originIndex=originIndex

class FilteredRow():

    def __init__(self,type,value,narrative,originIndex):
        self.type=type
        self.value=value
        self.narrative=narrative
        self.originIndex=originIndex

row1=FilteredRow('transfer',1000.0,'Vendor',10)
