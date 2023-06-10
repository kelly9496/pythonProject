import xlrd
import xlwt
import pandas as pd
data1 = xlrd.open_workbook(r"C:\Users\he kelly\Desktop\TW\TW Audit\2023\Tax audit\GL Dump query Jun.16 _ 2022.1.1-2023.1.31.xlsx")
#data1 = xlrd.open_workbook("HSBC SH.xlsx")
allSheets1 = data1.sheets()
table1 = allSheets1[0]
ncols1 = table1.ncols
print(ncols1)


list1 = table1.col_values(ncols1-1)
print("list1:", list1)


data2 = xlrd.open_workbook(r"C:\Users\he kelly\Desktop\TW\TW Audit\2023\Tax audit\GL Dump query Jun.16 _ 2022.1.1-2023.1.31.xlsx")
allSheets2 = data2.sheets()
table2 = allSheets2[1]
ncols2 = table2.ncols
print("ncols:", ncols2)
list2 = table2.col_values(ncols2-1)
print("list2:", list2)





#list = [1,2,3,4,-1,0,-4,-5,1];
#list.append(1)
record1 = []
record2 = []

# 将list1
# step1 将名称列和金额列找到
for i in range(ncols1):
    curList = table1.col_values(i)
    column_name=curList[0]
    if column_name == "Reference":
        list_reference = curList
    elif column_name == "Amount":
        list_amount = curList
    else:
        pass
print(list_reference)
print(list_amount)
# step2 将包含TS且名称相同的行找出，找出所有batch，计算batch总和
batch_list=[]
batch_amount_list = []
for i in range(1, len(list_reference)):
    if "TS" in list_reference[i]:
        if list_reference[i] in batch_list:
            batch_index = batch_list.index(list_reference[i])
            batch_amount_list[batch_index] += list_amount[i]
        else:
            batch_list.append(list_reference[i])
            batch_amount_list.append(list_amount[i])
print("batch list:", batch_list, batch_amount_list)

if len(batch_list):
    record_ts=[]
    for batchIndex in range(0,len(batch_amount_list)):
        if batchIndex in record_ts:
            continue
        for bIndex in range(1, len(list2)):
            # print(list[bIndex])
            if bIndex in record2:
                continue
            if float(batch_amount_list[batchIndex]) == float(list2[bIndex]):
                print(batch_amount_list[batchIndex], list2[bIndex])
                record_ts.append(batchIndex)
                record2.append(bIndex)
                break
    print("ts record2:", record2)



#print(len(list));
for aIndex in range(1,len(list1)):
    if aIndex in record1:
        continue
    #print(list[aIndex])
    for bIndex in range(1,len(list2)):
        #print(list[bIndex])
        if bIndex in record2:
            continue
        if float(list1[aIndex])==float(list2[bIndex]):
            print(list1[aIndex],list2[bIndex])
            record1.append(aIndex)
            record2.append(bIndex)
            break
result1=[]
for a in record1:
    a=a+1
    result1.append(a)
result2=[]
for b in record2:
    b=b+1
    result2.append(b)

print("第一个文件重复值的位置", result1)
print("第二个文件重复值的位置",result2)


workbook=xlwt.Workbook()
worksheet=workbook.add_sheet('My Sheet')
pattern = xlwt.Pattern() # Create the Pattern
pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
style = xlwt.XFStyle() # Create the Pattern
style.pattern = pattern # Add Pattern to Style

for aIndex in range(len(list1)):
    if aIndex in record1:
        worksheet.write(aIndex,0,list1[aIndex],style)
    else:
        worksheet.write(aIndex,0, list1[aIndex])

for bIndex in range(len(list2)):
    if bIndex in record2:
        worksheet.write(bIndex,1,list2[bIndex],style)
    else:
        worksheet.write(bIndex,1, list2[bIndex])
workbook.save('result for TW Sample.xls')

