import xlwt
import xlrd
data1 = xlrd.open_workbook(r"C:\Users\he kelly\Desktop\Bank reconciliation\202212\HK\101113\Jan. 5.xlsx")
allSheets1 = data1.sheets()
table1 = allSheets1[1]
ncols1 = table1.ncols
print(ncols1)
list1 = table1.col_values(-1)
print(list1)




#list = [1,2,3,4,-1,0,-4,-5,1];
#list.append(1)
record1 = []


#print(len(list));
for aIndex in range(1,len(list1)):
    if aIndex in record1:
        continue
    #print(list[aIndex])
    for bIndex in range(1,len(list1)):
        #print(list[bIndex])
        if bIndex in record1:
            continue
        if (float(list1[aIndex])+float(list1[bIndex]))==0:
            print(list1[aIndex],list1[bIndex])
            record1.append(aIndex)
            record1.append(bIndex)
            break
result1=[]
for a in record1:
    a=a+1
    result1.append(a)


print("第一个文件重复值的位置", result1)



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

workbook.save('result for SH101244 Dec net-offs.xls')