import pandas as pd
import exercise316

file_path_bank = r'C:\Users\he kelly\Desktop\Bank reconciliation\202302\HSBC TW.xlsx.xlsx'
file_path_APmapping = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\AP Mapping.xlsx'
file_path_GL = r'C:\Users\he kelly\Desktop\Bank reconciliation\202302\GL JAN-FEB TW.xlsx.xlsx'

bankData = pd.read_excel(file_path_bank, sheet_name='Sheet1')
map_Employee = pd.read_excel(file_path_APmapping, sheet_name='Employee')
glData = pd.read_excel(file_path_GL, sheet_name='TW 101245')
filteredBankData = bankData[bankData["TRN type"].str.contains("TRANSFER", "Transfer")]

def get_sub_set(nums):
    sub_sets = [[]]
    for x in nums:
        sub_sets.extend([item + [x] for item in sub_sets])
    return sub_sets

def Mapping_Employee (vendor_name, office):

    if office == '2801':
        Vendor_Site_Name = 'PRC'
    elif office == '2821':
        Vendor_Site_Name = 'SHZ'
    elif office == '2841':
        Vendor_Site_Name = 'BEI'
    elif office == '1601':
        Vendor_Site_Name = 'HKG'
    elif office == '6001':
        Vendor_Site_Name = 'TAI'
    else:
        pass

    tf = (map_Employee["Vendor Name"] == vendor_name) & (map_Employee["Vendor Site Name"] == Vendor_Site_Name)
    result = map_Employee.loc[tf, 'Bank Account Num']
    return result

#red pocket
glData_NoVendorNA=glData.dropna(axis=0, subset="Vendor Name")
# glData_employee = pd.DataFrame(columns=glData.columns)
# for i in range(0,10):
#      addedList = glData_NoVendorNA.loc[glData_NoVendorNA["Vendor Name"].str.endswith(f'{i}')]
#      glData_employee = pd.concat([glData_employee, addedList])
# print(glData_employee)

glData_employee = glData_NoVendorNA.loc[glData_NoVendorNA["Vendor Name"].str.contains("                ")]
glData_redPocket = glData_employee.loc[glData_employee["Memo"].str.contains("red-pocket")]

pro_glData_redPocket = glData_redPocket.groupby("Vendor Name")
Code = str(glData['Entity Cd'].iloc[1])

record_gl_E = []
record_bk_E = []

for i, j in pro_glData_redPocket:
    bankAccountSeries = Mapping_Employee(f"{i}",Code)
    if bankAccountSeries.size:
        bankAccountNumber = bankAccountSeries.iloc[0]
        bankListIndex = []
        for narrative in filteredBankData["Narrative"]:
            narrative_split=narrative.replace("\n", "").replace(" ", "")
            narrative_split = [item for item in narrative_split.split("/")]
            if f"{bankAccountNumber}" in narrative_split:
                bankList = (filteredBankData[filteredBankData["Narrative"].str.contains(f'{narrative}')])
                bankListIndex.append(bankList.index)
        bankListIndex_int = []
        for a in bankListIndex:
             for index in a:
                 bankListIndex_int.append(index)
        bankListIndex = (list(set(bankListIndex_int)))
        print(bankListIndex)
        glValue = j["Amount Avg Rate"].sum()
        subsets_Bank = get_sub_set(bankListIndex)
        print(subsets_Bank)
        for subset in subsets_Bank:
            subsetSum = 0
            if len(subset) >= 1:
                for index in subset:
                    bank = bankData.loc[index, :]
                    value = bank['Credit/Debit amount']
                    subsetSum += value
            if glValue == subsetSum:
                record_gl_E.append(j.index)
                record_bk_E.append(subset)
                break

print(record_gl_E)
print(record_bk_E)

wbBank = openpyxl.load_workbook(file_path_bank)
sheetBank = wbBank.worksheets[0]
for i in record_bk_E:
    for j in i:
        cellBank = sheetBank.cell(j+2, sheetBank.max_column-1)
        cellBank.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='ffc7ce')
wbBank.save(file_path_bank)


wbGL = openpyxl.load_workbook(file_path_GL)
sheetGL = wbGL["TW 101245"]
for i in record_gl_E:
    for j in i:
        cellGL = sheetGL.cell(j+2, sheetGL.max_column-1)
        cellGL.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='ffc7ce')
wbGL.save(file_path_GL)
