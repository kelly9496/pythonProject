import pandas as pd
import exercise316
import numpy as np



file_path_bank = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\088-169370-011 SH HSBC.xlsx'
file_path_APmapping = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\AP Mapping.xlsx'
file_path_GL = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\GL SH Dec.2022-Jan.2023.xlsx'
#file_path_Cmapping = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\Commercial Mapping.xlsx'

bankData = pd.read_excel(file_path_bank, sheet_name='Sheet1')
map_AP = pd.read_excel(file_path_APmapping, sheet_name='Supplier')
#map_Commercial = pd.read_excel(file_path_Cmapping, sheet_name='Sheet1')
glData = pd.read_excel(file_path_GL, sheet_name='GL')
filteredBankData = bankData[bankData["TRN type"].str.contains("TRANSFER", "transfer")]


def get_sub_set(nums):
    sub_sets = [[]]
    for x in nums:
        sub_sets.extend([item + [x] for item in sub_sets])
    return sub_sets




def Mapping_AP (vendor, office):

    if office == '2801':
        vendor_site_OU = 'China PRC OU'
    elif office == '2821':
        vendor_site_OU = 'Shenzhen OU'
    elif office == '2841':
        vendor_site_OU = 'Beijing OU'
    elif office == '1601':
        vendor_site_OU = 'Hong Kong OU'
    elif office == '6001':
        vendor_site_OU = 'Taiwan OU'
    else:
        pass

    # TODO
    tf = (map_AP["Vendor Name"] == vendor.upper()) & (map_AP["Vendor Site OU"] == vendor_site_OU)
    result = map_AP.loc[tf, 'Bank Account Num']
    return result



glData_AP = glData[glData["JE Headers Description"].str.contains("Payments")]
pro_glData_AP = glData_AP.groupby("Vendor Name")
Code = str(glData_AP['Entity Cd'].iloc[1])

record_bk=[]
record_gl=[]

print("开始循环")
for i, j in pro_glData_AP:
    bankAccountSeries = Mapping_AP(f'{i}', Code)
    if bankAccountSeries.size:
        bankAccountNumber = bankAccountSeries.iloc[0]
        for narrative in filteredBankData["Narrative"]:
             if f'{bankAccountNumber}' in narrative:
                  bankList = filteredBankData[filteredBankData["Narrative"].str.contains(f'{bankAccountNumber}')]
                  bankValueList = bankList["Credit/Debit amount"]
                  bankValueList_dic = bankValueList.to_dict()
                  print(bankValueList_dic)
                  glValue = j["Amount Avg Rate"].sum()
                  subsets_Bank = get_sub_set(bankValueList_dic)
                  print(subsets_Bank)
                  for subset in subsets_Bank:
                      subsetSum = 0
                      if len(subset) >= 1:
                          for index in subset:
                            subsetSum += bankValueList_dic.get(index)
                      if glValue == subsetSum:
                             record_gl.append(j.index)
                             record_bk.append(subset)
                             break

print(record_gl)
print(record_bk)

print("开始上色")
wbBank = openpyxl.load_workbook(file_path_bank)
sheetBank = wbBank.worksheets[0]
for i in record_bk:
    for j in i:
        cellBank = sheetBank.cell(j+2, sheetBank.max_column-1)
        cellBank.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='90EE90')
wbBank.save(file_path_bank)


wbGL = openpyxl.load_workbook(file_path_GL)
sheetGL = wbGL["GL"]
for i in record_gl:
    for j in i:
        cellGL = sheetGL.cell(j+2, sheetGL.max_column-1)
        cellGL.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='90EE90')
wbGL.save(file_path_GL)




#test = bankData[bankData["TRN type"].str.contains("TRANSFER", "transfer")]




