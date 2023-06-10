import pandas as pd
import exercise316
import numpy as np


file_path_bank = r'C:\Users\he kelly\Desktop\Bank reconciliation\202302\HSBC SH.xlsx'
file_path_APmapping = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\AP Mapping.xlsx'
file_path_GL = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\GL SH Dec.2022-Jan.2023.xlsx'
file_path_Cmapping = r'C:\Users\he kelly\Desktop\Bank reconciliation\202301\Commercial Mapping.xlsx'

bankData = pd.read_excel(file_path_bank, sheet_name='Sheet1')
map_AP = pd.read_excel(file_path_APmapping, sheet_name='Supplier')
map_Commercial = pd.read_excel(file_path_Cmapping, sheet_name='Sheet1')
glData = pd.read_excel(file_path_GL, sheet_name='GL')

list = ["SWEEP", "Sweep", "CHARGES", "Charges"]
filteredBankData = bankData
for item in list:
     filteredBankData = filteredBankData.loc[~filteredBankData["TRN type"].str.contains(f"{item}"),:]
print(filteredBankData)


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

    tf = (map_AP["Vendor Name"] == vendor.upper()) & (map_AP["Vendor Site OU"] == vendor_site_OU)
    result = map_AP.loc[tf, 'Bank Account Num']
    return result



glData_AP = glData[glData["JE Headers Description"].str.contains("Payments")]
pro_glData_AP = glData_AP.groupby("Vendor Name")
Code = str(glData_AP['Entity Cd'].iloc[1])

record_bk_AP = []
record_gl_AP = []

for i, j in pro_glData_AP:
    a=1
    bankAccountSeries = Mapping_AP(f'{i}', Code)
    if bankAccountSeries.size:
        bankAccountNumber = bankAccountSeries.iloc[0]
        for narrative in filteredBankData["Narrative"]:
             if f'{bankAccountNumber}' in narrative:
                  bankList = filteredBankData[filteredBankData["Narrative"].str.contains(f'{bankAccountNumber}')]
                  bankValueList = bankList["Credit/Debit amount"]
                  bankValueList_dic = bankValueList.to_dict()
                  glValue = j["Amount Avg Rate"].sum()
                  subsets_Bank = get_sub_set(bankValueList_dic)
                  for subset in subsets_Bank:
                      subsetSum = 0
                      if len(subset) >= 1:
                          for index in subset:
                            subsetSum += bankValueList_dic.get(index)
                      if glValue == subsetSum:
                             record_gl_AP.append(j.index)
                             record_bk_AP.append(subset)
                             break

glData_Commercial = glData[glData["JE Headers Description"].str.contains("Cash Receipts")]
pro_glData_Commercial = glData_Commercial.groupby("Vendor Name")

record_bk_C = []
record_gl_C = []
print("================", a)
for i, j in pro_glData_Commercial:
    a=2
    bankAccountSeries = map_Commercial.loc[map_Commercial["Client Name"] == f'{i}'.upper(), :]
    if bankAccountSeries.size:
          bankAccountName = bankAccountSeries["Client Name in Chinese"]
          bankListIndex=[]
          for name in bankAccountName:
              pro_name=name.strip()
              for narrative in filteredBankData["Narrative"]:
                  narrative_split = [item for item in narrative.replace("\n","").split("/")]
                  if f'{pro_name}' in narrative_split:
                     bankList=(filteredBankData[filteredBankData["Narrative"].str.contains(f'{narrative}')])
      #               print(bankList)
                     bankListIndex.append(bankList.index)
          bankListIndex_int=[]
          for a in bankListIndex:
              for index in a:
                  bankListIndex_int.append(index)
          bankListIndex=(list(set(bankListIndex_int)))
          glValue = j["Amount Avg Rate"].sum()
          subsets_Bank = get_sub_set(bankListIndex)
          for subset in subsets_Bank:
                subsetSum = 0
                if len(subset) >= 1:
                    for index in subset:
                        bank = bankData.loc[index, :]
                        value = bank['Credit/Debit amount']
                        subsetSum += value
                if glValue == subsetSum:
                      record_gl_C.append(j.index)
                      record_bk_C.append(subset)
                      break

print("================", a)

print("gl_AP")
print(record_gl_AP)
print("gl_C")
print(record_gl_C)
print("bk_AP")
print(record_bk_AP)
print("bk_C")
print(record_bk_C)

record_gl = record_gl_C + record_gl_AP
record_bk = record_bk_C + record_bk_AP
print("gl_AP+C")
print(record_gl)
print("bk_AP+C")
print(record_bk)




wbBank = openpyxl.load_workbook(file_path_bank)
sheetBank = wbBank.worksheets[0]
for i in record_bk:
    for j in i:
        cellBank = sheetBank.cell(j+2, sheetBank.max_column-1)
        cellBank.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='FFFF50')
wbBank.save(file_path_bank)


wbGL = openpyxl.load_workbook(file_path_GL)
sheetGL = wbGL["GL"]
for i in record_gl:
    for j in i:
        cellGL = sheetGL.cell(j+2, sheetGL.max_column-1)
        cellGL.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='FFFF50')
wbGL.save(file_path_GL)

#给charges上浅蓝色
#chargesList = bankData[bankData["TRN Type"].str.contains("Charges","CHARGES")]
#
# filteredBankData = bankData[bankData["TRN type"].str.contains("TRANSFER", "transfer")]








