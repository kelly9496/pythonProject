import os
import pandas
import pandas as pd
from datetime import datetime
import datetime as dt

class ExcelLog:
    def __init__(self, number):
        self.max_log_number = number | 10
        self.log_number = 0

    def log(self, dataframe, desc):
        if self.log_number >= self.max_log_number:
            return
        now = str(datetime.now()).replace(':', '_')
        dataframe.to_excel(rf'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\debug\{desc}_{now}.xlsx')
        self.log_number += 1

excel_log = ExcelLog(10)



# directory_BS = input("Please enter the folder directory of all the BS statements:")
# directory_GL = input("Please enter the folder directory of all the GL files:")
# directory_AP_Vendor = input("Please enter the file link of the AP_Vendor Mapping:")
# directory_AP_Employee = input("Please enter the file link of the AP_Employee Mapping:")
# directory_Commercial = input("Please enter the file link of the Commercial Mapping")


path_folder_BS = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Bank Statement'
path_folder_GL = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\GL'
directory_AP_Vendor = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\AP Mapping.xlsx'
directory_AP_Employee = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\Employee mapping.xlsx'
directory_Commercial = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\Cash receipt 2023.xlsx'


#获取所有bank信息
files_BS = os.listdir(rf'{path_folder_BS}')
df_bank = pd.DataFrame()
for file_BS in files_BS:
    # if file_BS.startswith('~$'):
    #     continue
    file_path_BS = os.path.join(path_folder_BS, file_BS)
    df_file_BS = pd.read_excel(file_path_BS, header=0)
    df_bank = pd.concat([df_bank, df_file_BS])
df_bank.fillna(0, inplace=True)
df_bank['Credit/Debit amount'] = df_bank.apply(lambda row: sum([row['Credit amount'], row['Debit amount']]), axis=1)
df_bank['Value date'] = df_bank['Value date'].astype('datetime64[ns]')




#获取所有GL信息
files_GL = os.listdir(rf'{path_folder_GL}')
df_GL = pd.DataFrame()
for file_GL in files_GL:
    file_path_GL = os.path.join(path_folder_GL, file_GL)
    df_file_GL = pd.read_excel(file_path_GL, header=1).reset_index()
    df_GL = pd.concat([df_GL, df_file_GL])
df_GL['JH Created Date']=df_GL['JH Created Date'].astype('datetime64[ns]').dt.date


#获取mapping
map_vendor = pd.read_excel(directory_AP_Vendor, header=0)
map_employee = pd.read_excel(directory_AP_Employee, header=1)

#读取Commercial mapping, 创建mapping dictionary
map_commercial = pd.read_excel(directory_Commercial, header=0)
map_commercial['Actual Receipt  Amount'].fillna(method='ffill', axis=0, inplace=True)
map_commercial['Receipt Dt'] = map_commercial['Receipt Dt'].astype('datetime64[ns]')
map_commercial['bank expense'] = map_commercial['bank expense'].astype('float')
tb_location = {'088-169370-011': 'PRC', '626-055784-001': 'Beijing', '622-512317-001': 'Shenzhen'}


#定义所需函数
def get_sub_set(nums):
    sub_sets = [[]]
    for x in nums:
        sub_sets.extend([item + [x] for item in sub_sets])
    return sub_sets


#PRC Section
#bank_mapping_PRC = {'626-055784-001': '101001', '622-512317-001': '101135', '088-169370-011': '101244'}
bank_mapping_PRC = {'088-169370-011': '101244'}

for account_number, account_cd in bank_mapping_PRC.items():

    #获取当前bank account的bank和gl数据
    bankData = df_bank[df_bank['Account number']==f'{account_number}']
    glData = df_GL[df_GL['Account Cd']==int(account_cd)]

    #处理无需mapping的type,并筛选需mapping的df
    bankData_charges = bankData[bankData['TRN type']=='CHARGES']
    bankData.loc[bankData_charges.index, 'notes'] = 'bank charges'
    bankData_interest = bankData[bankData['TRN type']=='INTEREST']
    bankData.loc[bankData_interest.index, 'notes'] = 'bank interest'
    bankData_sweep = bankData[bankData['TRN type']=='SWEEP']#sweep 加注释
    sweep_netoff = list(bankData_sweep.index)
    del sweep_netoff[0]
    del sweep_netoff[-1]
    bankData.loc[sweep_netoff, 'notes'] = 'sweep netoff'
    index_filtered = list(set(bankData.index).difference(set(list(bankData_charges.index)+list(bankData_interest.index)+sweep_netoff))) #改名字 index_emptyNotes
    bankData_filtered = bankData.iloc[index_filtered] #改名字


    #commercial mapping
    bankData_commercial = bankData_filtered.loc[bankData_filtered['Credit/Debit amount']>0, :]
    glData_commercial = glData[glData['JE Headers Description'].str.contains('Cash Receipts')]
    location = tb_location[bankData_commercial.loc[0, 'Account number']]
    id_number = 0
    mapped_index_commercial = []
    for ind, row in bankData_commercial.iterrows():
        #按五个筛选标准筛选map_commercial数据
        tf_map_commercial = map_commercial['location'].str.contains(f'{location}') & (map_commercial['Actual Receipt  Amount'] == row['Credit/Debit amount']) & (map_commercial['Receipt Dt'] == row['Value date']) & (map_commercial['Currency'] == 'CNY') & (map_commercial['Notification Email'] != '-')
        df_test = pd.DataFrame()
        df_test['location'] = map_commercial['location'].str.contains(f'{location}')
        df_test['amount'] = map_commercial['Actual Receipt  Amount'] == row['Credit/Debit amount']
        df_test['receipt dt'] = map_commercial['Receipt Dt'] == row['Value date']
        df_test['Currency'] = map_commercial['Currency'] == 'CNY'
        df_test['email'] = map_commercial['Notification Email'] != '-'
        df_test['actual receipt dt']=map_commercial['Receipt Dt']
        df_test['Value date'] = row['Value date']
        print(row['Value date'])
        excel_log.log(df_test, 'test条件')
        map_commercial_filtered = map_commercial.loc[tf_map_commercial, :]
        excel_log.log(map_commercial_filtered, 'step1初步筛选后的mapping表')
        if map_commercial_filtered.size:
            glData_commercial_mapped = pd.DataFrame()
            for i in range(len(map_commercial_filtered)):
                map_commercial_filtered['Notification Email'] = map_commercial_filtered['Notification Email'].astype('datetime64[ns]').dt.date
                date = map_commercial_filtered.iloc[i]['Notification Email']
                project_id = map_commercial_filtered.iloc[i]['Project ID']
                # charges = map_commercial_filtered.iloc[i]['bank expense']
                amount = map_commercial_filtered.iloc[i]['AR in Office Currency']
                s_date = date+dt.timedelta(days=7)
                e_date = date-dt.timedelta(days=7)
                date_condition = (glData_commercial['JH Created Date'] <= s_date) & (glData_commercial['JH Created Date'] >= e_date)
                project_id_condition = glData_commercial['Project Id'] == project_id
                amount_condition = glData_commercial['Amount Func Cur'] == amount
                df_glCommercial_mapped = glData_commercial[date_condition & project_id_condition & amount_condition]
                glData_commercial_mapped = pd.concat([glData_commercial_mapped, df_glCommercial_mapped])
            excel_log.log(glData_commercial_mapped, 'step2gl')
            gl_commercial_sum = glData_commercial_mapped['Amount Func Cur'].sum()
            bk_commercial_amount = row['Credit/Debit amount']
            check = bk_commercial_amount - gl_commercial_sum
            print(check)
            if check == 0:
                id_number = id_number + 1
                bankData.loc[ind, 'notes'] = f'commercial netoff {id_number}'
                glData.loc[glData_commercial_mapped.index, 'notes'] = f'commercial netoff {id_number}'
            if (check != 0) & (check <= 1000):
                df_charges_mapped = glData_commercial[(glData_commercial['Vendor Name'] == '') & (glData_commercial['Amount Func Cur'] == check)]
                if df_charges_mapped.size:
                    id_number = id_number + 1
                    bankData.loc[ind, 'notes'] = f'commercial netoff {id_number}'
                    glData.loc[glData_commercial_mapped.index, 'notes'] = f'commercial netoff {id_number}'
            else:
                continue

    bankData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\bank.xlsx')
    glData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\gl.xlsx')









    #gl mapping columns: 'Project Id' 'Amount Func Cur'

    # vendor_site_OU = {'2801':'China PRC OU', '2821':'Shenzhen OU', '2841':'Beijing OU', '1601':'Hong Kong OU', '6001':'Taiwan OU'}
    # print(vendor_site_OU['2801'])
    # tf = (map_AP["Vendor Name"] == vendor.upper()) & (map_AP["Vendor Site OU"] == vendor_site_OU)
    # bankAccountNum = map_AP.loc[tf, 'Bank Account Num']

    #
    #
    # glData_AP = glData[glData["JE Headers Description"].str.contains("Payments")]
    # pro_glData_AP = glData_AP.groupby("Vendor Name")
    # Code = str(glData_AP['Entity Cd'].iloc[1])
    #
    # record_bk_AP = []
    # record_gl_AP = []
    #
    # for i, j in pro_glData_AP:
    #     a = 1
    #     bankAccountSeries = Mapping_AP(f'{i}', Code)
    #     if bankAccountSeries.size:
    #         bankAccountNumber = bankAccountSeries.iloc[0]
    #         for narrative in filteredBankData["Narrative"]:
    #             if f'{bankAccountNumber}' in narrative:
    #                 bankList = filteredBankData[filteredBankData["Narrative"].str.contains(f'{bankAccountNumber}')]
    #                 bankValueList = bankList["Credit/Debit amount"]
    #                 bankValueList_dic = bankValueList.to_dict()
    #                 glValue = j["Amount Avg Rate"].sum()
    #                 subsets_Bank = get_sub_set(bankValueList_dic)
    #                 for subset in subsets_Bank:
    #                     subsetSum = 0
    #                     if len(subset) >= 1:
    #                         for index in subset:
    #                             subsetSum += bankValueList_dic.get(index)
    #                     if glValue == subsetSum:
    #                         record_gl_AP.append(j.index)
    #                         record_bk_AP.append(subset)
    #                         break
    #
    # glData_Commercial = glData[glData["JE Headers Description"].str.contains("Cash Receipts")]
    # pro_glData_Commercial = glData_Commercial.groupby("Vendor Name")
    #
    # record_bk_C = []
    # record_gl_C = []
    # print("================", a)
    # for i, j in pro_glData_Commercial:
    #     a = 2
    #     bankAccountSeries = map_Commercial.loc[map_Commercial["Client Name"] == f'{i}'.upper(), :]
    #     if bankAccountSeries.size:
    #         bankAccountName = bankAccountSeries["Client Name in Chinese"]
    #         bankListIndex = []
    #         for name in bankAccountName:
    #             pro_name = name.strip()
    #             for narrative in filteredBankData["Narrative"]:
    #                 narrative_split = [item for item in narrative.replace("\n", "").split("/")]
    #                 if f'{pro_name}' in narrative_split:
    #                     bankList = (filteredBankData[filteredBankData["Narrative"].str.contains(f'{narrative}')])
    #                     #               print(bankList)
    #                     bankListIndex.append(bankList.index)
    #         bankListIndex_int = []
    #         for a in bankListIndex:
    #             for index in a:
    #                 bankListIndex_int.append(index)
    #
    #         temp = []
    #         for item in bankListIndex_int:
    #             if not item in temp:
    #                 temp.append(item)
    #         bankListIndex = temp
    #         glValue = j["Amount Avg Rate"].sum()
    #         subsets_Bank = get_sub_set(bankListIndex)
    #         for subset in subsets_Bank:
    #             subsetSum = 0
    #             if len(subset) >= 1:
    #                 for index in subset:
    #                     bank = bankData.loc[index, :]
    #                     value = bank['Credit/Debit amount']
    #                     subsetSum += value
    #             if glValue == subsetSum:
    #                 record_gl_C.append(j.index)
    #                 record_bk_C.append(subset)
    #                 break
    #
    # print("================", a)
    #
    # print("gl_AP")
    # print(record_gl_AP)
    # print("gl_C")
    # print(record_gl_C)
    # print("bk_AP")
    # print(record_bk_AP)
    # print("bk_C")
    # print(record_bk_C)
    #
    # record_gl = record_gl_C + record_gl_AP
    # record_bk = record_bk_C + record_bk_AP
    # print("gl_AP+C")
    # print(record_gl)
    # print("bk_AP+C")
    # print(record_bk)
    #
    # wbBank = openpyxl.load_workbook(file_path_bank)
    # sheetBank = wbBank.worksheets[0]
    # for i in record_bk:
    #     for j in i:
    #         cellBank = sheetBank.cell(j + 2, sheetBank.max_column)
    #         cellBank.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='FFFF50')
    # wbBank.save(file_path_bank)
    #
    # wbGL = openpyxl.load_workbook(file_path_GL)
    # sheetGL = wbGL["Drill"]
    # for i in record_gl:
    #     for j in i:
    #         cellGL = sheetGL.cell(j + 2, sheetGL.max_column - 1)
    #         cellGL.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='FFFF50')
    # wbGL.save(file_path_GL)




