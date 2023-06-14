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
df_bank['Value date']=df_bank['Value date'].apply(lambda x: datetime.strptime(x, '%d/%m/%Y'))




#获取所有GL信息
files_GL = os.listdir(rf'{path_folder_GL}')
df_GL = pd.DataFrame()
for file_GL in files_GL:
    file_path_GL = os.path.join(path_folder_GL, file_GL)
    df_file_GL = pd.read_excel(file_path_GL, header=1).reset_index()
    df_GL = pd.concat([df_GL, df_file_GL])


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


def common_data(list1, list2):
    result = False
    for x in list1:
        for y in list2:
            if x == y:
                result = True
    return result

#输入字典键的子集，返回每个键子集对应的字典值的列表的集合
# def key_to_value(subsets_key, dict):
#     subsets_value=[]
#     for subset in subsets_key:
#         subset_value=[]
#         for key in subset:
#             value = dict[key]
#             subset_value.append(value)
#         subsets_value.append(subset_value)
#     return subsets_value

# 输入字典键的列表，返回每个键对应的字典值的列表
def key_to_value(subset_key, dict):
    subset_value = []
    for key in subset_key:
        value = dict[key]
        subset_value.append(value)
    return subset_value




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
    bankData_commercial = bankData_filtered.loc[bankData_filtered['Credit/Debit amount']>0, :] #排除bk金额小于等于0.03的item
    glData_commercial = glData[glData['JE Headers Description'].str.contains('Cash Receipts')]
    location = tb_location[bankData_commercial.loc[0, 'Account number']]
    #处理commercial mapping表，筛出本entity RPApo账的部分，并按project ID分类
    map_commercial_RPA = map_commercial[map_commercial['location'].str.contains(f'{location}') & (map_commercial['Currency'] == 'CNY') & (map_commercial['Notification Email'] != '-')]
    excel_log.log(map_commercial_RPA, 'map_commercial_PRC')
    map_commercial_RPA = map_commercial_RPA.groupby(['Receipt Dt', 'Client Name'])

    #设定初始值
    id_number = 0
    mapped_index_commercial = []
    bank_charges = 0 #bk金额小于等于0.03 或 bk-gl金额小于等于0.03
    mapped_glIndex_commercial = []
    mapped_bankIndex_commercial = []

    #第一轮commercial mapping
    for ind, row in bankData_commercial.iterrows():
        # if ind != 833:
        #     continue
        bank_value = row['Credit/Debit amount']
        bank_receipt_date = row['Value date']
        for group_condition, df_map in map_commercial_RPA:
            if group_condition[0] == bank_receipt_date:
                map_sum_byProject = df_map.groupby('Project ID').sum('AR in Office Currency')['AR in Office Currency'].to_dict()
                project_code_list = list(map_sum_byProject.keys())
                project_code_subsets = get_sub_set(project_code_list)
                for subset in project_code_subsets:
                    subset_value_map = key_to_value(subset, map_sum_byProject)
                    #如果mapping表中receipt date匹上的project code的subset值的汇总和银行匹配上
                    if (sum(subset_value_map) - bank_value <= 0.03) & (sum(subset_value_map) - bank_value >= -0.03) :
                        glIndex_mappedToInd = []
                        print(subset_value_map)
                        #对加总值匹上的project code进行循环
                        for project_id in subset:
                            print(project_id)
                            df_map_grouped = df_map.groupby(['Notification Email', 'Project ID'])
                            for filter_condition, df in df_map_grouped:
                                if filter_condition[1] == project_id:
                                    map_clear_date = df.iloc[0]['Notification Email']
                                    #获得某一入账时间的project code对应的mapping表总和，该值与GL的子集进行比对
                                    sum_value_map = df['AR in Office Currency'].sum()
                                    print('sum_value_map', sum_value_map)
                                    #筛出还未mapping过的glData
                                    glData_commercial_filtered = glData_commercial.loc[glData_commercial.index.difference(mapped_glIndex_commercial)]
                                    #对还未mapping上的glData用入账时间和project code进行初步筛选
                                    glData_commercial_filtered = glData_commercial_filtered[(glData_commercial_filtered['JH Created Date'] < map_clear_date+dt.timedelta(days=8)) & (glData_commercial_filtered['JH Created Date'] > map_clear_date-dt.timedelta(days=8)) & (glData_commercial_filtered['Project Id'] == f'{project_id}')]
                                    value_list_gl = glData_commercial_filtered['Amount Func Cur'].to_dict()
                                    print('value_list_gl', value_list_gl)
                                    index_list_gl = list(value_list_gl.keys())
                                    print('index_list_gl', index_list_gl)
                                    subsets_index_gl = get_sub_set(index_list_gl)
                                    print(subsets_index_gl)
                                    for subset_index_gl in subsets_index_gl:
                                        subset_value_gl = key_to_value(subset_index_gl, value_list_gl)
                                        #若筛出的glData的某个子集之和等于某一入账时间的project code对应的mapping表总和
                                        if sum(subset_value_gl) == sum_value_map:
                                            print(f'{subset_index_gl}', 'mapped')
                                            # id_number = id_number+1
                                            for index in subset_index_gl:
                                                if (index in glIndex_mappedToInd):
                                                    print(f'{index} previously mapped')
                                                    print('mapped_glIndex_commercial', mapped_glIndex_commercial)
                                                    break
                                                else:
                                                    print('recorded index:', index)
                                                    # bankData.loc[ind, 'notes'] = f'commercial netoff {id_number}'
                                                    # glData.loc[index, 'notes'] = f'commercial netoff {id_number}'
                                                    # bank_charges = bank_charges + sum(subset_value_map) - bank_value
                                                    # mapped_glIndex_commercial.append(index)
                                                    glIndex_mappedToInd.append(index)
                                                    break
                                                    # mapped_bankIndex_commercial.append(ind)
                        print('glIndex_mappedToInd', glIndex_mappedToInd)
                        glData_sum_mappedToInd = glData_commercial.loc[glIndex_mappedToInd]['Amount Func Cur'].sum()
                        check = glData_sum_mappedToInd - bank_value
                        if (check <= 0.03) & (check >= -0.03):
                            if (ind in mapped_bankIndex_commercial) or common_data(glIndex_mappedToInd, mapped_glIndex_commercial):
                                pass
                            else:
                                id_number = id_number + 1
                                print('id_number', id_number)
                                bankData.loc[ind, 'notes'] = f'commercial netoff {id_number}'
                                glData.loc[glIndex_mappedToInd, 'notes'] = f'commercial netoff {id_number}'
                                mapped_bankIndex_commercial.append(ind)
                                mapped_glIndex_commercial = mapped_glIndex_commercial + glIndex_mappedToInd



                        print('glIndex_mappedToInd', glIndex_mappedToInd)
                        print('glData_sum_mappedToInd', glData_sum_mappedToInd)
                        print(check)


    bankData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\bank.xlsx')
    glData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\gl.xlsx')

    # #获取第一轮mapping之后剩余部分的glData和bankData
    # bankIndex_commercial_left = list(set(list(bankData_commercial.index)).difference(set(mapped_bankIndex_commercial)))
    # glIndex_commercial_left = list(set(list(glData_commercial.index)).difference(set(mapped_glIndex_commercial)))
    #
    #
    # glData_commercial_left = glData_commercial.loc[glIndex_commercial_left, :]
    # bankData_commercial_left = bankData_commercial.loc[bankIndex_commercial_left, :]
    # bankData_commercial_left['Narrative'] = bankData_commercial_left['Narrative'].map(lambda x: ''.join(line.strip() for line in x.splitlines()))
    # glData_commercial_left = glData_commercial_left.groupby('Vendor Name')
    #
    # mapped_glIndex_commercial_2 = []
    # mapped_bankIndex_commercial_2 = []
    #
    # for client, df_left in glData_commercial_left:
    #     sum_gl = df_left['Amount Func Cur'].sum()
    #     bankAccountName = map_commercial.loc[map_commercial['Client Name'] == f'{client}', 'Client Name in Chinese']
    #     if len(set(bankAccountName.to_list())):
    #         for name in set(bankAccountName.to_list()):
    #             pro_name = name.strip()
    #             tf = bankData_commercial_left["Narrative"].str.contains(f'{pro_name}', regex = False)
    #             bank_value_list = bankData_commercial_left.loc[tf, 'Credit/Debit amount'].to_dict()
    #             subsets_value_bk = get_sub_set(bank_value_list.values())
    #             for subset_bk in subsets_value_bk:
    #                 if ((sum(subset_bk) - sum_gl) <= 0.03) & ((sum(subset_bk) - sum_gl) >= -0.03):
    #                     for item_bk in subset_bk:
    #                         index_bk = list(filter(lambda x: bank_value_list[x] == item_bk, bank_value_list))
    #                         if (j in mapped_bankIndex_commercial_2 for j in index_bk) or (a in mapped_glIndex_commercial_2 for a in list(df_left.index)):
    #                             pass
    #                         else:
    #                             id_number = id_number + 1
    #                             bankData.loc[index_bk, 'notes'] = f'commercial netoff {id_number}'
    #                             glData.loc[df_left.index, 'notes'] = f'commercial netoff {id_number}'
    #                             bank_charges = bank_charges + sum_gl - sum(subset_bk)
    #                             print(bank_charges)
    #                             mapped_glIndex_commercial_2 = mapped_bankIndex_commercial_2 + list(df_left.index)
    #                             mapped_bankIndex_commercial_2 = mapped_bankIndex_commercial_2 + index_bk


    # bankData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\bank.xlsx')
    # glData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\gl.xlsx')
















    #     tf_map_commercial =  (map_commercial['Actual Receipt  Amount'] == row['Credit/Debit amount']) & (map_commercial['Receipt Dt'] == row['Value date'])
    #     df_test = pd.DataFrame()
    #     df_test['amount'] = map_commercial['Actual Receipt  Amount'] == row['Credit/Debit amount']
    #     df_test['actual receipt dt']=map_commercial['Receipt Dt']
    #     df_test['Value date'] = row['Value date']
    #     excel_log.log(df_test, 'test条件')
    #     map_commercial_filtered = map_commercial.loc[tf_map_commercial, :]
    #     excel_log.log(map_commercial_filtered, 'step1初步筛选后的mapping表')
    #     if map_commercial_filtered.size:
    #         glData_commercial_mapped = pd.DataFrame()
    #         for i in range(len(map_commercial_filtered)):
    #             map_commercial_filtered['Notification Email'] = map_commercial_filtered['Notification Email'].astype('datetime64[ns]')#.dt.date
    #             date = map_commercial_filtered.iloc[i]['Notification Email']
    #             project_id = map_commercial_filtered.iloc[i]['Project ID']
    #             # charges = map_commercial_filtered.iloc[i]['bank expense']
    #             amount = map_commercial_filtered.iloc[i]['AR in Office Currency']
    #             s_date = date+dt.timedelta(days=8)
    #             e_date = date-dt.timedelta(days=8)
    #             date_condition = (glData_commercial['JH Created Date'] <= s_date) & (glData_commercial['JH Created Date'] >= e_date)
    #             project_id_condition = glData_commercial['Project Id'] == project_id
    #             amount_condition = glData_commercial['Amount Func Cur'] == amount
    #             df_glCommercial_mapped = glData_commercial[date_condition & project_id_condition & amount_condition]
    #             glData_commercial_mapped = pd.concat([glData_commercial_mapped, df_glCommercial_mapped])
    #         excel_log.log(glData_commercial_mapped, 'step2gl')
    #         gl_commercial_sum = glData_commercial_mapped['Amount Func Cur'].sum()
    #         bk_commercial_amount = row['Credit/Debit amount']
    #         check = bk_commercial_amount - gl_commercial_sum
    #         print(check)
    #         if check == 0:
    #             id_number = id_number + 1
    #             bankData.loc[ind, 'notes'] = f'commercial netoff {id_number}'
    #             glData.loc[glData_commercial_mapped.index, 'notes'] = f'commercial netoff {id_number}'
    #         if (check != 0) & (check <= 1000):
    #             df_charges_mapped = glData_commercial[(glData_commercial['Vendor Name'] == '') & (glData_commercial['Amount Func Cur'] == check)]
    #             if df_charges_mapped.size:
    #                 id_number = id_number + 1
    #                 bankData.loc[ind, 'notes'] = f'commercial netoff {id_number}'
    #                 glData.loc[glData_commercial_mapped.index, 'notes'] = f'commercial netoff {id_number}'
    #         else:
    #             continue
    #
    # bankData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\bank.xlsx')
    # glData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\gl.xlsx')









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




