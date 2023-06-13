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
    #处理commercial mapping表，筛出本entity RPApo账的部分，并按project ID分类
    map_commercial_RPA = map_commercial[map_commercial['location'].str.contains(f'{location}') & (map_commercial['Currency'] == 'CNY') & (map_commercial['Notification Email'] != '-')]
    excel_log.log(map_commercial_RPA, 'map_commercial_PRC')
    map_commercial_RPA = map_commercial_RPA.groupby(['Receipt Dt'])


    id_number = 0
    mapped_index_commercial = []
    for ind, row in bankData_commercial.iterrows():
        bank_value = row['Credit/Debit amount']
        bank_receipt_date = row['Value date']
        for map_receipt_date, df_map in map_commercial_RPA:
            if map_receipt_date == bank_receipt_date:
                map_sum_byProject = df_map.groupby('Project ID').sum('AR in Office Currency')['AR in Office Currency'].to_dict()
                value_list_map = list(map_sum_byProject.values())
                subsets_value_map = get_sub_set(value_list_map)
                for subset in subsets_value_map:
                    #如果subset值的汇总和银行匹配上
                    if sum(subset) == bank_value:
                        for value in subset:
                            project_id = list(filter(lambda k: map_sum_byProject[k] == value, map_sum_byProject))
                            df_map_grouped = df_map.groupby(['Notification Email', 'Project ID'])
                            for filter_condition, df in df_map_grouped:
                                for code in project_id:
                                    if code in filter_condition:
                                        map_clear_date = df.iloc[0]['Notification Email']
                                        #获得与GL进行子集比对的sum_value
                                        sum_value_map = df['AR in Office Currency'].sum()
                                        glData_commercial_filtered = glData_commercial[(glData_commercial['JH Created Date'] < map_clear_date+dt.timedelta(days=8)) & (glData_commercial['JH Created Date'] > map_clear_date-dt.timedelta(days=8)) & (glData_commercial['Project Id'] == f'{code}')]
                                        value_list_gl = glData_commercial_filtered['Amount Func Cur'].to_dict()
                                        subsets_value_gl = get_sub_set(value_list_gl.values())
                                        for subset_gl in subsets_value_gl:
                                            if sum(subset_gl) == sum_value_map:
                                                for item_gl in subset_gl:
                                                    index = list(filter(lambda x: value_list_gl[x] == item_gl, value_list_gl))
                                                    id_number = id_number+1
                                                    bankData.loc[ind, 'notes'] = f'commercial netoff {id_number}'
                                                    glData.loc[index, 'notes'] = f'commercial netoff {id_number}'

    bankData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\bank.xlsx')
    glData.to_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\gl.xlsx')