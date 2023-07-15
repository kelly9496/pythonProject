import os
import pandas
import pandas as pd
from datetime import datetime
import datetime as dt
import pdfplumber
import re
from decimal import Decimal

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

def key_to_value(subset_key, dict):
    subset_value = []
    for key in subset_key:
        value = dict[key]
        subset_value.append(value)
    return subset_value

def extract_reimPayment_info(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text = text + page.extract_text()
    list_staffName = []
    list_staffNo = []
    list_amount = []
    list_date = []
    totalAmount = []
    entityName = []
    number_PIR = []
    for line in text.splitlines():
        # 抓取员工人名
        re_staffName = re.compile(r'Trading Partner (.+) Processing')
        match_staffName = re_staffName.search(line)
        if match_staffName:
            staffName = match_staffName.group(1)
            list_staffName.append(staffName)
        # 抓取员工工号
        re_staffNo = re.compile(r'^Number (\d+)$')
        match_staffNo = re_staffNo.search(line)
        if match_staffNo:
            staffNo = match_staffNo.group(1)
            list_staffNo.append(staffNo)
        # 抓取付款金额
        re_amount = re.compile(r'Payment Amount (.+) Supplier Number')
        match_amount = re_amount.search(line)
        if match_amount:
            amount = match_amount.group(1)
            amount = amount.replace(',', '')
            amount = float(f'{amount}')
            list_amount.append(amount)
        # 抓取付款日期
        re_date = re.compile(r'Payment Date (.+) Payment Method')
        match_date = re_date.search(line)
        if match_date:
            date = match_date.group(1)
            list_date.append(date)
        # 抓取batch付款总金额
        re_sum = re.compile(r'Total (.*[0-9]+[.][0-9]{2})$')
        match_sum = re_sum.search(line)
        if match_sum:
            total = match_sum.group(1)
            totalAmount = total
        # 抓取entity
        re_entity = re.compile(r'Legal Entity (.+)$')
        match_entity = re_entity.search(line)
        if match_entity:
            entity = match_entity.group(1)
            entityName = entity
        # 抓取Payment Instruction Reference No.
        re_PIR = re.compile(r'Payment Instruction Reference (\d+)')
        match_PIR = re_PIR.search(line)
        if match_PIR:
            number_PIR = match_PIR.group(1)
    df_reimPayment = pd.DataFrame()
    if len(list_staffName):
        df_reimPayment['Staff Name'] = pd.DataFrame(list_staffName)
    if len(list_staffNo):
        df_reimPayment['Staff No'] = pd.DataFrame(list_staffNo)
    if len(list_amount):
        df_reimPayment['Payment Amount'] = pd.DataFrame(list_amount)
    if len(list_date):
        df_reimPayment['Payment Date'] = pd.DataFrame(list_date)
        df_reimPayment['Payment Date'] = pd.to_datetime(df_reimPayment['Payment Date'])
        df_reimPayment['Month'] = df_reimPayment['Payment Date'].dt.month
        month_conversion = {1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUN', 7: 'JUL', 8: 'AUG', 9: 'SEP',
                            10: 'OCT', 11: 'NOV', 12: 'DEC'}
        df_reimPayment['Month'] = df_reimPayment['Month'].map(lambda x: month_conversion[x])
    if len(totalAmount):
        df_reimPayment['Batch Amount'] = totalAmount
    if len(entityName):
        df_reimPayment['Entity'] = entityName
    if len(number_PIR):
        df_reimPayment['PIR Number'] = number_PIR
    return df_reimPayment

# directory_BS = input("Please enter the folder directory of all the BS statements:")
# directory_GL = input("Please enter the folder directory of all the GL files:")
# directory_AP_Vendor = input("Please enter the file link of the AP_Vendor Mapping:")
# directory_AP_Employee = input("Please enter the file link of the AP_Employee Mapping:")
# directory_Commercial = input("Please enter the file link of the Commercial Mapping")


path_folder_BS = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Bank Statement'
path_folder_GL = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\GL'
path_folder_reimRegister = r'C:\Users\he kelly\Desktop\Alteryx & Python\reimbursement\New folder\SH TS 2023'
directory_AP_Vendor = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\AP Mapping.xlsx'
directory_AP_Employee = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\Employee mapping.xlsx'
directory_Commercial = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\Cash receipt 2023.xlsx'

now = str(datetime.now()).split('.')[0]

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
df_bank['Narrative'] = df_bank['Narrative'].map(lambda x: ''.join(line.strip() for line in x.splitlines()))



#获取所有GL信息
files_GL = os.listdir(rf'{path_folder_GL}')
df_GL = pd.DataFrame()
for file_GL in files_GL:
    file_path_GL = os.path.join(path_folder_GL, file_GL)
    df_file_GL = pd.read_excel(file_path_GL, header=1).reset_index()
    df_GL = pd.concat([df_GL, df_file_GL])


# #获取vendor mapping
# map_vendor = pd.read_excel(directory_AP_Vendor, header=0)
# map_vendor['Vendor Name'] = map_vendor['Vendor Name'].map(lambda x: x.upper())

#获取employee mapping
map_employee = pd.read_excel(directory_AP_Employee, header=1)
files_reim = os.listdir(rf'{path_folder_reimRegister}')
file_paths_reimRegister = []
for root, dirs, files in os.walk(path_folder_reimRegister):
    for file in files:
        file_path_reim = os.path.join(root, file)
        file_format = os.path.splitext(file_path_reim)[1]
        if file_format == '.pdf':
            file_paths_reimRegister.append(file_path_reim)
df_reimPay = pd.DataFrame()
for i in file_paths_reimRegister:
    df = extract_reimPayment_info(i)
    df_reimPay = pd.concat([df_reimPay, df])
# df_reimPay = pd.read_excel(r'C:\Users\he kelly\Desktop\Alteryx & Python\reimbursement\reimbursement summary.xlsx', sheet_name='Sheet1')
accountNo_to_entity = {'626-055784-001': 'Beijing LE', '622-512317-001': 'BCG Shenzhen LE', '088-169370-011': 'China PRC LE'}

# #读取Commercial mapping, 创建mapping dictionary
# map_commercial = pd.read_excel(directory_Commercial, header=0)
# map_commercial['Actual Receipt  Amount'].fillna(method='ffill', axis=0, inplace=True)
# map_commercial['Receipt Dt'] = map_commercial['Receipt Dt'].astype('datetime64[ns]')
# map_commercial['bank expense'] = map_commercial['bank expense'].astype('float')
# map_commercial['Client Name'] = map_commercial['Client Name'].dropna().map(lambda x: x.upper())
# tb_location = {'088-169370-011': 'PRC', '626-055784-001': 'Beijing', '622-512317-001': 'Shenzhen'}



#PRC Section
# bank_mapping_PRC = {'626-055784-001': '101001', '622-512317-001': '101135', '088-169370-011': '101244'}
bank_mapping_PRC = {'088-169370-011': '101244'}

for account_number, account_cd in bank_mapping_PRC.items():

    #获取当前bank account的bank和gl数据
    bankData = df_bank[df_bank['Account number']==f'{account_number}']
    glData = df_GL[df_GL['Account Cd']==int(account_cd)]
    print(glData)


    glData_reimbursement = pd.DataFrame()
    staff_invoice_indication = ['HLYERR', 'TB', 'RVCR', 'CM', 'CR']
    for item in staff_invoice_indication:
        df_staff = glData[glData['Invoice Number'].str.contains(f'{item}', regex=False, case=False, na=False)]
        print(item, df_staff)
        glData_reimbursement = pd.concat([glData_reimbursement, df_staff])
    glData_reimbursement['Staff Name'] = glData_reimbursement['Vendor Name'].map(lambda x: x.split('      ')[0])
    print(glData_reimbursement['Staff Name'])



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

    #在bkData中filter出TS batch
    bankData_SBID = bankData[bankData['Bank reference'].str.contains('SBID', regex=False, case=False, na=False)]
    bankData_TSBatch = bankData_SBID
    bankData_TSBatch['Keyword'] = bankData_TSBatch['Narrative'].map(lambda x: x.split('/')[2])
    keyword_nonTS = ['COL', 'Intern', 'PTA', 'Payroll', 'Bonus', 'Cash advance']
    for keyword in keyword_nonTS:
        bankData_keyword = bankData_TSBatch[bankData_TSBatch['Keyword'].str.contains(f'{keyword}', regex=False, case=False, na=False)]
        bankData_TSBatch = bankData_TSBatch.loc[bankData_TSBatch.index.difference(bankData_keyword.index)]


    #开始reimbursement mapping, 设定初始值
    id_number_reim = 0
    mapped_glIndex_reim = []
    mapped_bankIndex_reim = []
    valueMapped_bankIndex_reim = []
    entity = accountNo_to_entity[f'{account_number}']
    df_reimPay_filtered = df_reimPay[df_reimPay['Entity'].str.contains(f'{entity}')]
    for month, df_reimPay_perM in df_reimPay_filtered.groupby('Month'):
        # 可以增加input -> covered month periods
        # if month in ['MAR', 'APR', 'MAY', 'JUN', 'JUL']:
        #     continue
        print(month)
        list_staffName = set(df_reimPay_perM['Staff Name'].to_list())
        gl_mapped = []
        gl_mapped_index = []
        for staff in list_staffName:
            # if staff != 'FRANK ZHU':
            #     continue
            print(staff)
            sum_pay = df_reimPay_perM.loc[df_reimPay_perM['Staff Name'] == f'{staff}', 'Payment Amount'].sum()
            print(sum_pay)
            gl_perStaff_mapped = False
            #将reimbursement payment info与gl进行比对
            #按月份和员工名筛出GL里的金额
            glValue_list_reim = glData_reimbursement.loc[(glData_reimbursement['Staff Name'] == f'{staff}') & (glData_reimbursement['JE Headers Description'].str.contains(f'{month}', case=False)), 'Amount Func Cur'].to_dict()
            sum_gl_reim = sum(glValue_list_reim.values())
            print(sum_gl_reim)
            if sum_pay + sum_gl_reim < 0.02 and sum_pay + sum_gl_reim > -0.02:
                if common_data(mapped_glIndex_reim, glValue_list_reim.keys()):
                    pass
                else:
                    gl_perStaff_mapped = True
                    gl_staffMapped_index = list(glValue_list_reim.keys())
                    gl_mapped_index = gl_mapped_index + gl_staffMapped_index
                    print('mapped index', gl_staffMapped_index)
            else:
                print('subset')
                subsets_glIndex_reim = get_sub_set(glValue_list_reim)
                for subset_glIndex_reim in subsets_glIndex_reim:
                    subset_glValue_reim = key_to_value(subset_glIndex_reim, glValue_list_reim)
                    if sum_pay + sum(subset_glValue_reim) < 0.02 and sum_pay + sum_gl_reim > -0.02:
                        if common_data(mapped_glIndex_reim, subset_glIndex_reim):
                            pass
                        else:
                            gl_perStaff_mapped = True
                            gl_staffMapped_index = list(subset_glIndex_reim)
                            gl_mapped_index = gl_mapped_index + gl_staffMapped_index
                            print('subset mapped', gl_mapped_index)
                            break
            gl_mapped.append(gl_perStaff_mapped)
        #将reimbursement payment info与bk进行比对
        list_PIRnumber = set(df_reimPay_perM['PIR Number'].to_list())
        print(list_PIRnumber)
        bk_mapped = []
        bk_mapped_index = []
        valueMappedIndex_to_PIR = {}
        for number_PIR in list_PIRnumber:
            # 优化点：把paymentDate换成Payment Instruction Reference
            print(number_PIR)
            sum_pay_perPIR = df_reimPay_perM.loc[df_reimPay_perM['PIR Number'] == f'{number_PIR}', 'Payment Amount'].sum()
            print(sum_pay_perPIR)
            bkValue_list_reim = bankData.loc[bankData['Credit/Debit amount'] == round(-sum_pay_perPIR, 2), 'Credit/Debit amount'].to_dict()
            print(bkValue_list_reim)
            bk_perPIR_mapped = False
            if len(bkValue_list_reim) == 1:
                if common_data(mapped_bankIndex_reim, list(bkValue_list_reim.keys())):
                    pass
                else:
                    bk_perPIR_mapped = True
                    bk_mapped_index = bk_mapped_index + list(bkValue_list_reim.keys())
                    print('bk_mapped_index', bk_mapped_index)
            if len(bkValue_list_reim) >= 2:
                index_in_bkTSBatch = set(bkValue_list_reim.keys()).intersection(set(bankData_TSBatch.index.tolist()))
                print(index_in_bkTSBatch)
                if len(index_in_bkTSBatch) == 1:
                    bk_perPIR_mapped = True
                    bk_mapped_index = bk_mapped_index + index_in_bkTSBatch
                    print('bk_mapped_index-choose from 2 index', bk_mapped_index)
                else:
                    bk_perPIR_mapped = True
                    valueMappedIndex_to_PIR = dict.fromkeys(list(bkValue_list_reim.keys()), f'{number_PIR}')
                    # bk_valueMapped_index = bk_valueMapped_index + list(bkValue_list_reim.keys())
                    print('bk_valueMapped_index', valueMappedIndex_to_PIR)
            bk_mapped.append(bk_perPIR_mapped)
        print(valueMappedIndex_to_PIR)
        print('bk Condition:', False not in bk_mapped, bk_mapped)
        print('gl Condition:', False not in gl_mapped, gl_mapped)
        if (False not in bk_mapped) and (False not in gl_mapped):
            id_number_reim += 1
            bankData.loc[bk_mapped_index, 'notes'] = f'reimbursement netoff {now} {id_number_reim}'
            for key in valueMappedIndex_to_PIR.keys():
                bankData.loc[key, 'notes'] = f'reimbursement netoff {now} {id_number_reim} - value map {valueMappedIndex_to_PIR[key]}'
            # bankData.loc[list(valueMappedIndex_to_PIR.keys()), 'notes'] = f'reimbursement netoff {now} {id_number_reim} - value map {}'
            glData.loc[gl_mapped_index, 'notes'] = f'reimbursement netoff {now} {id_number_reim}'
            mapped_bankIndex_reim = mapped_bankIndex_reim + bk_mapped_index
            mapped_glIndex_reim = mapped_glIndex_reim + gl_mapped_index

    now_for_folder = now.replace(':', ' ')
    os.makedirs(rf'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\{now_for_folder}\PRC')
    bankData.to_excel(fr'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\{now_for_folder}\PRC\bank_{account_number}.xlsx')
    glData.to_excel(fr'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\{now_for_folder}\PRC\gl_{account_cd}.xlsx')


















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




