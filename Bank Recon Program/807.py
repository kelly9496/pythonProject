import os
import pandas as pd
from datetime import datetime
import datetime as dt
import pdfplumber
import re

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
    print('nums', nums)
    # if len(nums) >=14:
    #     nums = list(nums)[:14]
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

def repayment_mapping(bankData_filtered, bankData, order):

    # AP退款重付
    id_number_repayment = 0
    bankIndex_repayment_netoff = []
    # 找出含有RJCT的行
    bankData_RJCT = bankData_filtered[bankData_filtered['Narrative'].str.contains('RJCT', regex=False, na=False)]
    bankData_RJCT['bankAccountName'] = bankData_RJCT['Narrative'].map(lambda x: x.split('/')[2])
    # 删掉有return字眼的行
    bankData_RJCT.drop(bankData_RJCT[bankData_RJCT['bankAccountName'].str.contains('RETURN')].index, inplace=True)
    # excel_log.log(bankData_RJCT, 'return')
    for ind_RJCT, row_RJCT in bankData_RJCT.iterrows():
        condition_bkAccountName = bankData_filtered['Narrative'].str.contains(f'{row_RJCT["bankAccountName"]}')
        if order == 'first':
            condition_date = bankData_filtered['Value date'] > row_RJCT['Value date']
        if order == 'last':
            condition_date = bankData_filtered['Value date'] >= row_RJCT['Value date']
        condition_amount = bankData_filtered['Credit/Debit amount'] == -row_RJCT['Credit/Debit amount']
        bankData_repayment = bankData_filtered[condition_bkAccountName & condition_date & condition_amount]
        # excel_log.log(bankData_repayment, 'repayment')
        if len(bankData_repayment) == 1:
            # 将indexint64转为int
            ind_repayment = bankData_repayment.index.values[0]
            if ind_RJCT in bankIndex_repayment_netoff or ind_repayment in bankIndex_repayment_netoff:
                pass
            else:
                id_number_repayment = id_number_repayment + 1
                bankData.loc[ind_RJCT, 'notes'] = f'repayment netoff {now} {order} {id_number_repayment}'
                bankData.loc[ind_repayment, 'notes'] = f'repayment netoff {now} {order} {id_number_repayment}'
                bankIndex_repayment_netoff.append(ind_RJCT)
                bankIndex_repayment_netoff.append(ind_repayment)

        if len(bankData_repayment) > 1:
            ind_repayment = bankData_repayment.index.values
            if bankData_repayment.loc[ind_repayment[0], 'Value date'] >= bankData_repayment.loc[
                ind_repayment[1], 'Value date']:
                ind_repayment_selected = ind_repayment[-1]
            else:
                ind_repayment_selected = ind_repayment[0]
            if ind_RJCT in bankIndex_repayment_netoff or ind_repayment_selected in bankIndex_repayment_netoff:
                pass
            else:
                id_number_repayment = id_number_repayment + 1
                bankData.loc[ind_RJCT, 'notes'] = f'repayment netoff {now} {order} {id_number_repayment}'
                bankData.loc[ind_repayment_selected, 'notes'] = f'repayment netoff {now} {order} {id_number_repayment}'
                bankIndex_repayment_netoff.append(ind_RJCT)
                bankIndex_repayment_netoff.append(ind_repayment_selected)

    return bankIndex_repayment_netoff

def commercial_mapping(bankData_commercial, bankData, glData_commercial, glData, map_commercial, account_cd):
    # location = tb_location[bankData_commercial.iloc[1]['Account number']]
    # 处理commercial mapping表，筛出本entity RPApo账的部分，并按project ID分类
    map_commercial_RPA = map_commercial[
        map_commercial['location'].str.contains(f'{location}') & (map_commercial['Currency'] == 'CNY') & (
                    map_commercial['Notification Email'] != '-')]
    # excel_log.log(map_commercial_RPA, 'map_commercial_PRC')
    map_commercial_RPA = map_commercial_RPA.groupby(['Receipt Dt', 'Client Name'])

    # 设定初始值
    id_number_commercial = 0
    mapped_index_commercial = []
    mapped_glIndex_commercial_1 = []
    mapped_bankIndex_commercial_1 = []

    # 第一轮commercial mapping
    for ind, row in bankData_commercial.iterrows():
        bank_value = row['Credit/Debit amount']
        bank_receipt_date = row['Value date']
        for group_condition, df_map in map_commercial_RPA:
            if group_condition[0] == bank_receipt_date:
                map_sum_byProject = df_map.groupby('Project ID').sum('AR in Office Currency')[
                    'AR in Office Currency'].to_dict()
                project_code_list = list(map_sum_byProject.keys())
                project_code_subsets = get_sub_set(project_code_list)
                for subset in project_code_subsets:
                    subset_value_map = key_to_value(subset, map_sum_byProject)
                    mapped_commercial1 = False
                    if account_cd == '101245':
                        list_bankCharge = [0, 10, 20, 25, 30, 35, 40, 50, 60, 70, 80]
                        if sum(subset_value_map) - bank_value in list_bankCharge:
                            mapped_commercial1 = True
                    else:
                        if (sum(subset_value_map) - bank_value <= 0.03) & (sum(subset_value_map) - bank_value >= -0.03):
                            mapped_commercial1 = True
                    # 如果mapping表中receipt date匹上的project code的subset值的汇总和银行匹配上
                    if mapped_commercial1:
                        glIndex_mappedToInd = []
                        # 对加总值匹上的project code进行循环
                        for project_id in subset:
                            df_map_grouped = df_map.groupby(['Notification Email', 'Project ID'])
                            for filter_condition, df in df_map_grouped:
                                if filter_condition[1] == project_id:
                                    map_clear_date = df.iloc[0]['Notification Email']
                                    # 获得某一入账时间的project code对应的mapping表总和，该值与GL的子集进行比对
                                    sum_value_map = df['AR in Office Currency'].sum()
                                    # print('sum_value_map', sum_value_map)
                                    # 筛出还未mapping过的glData
                                    glData_commercial_filtered = glData_commercial.loc[
                                        glData_commercial.index.difference(mapped_glIndex_commercial_1)]
                                    # 对还未mapping上的glData用入账时间和project code进行初步筛选
                                    glData_commercial_filtered = glData_commercial_filtered[(glData_commercial_filtered[
                                                                                                 'JH Created Date'] < map_clear_date + dt.timedelta(
                                        days=8)) & (glData_commercial_filtered[
                                                        'JH Created Date'] > map_clear_date - dt.timedelta(days=8)) & (
                                                                                                    glData_commercial_filtered[
                                                                                                        'Project Id'] == f'{project_id}')]
                                    value_list_gl = glData_commercial_filtered['Amount Func Cur'].to_dict()
                                    # print('value_list_gl', value_list_gl)
                                    index_list_gl = list(value_list_gl.keys())
                                    # print('index_list_gl', index_list_gl)
                                    subsets_index_gl = get_sub_set(index_list_gl)
                                    # print(subsets_index_gl)
                                    for subset_index_gl in subsets_index_gl:
                                        subset_value_gl = key_to_value(subset_index_gl, value_list_gl)
                                        # 若筛出的glData的某个子集之和等于某一入账时间的project code对应的mapping表总和
                                        if sum(subset_value_gl) == sum_value_map:
                                            # print(f'{subset_index_gl}', 'mapped')
                                            # id_number = id_number+1
                                            for index in subset_index_gl:
                                                if (index in glIndex_mappedToInd):
                                                    # print(f'{index} previously mapped')
                                                    # print('mapped_glIndex_commercial', mapped_glIndex_commercial_1)
                                                    break
                                                else:
                                                    # print('recorded index:', index)
                                                    glIndex_mappedToInd.append(index)
                                                    break

                        # print('glIndex_mappedToInd', glIndex_mappedToInd)
                        glData_sum_mappedToInd = glData_commercial.loc[glIndex_mappedToInd]['Amount Func Cur'].sum()
                        check = glData_sum_mappedToInd - bank_value
                        check_successful = False
                        if account_cd == '101245':
                            list_bankCharge = [0, 10, 20, 25, 30, 35, 40, 50, 60, 70, 80]
                            if check in list_bankCharge:
                                check_successful = True
                        else:
                            if (check <= 0.03) & (check >= -0.03):
                                check_successful = True
                        if check_successful:
                            if (ind in mapped_bankIndex_commercial_1) or common_data(glIndex_mappedToInd,
                                                                                     mapped_glIndex_commercial_1):
                                pass
                            else:
                                id_number_commercial = id_number_commercial + 1
                                # print('id_number', id_number_commercial)
                                bankData.loc[ind, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                                glData.loc[
                                    glIndex_mappedToInd, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                                mapped_bankIndex_commercial_1.append(ind)
                                mapped_glIndex_commercial_1 = mapped_glIndex_commercial_1 + glIndex_mappedToInd

    # #获取第一轮mapping之后剩余部分的glData和bankData
    bankIndex_commercial_left = list(
        set(list(bankData_commercial.index)).difference(set(mapped_bankIndex_commercial_1)))
    glIndex_commercial_left = list(set(list(glData_commercial.index)).difference(set(mapped_glIndex_commercial_1)))

    glData_commercial_left = glData_commercial.loc[glIndex_commercial_left, :]
    bankData_commercial_left = bankData_commercial.loc[bankIndex_commercial_left, :]
    # bankData_commercial_left['Narrative'] = bankData_commercial_left['Narrative'].map(lambda x: ''.join(line.strip() for line in x.splitlines()))
    glData_commercial_left = glData_commercial_left.groupby('Vendor Name')

    mapped_glIndex_commercial_2 = []
    mapped_bankIndex_commercial_2 = []

    # 第二轮commercial mapping
    for client, df_left in glData_commercial_left:
        sum_gl = df_left['Amount Func Cur'].sum()
        bankAccountName_client = map_commercial.loc[
            map_commercial['Client Name'] == f'{client}'.upper(), 'Client Name in Chinese']
        if len(set(bankAccountName_client.to_list())):
            for name in set(bankAccountName_client.to_list()):
                pro_name = name.strip()
                tf = bankData_commercial_left["Narrative"].str.contains(f'{pro_name}', regex=False, case=False)
                bkValue_list_commercial = bankData_commercial_left.loc[tf, 'Credit/Debit amount'].to_dict()
                subsets_bkIndex_commercial = get_sub_set(bkValue_list_commercial.keys())
                for subset_bkIndex_commercial in subsets_bkIndex_commercial:
                    subset_value_bk = key_to_value(subset_bkIndex_commercial, bkValue_list_commercial)
                    mapped_commercial2 = False
                    if account_cd == '101245':
                        list_bankCharge = [0, 10, 20, 25, 30, 35, 40, 50, 60, 70, 80]
                        if (sum(subset_value_bk) - sum_gl) in list_bankCharge:
                            mapped_commercial2 = True
                    else:
                        if ((sum(subset_value_bk) - sum_gl) <= 0.03) & ((sum(subset_value_bk) - sum_gl) >= -0.03):
                            mapped_commercial2 = True
                    if mapped_commercial2:
                        if common_data(subset_bkIndex_commercial, mapped_bankIndex_commercial_2) or (
                        common_data(list(df_left.index), mapped_glIndex_commercial_2)):
                            pass
                        else:
                            id_number_commercial = id_number_commercial + 1
                            bankData.loc[
                                subset_bkIndex_commercial, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                            glData.loc[df_left.index, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                            mapped_glIndex_commercial_2 = mapped_bankIndex_commercial_2 + list(df_left.index)
                            mapped_bankIndex_commercial_2 = mapped_bankIndex_commercial_2 + subset_bkIndex_commercial

    mapped_glIndex_commercial = mapped_glIndex_commercial_1 + mapped_glIndex_commercial_2
    mapped_bankIndex_commercial = mapped_bankIndex_commercial_1 + mapped_bankIndex_commercial_2

    return mapped_bankIndex_commercial, mapped_glIndex_commercial


def AP_mapping(bankData_AP, bankData, glData_vendor, glData, map_vendor_byEntity, account_cd, list_bankCharge, nameNum, vendorStaff):
    #修改点 编号不连续
    id_number_AP = 0
    mapped_glIndex_vendor = []
    mapped_bankIndex_vendor = []

    # AP Mapping 1 GL总和匹BK子集
    for vendor, df in glData_vendor.groupby('Vendor Name'):
        # if vendor != 'Atkins Insights Co.,Ltd':
        #     continue
        # print('Start: 1 GL总和匹BK子集')
        glSum_vendor = df['Amount Func Cur'].sum()
        bankAccountNum_vendor = map_vendor_byEntity.loc[map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), f'Bank Account {nameNum}']
        print('bankAccountNum_vendor', bankAccountNum_vendor)
        # 修改点：当一个vendor name匹配一个银行账号或匹配多个银行账号的情况可以合并成一个条件if len(set(bankAccountNum_vendor.to_list())) 来写
        # 当一个vendor name匹配一个银行账号时
        # if len(set(bankAccountNum_vendor.to_list())) == 1:
        #     for num in set(bankAccountNum_vendor.to_list()):
        #         pro_num = str(num).strip()
        #         # tf = bankData_AP["Narrative"].str.contains(f'{pro_num}', regex=False, case=False)
        #         bkValue_list_AP = bankData_AP.loc[bankData_AP["Narrative"].str.contains(f'{pro_num}', regex=False, case=False), 'Credit/Debit amount'].to_dict()
        #         print('bkValue_list_AP', bkValue_list_AP)
        #         subsets_bkIndex_vendor = get_sub_set(bkValue_list_AP.keys())
        #         for subset_bkIndex_vendor in subsets_bkIndex_vendor:
        #             subset_bkValue_vendor = key_to_value(subset_bkIndex_vendor, bkValue_list_AP)
        #             mapped_first = False
        #             if account_cd == '101245':
        #                 if (sum(subset_bkValue_vendor) - glSum_vendor) in list_bankCharge:
        #                     mapped_first = True
        #             else:
        #                 if ((sum(subset_bkValue_vendor) - glSum_vendor) <= 0.03) & ((sum(subset_bkValue_vendor) - glSum_vendor) >= -0.03):
        #                     mapped_first = True
        #             if mapped_first:
        #                 if common_data(subset_bkValue_vendor, mapped_bankIndex_vendor) or common_data(list(df.index), mapped_glIndex_vendor):
        #                     print('duplicate index')
        #                     pass
        #                 else:
        #                     print('mapped')
        #                     id_number_AP = id_number_AP + 1
        #                     bankData.loc[subset_bkIndex_vendor, 'notes'] = f'AP netoff {now} {id_number_AP}'
        #                     glData.loc[df.index, 'notes'] = f'AP netoff {now} {id_number_AP}'
        #                     mapped_glIndex_vendor = mapped_glIndex_vendor + list(df.index)
        #                     mapped_bankIndex_vendor = mapped_bankIndex_vendor + subset_bkIndex_vendor
        #                     break

        # 当一个vendor name匹配多个银行账号时
        if len(set(bankAccountNum_vendor.to_list())):
            bkIndex_vendor = []
            dic_bkValue_AP = {}
            for num in set(bankAccountNum_vendor.to_list()):
                pro_num = str(num).strip()
                bkValue_list_AP = bankData_AP.loc[bankData_AP["Narrative"].str.contains(f'{pro_num}', regex=False, case=False), 'Credit/Debit amount'].to_dict()
                bkIndex_vendor = bkIndex_vendor + list(bkValue_list_AP.keys())
                dic_bkValue_AP.update(bkValue_list_AP)
            print(bkIndex_vendor)
            subsets_bkIndex_vendor_2 = get_sub_set(bkIndex_vendor)
            for subset_bkIndex_vendor_2 in subsets_bkIndex_vendor_2:
                subset_bkValue_vendor_2 = key_to_value(subset_bkIndex_vendor_2, dic_bkValue_AP)
                mapped_first = False
                if account_cd == '101245':
                    if (glSum_vendor - sum(subset_bkValue_vendor_2)) in list_bankCharge:
                        mapped_first = True
                else:
                    if ((sum(subset_bkValue_vendor_2) - glSum_vendor) <= 0.03) & (
                            (sum(subset_bkValue_vendor_2) - glSum_vendor) >= -0.03):
                        mapped_first = True
                if mapped_first:
                    if common_data(subset_bkValue_vendor_2, mapped_bankIndex_vendor) or common_data(list(df.index), mapped_glIndex_vendor):
                        pass
                    else:
                        id_number_AP = id_number_AP + 1
                        bankData.loc[subset_bkIndex_vendor_2, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                        glData.loc[df.index, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                        mapped_glIndex_vendor = mapped_glIndex_vendor + list(df.index)
                        mapped_bankIndex_vendor = mapped_bankIndex_vendor + subset_bkIndex_vendor_2
                        break

    # AP gl和bk挖去第一次匹配后的结果
    bankData_AP_left1 = bankData_AP.loc[bankData_AP.index.difference(mapped_bankIndex_vendor)]
    glData_vendor_left1 = glData_vendor.loc[glData_vendor.index.difference(mapped_glIndex_vendor)]

    # AP Mapping 2 bk总和匹gl子集
    for vendor, df in glData_vendor_left1.groupby('Vendor Name'):
        # if vendor != 'Atkins Insights Co.,Ltd':
        #     continue
        print('Start: 2 bk总和匹gl子集')
        bankAccountName_vendor = map_vendor_byEntity.loc[map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), f'Bank Account {nameNum}']
        print('bankAccountName_vendor', vendor, bankAccountName_vendor)
        dic_bkValue_AP_2 = {}
        if len(set(bankAccountName_vendor.to_list())) >= 1:
            for name in set(bankAccountName_vendor.to_list()):
                bkValue_list_AP_2 = bankData_AP_left1.loc[bankData_AP_left1['Narrative'].str.contains(f'{str(name).strip()}', regex=False, case=False, na=False), 'Credit/Debit amount'].to_dict()
                dic_bkValue_AP_2.update(bkValue_list_AP_2)
        bk_sum = sum(dic_bkValue_AP_2.values())
        bk_Index = list(dic_bkValue_AP_2.keys())
        glValue_list_vendor = df['Amount Func Cur'].to_dict()
        subsets_glIndex_vendor = get_sub_set(glValue_list_vendor.keys())
        for subset_glIndex_vendor in subsets_glIndex_vendor:
            subset_glValue_vendor = key_to_value(subset_glIndex_vendor, glValue_list_vendor)
            mapped_second = False
            if account_cd == '101245':
                if sum(subset_glValue_vendor) - bk_sum in list_bankCharge:
                    mapped_second = True
            else:
                if (sum(subset_glValue_vendor) - bk_sum <= 0.03) & (sum(subset_glValue_vendor) - bk_sum >= -0.03):
                    mapped_second = True
            if mapped_second:
                if common_data(subset_glIndex_vendor, mapped_glIndex_vendor) or common_data(bk_Index, mapped_bankIndex_vendor):
                    pass
                else:
                    id_number_AP = id_number_AP + 1
                    bankData.loc[bk_Index, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                    glData.loc[subset_glIndex_vendor, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                    mapped_glIndex_vendor = mapped_glIndex_vendor + list(subset_glIndex_vendor)
                    mapped_bankIndex_vendor = mapped_bankIndex_vendor + bk_Index
                    break

    # AP gl和bk挖去第二次匹配后的结果
    bankData_AP_left2 = bankData_AP.loc[bankData_AP.index.difference(mapped_bankIndex_vendor)]
    glData_vendor_left2 = glData_vendor.loc[glData_vendor.index.difference(mapped_glIndex_vendor)]

    # vendor gl子集和bk子集匹配
    for vendor, df in glData_vendor_left2.groupby('Vendor Name'):
        # if vendor != 'Atkins Insights Co.,Ltd':
        #     continue
        print('Start: 3 gl子集和bk子集匹配')
        bankAccountName_vendor_2 = map_vendor_byEntity.loc[
            map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), f'Bank Account {nameNum}']
        print('bankAccountName_vendor', bankAccountName_vendor_2)
        dic_bkValue_AP_3 = {}
        if len(set(bankAccountName_vendor_2.to_list())) >= 1:
            for name in set(bankAccountName_vendor_2.to_list()):
                bkValue_list_AP_3 = bankData_AP_left2.loc[
                    bankData_AP_left2['Narrative'].str.contains(f'{str(name).strip()}', regex=False, case=False,
                                                                na=False), 'Credit/Debit amount'].to_dict()
                dic_bkValue_AP_3.update(bkValue_list_AP_3)
        subsets_bkIndex_vendor_3 = get_sub_set(dic_bkValue_AP_3.keys())
        glValue_list_vendor_2 = df['Amount Func Cur'].to_dict()
        subsets_glIndex_vendor_2 = get_sub_set(glValue_list_vendor_2.keys())
        for subset_bkIndex_vendor_3 in subsets_bkIndex_vendor_3:
            if common_data(subset_bkIndex_vendor_3, mapped_bankIndex_vendor):
                continue
            subset_bkValue_vendor_3 = key_to_value(subset_bkIndex_vendor_3, dic_bkValue_AP_3)
            for subset_glIndex_vendor_2 in subsets_glIndex_vendor_2:
                if common_data(subset_glIndex_vendor_2, mapped_glIndex_vendor):
                    continue
                subset_glValue_vendor_2 = key_to_value(subset_glIndex_vendor_2, glValue_list_vendor_2)
                mapped_third = False
                if account_cd == '101245':
                    if sum(subset_glValue_vendor_2) - sum(subset_bkValue_vendor_3) in list_bankCharge:
                        mapped_third = True
                else:
                    if (sum(subset_glValue_vendor_2) - sum(subset_bkValue_vendor_3) <= 0.03) & (
                            sum(subset_glValue_vendor_2) - sum(subset_bkValue_vendor_3) >= -0.03):
                        mapped_third = True
                if mapped_third:
                    if common_data(subset_glIndex_vendor_2, mapped_glIndex_vendor) or common_data(
                            subset_bkIndex_vendor_3, mapped_bankIndex_vendor):
                        print("duplicate index")
                        pass
                    else:
                        print("mapped")
                        id_number_AP = id_number_AP + 1
                        bankData.loc[subset_bkIndex_vendor_3, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                        glData.loc[subset_glIndex_vendor_2, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                        mapped_glIndex_vendor = mapped_glIndex_vendor + list(subset_glIndex_vendor_2)
                        mapped_bankIndex_vendor = mapped_bankIndex_vendor + list(subset_bkIndex_vendor_3)

    return mapped_bankIndex_vendor, mapped_glIndex_vendor



# def AP_mapping(bankData_AP, bankData, glData_vendor, glData, map_vendor_byEntity, account_cd, list_bankCharge):
#     #修改点 编号不连续
#     id_number_AP = 0
#     mapped_glIndex_vendor = []
#     mapped_bankIndex_vendor = []
#
#     # AP Mapping 1 GL总和匹BK子集
#     for vendor, df in glData_vendor.groupby('Vendor Name'):
#         # if vendor != 'Atkins Insights Co.,Ltd':
#         #     continue
#         # print('Start: 1 GL总和匹BK子集')
#         glSum_vendor = df['Amount Func Cur'].sum()
#         bankAccountNum_vendor = map_vendor_byEntity.loc[map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), 'Bank Account Num']
#         print('bankAccountNum_vendor', bankAccountNum_vendor)
#         # 修改点：当一个vendor name匹配一个银行账号或匹配多个银行账号的情况可以合并成一个条件if len(set(bankAccountNum_vendor.to_list())) 来写
#         # 当一个vendor name匹配一个银行账号时
#         # if len(set(bankAccountNum_vendor.to_list())) == 1:
#         #     for num in set(bankAccountNum_vendor.to_list()):
#         #         pro_num = str(num).strip()
#         #         # tf = bankData_AP["Narrative"].str.contains(f'{pro_num}', regex=False, case=False)
#         #         bkValue_list_AP = bankData_AP.loc[bankData_AP["Narrative"].str.contains(f'{pro_num}', regex=False, case=False), 'Credit/Debit amount'].to_dict()
#         #         print('bkValue_list_AP', bkValue_list_AP)
#         #         subsets_bkIndex_vendor = get_sub_set(bkValue_list_AP.keys())
#         #         for subset_bkIndex_vendor in subsets_bkIndex_vendor:
#         #             subset_bkValue_vendor = key_to_value(subset_bkIndex_vendor, bkValue_list_AP)
#         #             mapped_first = False
#         #             if account_cd == '101245':
#         #                 if (sum(subset_bkValue_vendor) - glSum_vendor) in list_bankCharge:
#         #                     mapped_first = True
#         #             else:
#         #                 if ((sum(subset_bkValue_vendor) - glSum_vendor) <= 0.03) & ((sum(subset_bkValue_vendor) - glSum_vendor) >= -0.03):
#         #                     mapped_first = True
#         #             if mapped_first:
#         #                 if common_data(subset_bkValue_vendor, mapped_bankIndex_vendor) or common_data(list(df.index), mapped_glIndex_vendor):
#         #                     print('duplicate index')
#         #                     pass
#         #                 else:
#         #                     print('mapped')
#         #                     id_number_AP = id_number_AP + 1
#         #                     bankData.loc[subset_bkIndex_vendor, 'notes'] = f'AP netoff {now} {id_number_AP}'
#         #                     glData.loc[df.index, 'notes'] = f'AP netoff {now} {id_number_AP}'
#         #                     mapped_glIndex_vendor = mapped_glIndex_vendor + list(df.index)
#         #                     mapped_bankIndex_vendor = mapped_bankIndex_vendor + subset_bkIndex_vendor
#         #                     break
#
#         # 当一个vendor name匹配多个银行账号时
#         if len(set(bankAccountNum_vendor.to_list())):
#             bkIndex_vendor = []
#             dic_bkValue_AP = {}
#             for num in set(bankAccountNum_vendor.to_list()):
#                 pro_num = str(num).strip()
#                 bkValue_list_AP = bankData_AP.loc[bankData_AP["Narrative"].str.contains(f'{pro_num}', regex=False, case=False), 'Credit/Debit amount'].to_dict()
#                 bkIndex_vendor = bkIndex_vendor + list(bkValue_list_AP.keys())
#                 dic_bkValue_AP.update(bkValue_list_AP)
#             print(bkIndex_vendor)
#             subsets_bkIndex_vendor_2 = get_sub_set(bkIndex_vendor)
#             for subset_bkIndex_vendor_2 in subsets_bkIndex_vendor_2:
#                 subset_bkValue_vendor_2 = key_to_value(subset_bkIndex_vendor_2, dic_bkValue_AP)
#                 mapped_first = False
#                 if account_cd == '101245':
#                     if (glSum_vendor - sum(subset_bkValue_vendor_2)) in list_bankCharge:
#                         mapped_first = True
#                 else:
#                     if ((sum(subset_bkValue_vendor_2) - glSum_vendor) <= 0.03) & (
#                             (sum(subset_bkValue_vendor_2) - glSum_vendor) >= -0.03):
#                         mapped_first = True
#                 if mapped_first:
#                     if common_data(subset_bkValue_vendor_2, mapped_bankIndex_vendor) or common_data(list(df.index), mapped_glIndex_vendor):
#                         pass
#                     else:
#                         id_number_AP = id_number_AP + 1
#                         bankData.loc[subset_bkIndex_vendor_2, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                         glData.loc[df.index, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                         mapped_glIndex_vendor = mapped_glIndex_vendor + list(df.index)
#                         mapped_bankIndex_vendor = mapped_bankIndex_vendor + subset_bkIndex_vendor_2
#                         break
#
#     # AP gl和bk挖去第一次匹配后的结果
#     bankData_AP_left1 = bankData_AP.loc[bankData_AP.index.difference(mapped_bankIndex_vendor)]
#     glData_vendor_left1 = glData_vendor.loc[glData_vendor.index.difference(mapped_glIndex_vendor)]
#
#     # AP Mapping 2 bk总和匹gl子集
#     for vendor, df in glData_vendor_left1.groupby('Vendor Name'):
#         # if vendor != 'Atkins Insights Co.,Ltd':
#         #     continue
#         print('Start: 2 bk总和匹gl子集')
#         bankAccountName_vendor = map_vendor_byEntity.loc[map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), 'Bank Account Name']
#         print('bankAccountName_vendor', vendor, bankAccountName_vendor)
#         dic_bkValue_AP_2 = {}
#         if len(set(bankAccountName_vendor.to_list())) >= 1:
#             for name in set(bankAccountName_vendor.to_list()):
#                 bkValue_list_AP_2 = bankData_AP_left1.loc[bankData_AP_left1['Narrative'].str.contains(f'{str(name).strip()}', regex=False, case=False, na=False), 'Credit/Debit amount'].to_dict()
#                 dic_bkValue_AP_2.update(bkValue_list_AP_2)
#         bk_sum = sum(dic_bkValue_AP_2.values())
#         bk_Index = list(dic_bkValue_AP_2.keys())
#         glValue_list_vendor = df['Amount Func Cur'].to_dict()
#         subsets_glIndex_vendor = get_sub_set(glValue_list_vendor.keys())
#         for subset_glIndex_vendor in subsets_glIndex_vendor:
#             subset_glValue_vendor = key_to_value(subset_glIndex_vendor, glValue_list_vendor)
#             mapped_second = False
#             if account_cd == '101245':
#                 if sum(subset_glValue_vendor) - bk_sum in list_bankCharge:
#                     mapped_second = True
#             else:
#                 if (sum(subset_glValue_vendor) - bk_sum <= 0.03) & (sum(subset_glValue_vendor) - bk_sum >= -0.03):
#                     mapped_second = True
#             if mapped_second:
#                 if common_data(subset_glIndex_vendor, mapped_glIndex_vendor) or common_data(bk_Index, mapped_bankIndex_vendor):
#                     pass
#                 else:
#                     id_number_AP = id_number_AP + 1
#                     bankData.loc[bk_Index, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                     glData.loc[subset_glIndex_vendor, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                     mapped_glIndex_vendor = mapped_glIndex_vendor + list(subset_glIndex_vendor)
#                     mapped_bankIndex_vendor = mapped_bankIndex_vendor + bk_Index
#                     break
#
#     # AP gl和bk挖去第二次匹配后的结果
#     bankData_AP_left2 = bankData_AP.loc[bankData_AP.index.difference(mapped_bankIndex_vendor)]
#     glData_vendor_left2 = glData_vendor.loc[glData_vendor.index.difference(mapped_glIndex_vendor)]
#
#     # vendor gl子集和bk子集匹配
#     for vendor, df in glData_vendor_left2.groupby('Vendor Name'):
#         # if vendor != 'Atkins Insights Co.,Ltd':
#         #     continue
#         print('Start: 3 gl子集和bk子集匹配')
#         bankAccountName_vendor_2 = map_vendor_byEntity.loc[
#             map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), 'Bank Account Name']
#         print('bankAccountName_vendor', bankAccountName_vendor_2)
#         dic_bkValue_AP_3 = {}
#         if len(set(bankAccountName_vendor_2.to_list())) >= 1:
#             for name in set(bankAccountName_vendor_2.to_list()):
#                 bkValue_list_AP_3 = bankData_AP_left2.loc[
#                     bankData_AP_left2['Narrative'].str.contains(f'{str(name).strip()}', regex=False, case=False,
#                                                                 na=False), 'Credit/Debit amount'].to_dict()
#                 dic_bkValue_AP_3.update(bkValue_list_AP_3)
#         subsets_bkIndex_vendor_3 = get_sub_set(dic_bkValue_AP_3.keys())
#         glValue_list_vendor_2 = df['Amount Func Cur'].to_dict()
#         subsets_glIndex_vendor_2 = get_sub_set(glValue_list_vendor_2.keys())
#         for subset_bkIndex_vendor_3 in subsets_bkIndex_vendor_3:
#             if common_data(subset_bkIndex_vendor_3, mapped_bankIndex_vendor):
#                 continue
#             subset_bkValue_vendor_3 = key_to_value(subset_bkIndex_vendor_3, dic_bkValue_AP_3)
#             for subset_glIndex_vendor_2 in subsets_glIndex_vendor_2:
#                 if common_data(subset_glIndex_vendor_2, mapped_glIndex_vendor):
#                     continue
#                 subset_glValue_vendor_2 = key_to_value(subset_glIndex_vendor_2, glValue_list_vendor_2)
#                 mapped_third = False
#                 if account_cd == '101245':
#                     if sum(subset_glValue_vendor_2) - sum(subset_bkValue_vendor_3) in list_bankCharge:
#                         mapped_third = True
#                 else:
#                     if (sum(subset_glValue_vendor_2) - sum(subset_bkValue_vendor_3) <= 0.03) & (
#                             sum(subset_glValue_vendor_2) - sum(subset_bkValue_vendor_3) >= -0.03):
#                         mapped_third = True
#                 if mapped_third:
#                     if common_data(subset_glIndex_vendor_2, mapped_glIndex_vendor) or common_data(
#                             subset_bkIndex_vendor_3, mapped_bankIndex_vendor):
#                         print("duplicate index")
#                         pass
#                     else:
#                         print("mapped")
#                         id_number_AP = id_number_AP + 1
#                         bankData.loc[subset_bkIndex_vendor_3, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                         glData.loc[subset_glIndex_vendor_2, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                         mapped_glIndex_vendor = mapped_glIndex_vendor + list(subset_glIndex_vendor_2)
#                         mapped_bankIndex_vendor = mapped_bankIndex_vendor + list(subset_bkIndex_vendor_3)
#
#     return mapped_bankIndex_vendor, mapped_glIndex_vendor

def reimbursement_mapping(bankData_potentialTS, bankData_TSBatch, bankData, glData_reimbursement, glData, df_reimPay, account_number, month_period):
    # 设置初始值
    id_number_reim = 0
    mapped_glIndex_reim = []
    mapped_bankIndex_reim = []
    # 获取本entity下的报销mapping表
    entity = accountNo_to_entity[f'{account_number}']
    df_reimPay_filtered = df_reimPay[df_reimPay['Entity'].str.contains(f'{entity}')]
    # 将报销mapping表中的信息按月和GL匹配
    for month, df_reimPay_perM in df_reimPay_filtered.groupby('Month'):
        # 跳过不在month period里的月份
        if month not in month_period:
            continue
        print(month)
        list_staffName = set(df_reimPay_perM['Staff Name'].to_list())
        gl_mapped = []
        gl_mapped_index = []
        for staff in list_staffName:
            print(staff)
            # 获取该月份mapping表中每一位员工的payment amount总和
            sum_pay = df_reimPay_perM.loc[df_reimPay_perM['Staff Name'] == f'{staff}', 'Payment Amount'].sum()
            print(sum_pay)
            # 将reimbursement payment info与gl进行比对
            gl_perStaff_mapped = False
            # 按月份和员工名筛出GL里的金额，形成带有index和payment amount的字典
            glData_reim_staffperM = glData_reimbursement.loc[(glData_reimbursement['Staff Name'] == f'{staff}') & (
                glData_reimbursement['JE Headers Description'].str.contains(f'{month}', case=False))]
            glValue_list_reim = glData_reim_staffperM['Amount Func Cur'].to_dict()
            # glValue_list_reim = glData_reimbursement.loc[(glData_reimbursement['Staff Name'] == f'{staff}') & (glData_reimbursement['JE Headers Description'].str.contains(f'{month}', case=False)), 'Amount Func Cur'].to_dict()
            sum_gl_reim = sum(glValue_list_reim.values())
            print(sum_gl_reim)
            # 判断是否匹配上
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
                    if sum_pay + sum(subset_glValue_reim) < 0.02 and sum_pay + sum(subset_glValue_reim) > -0.02:
                        if common_data(mapped_glIndex_reim, subset_glIndex_reim):
                            pass
                        else:
                            gl_perStaff_mapped = True
                            gl_staffMapped_index = list(subset_glIndex_reim)
                            gl_mapped_index = gl_mapped_index + gl_staffMapped_index
                            print('subset mapped', gl_mapped_index)
                            break
            # to be tested
            #     #把glData_reim按JE Header Id分组，形成invoice到value的字典，以减少子集计算量
            #     JE_list = set(glData_reimbursement.loc[glValue_list_reim.keys(), 'JE Header Id'].to_list())
            #     print('JE Header Id', JE_list)
            #     #获取所有JE的子集
            #     subsets_glJE_reim = get_sub_set(JE_list)
            #     print(subsets_glJE_reim)
            #     for subset_glJE_reim in subsets_glJE_reim:
            #         subset_glIndex_reim = glData_reim_staffperM.loc[glData_reim_staffperM['JE Header Id'].isin(subset_glJE_reim), 'Amount Func Cur'].index.values
            #         subset_glValue_reim = glData_reim_staffperM.loc[subset_glIndex_reim, 'Amount Func Cur'].sum()
            #         print('index', subset_glIndex_reim)
            #         print('Sum', subset_glValue_reim)
            #         if sum_pay + sum(subset_glValue_reim) < 0.02 and sum_pay + sum(subset_glValue_reim) > -0.02:
            #             if common_data(mapped_glIndex_reim, subset_glIndex_reim):
            #                 pass
            #             else:
            #                 gl_perStaff_mapped = True
            #                 gl_staffMapped_index = list(subset_glIndex_reim)
            #                 gl_mapped_index = gl_mapped_index + gl_staffMapped_index
            #                 print('subset mapped', gl_mapped_index)
            #                 break
            gl_mapped.append(gl_perStaff_mapped)
        # 将reimbursement payment info与bk进行比对
        if account_number == '001-221076-031':
            list_bankCharge = [0, -10, -20, -30, -40]
            bk_mapped = []
            bk_mapped_index = []
            bk_valueMapped_index = []
            valueMappedIndex_to_PIR = {}
            exactMappedIndex_to_PIR = {}
            for staff in list_staffName:
                # if staff != 'JAY HUANG':
                #     continue
                print(staff)
                list_staffPay = df_reimPay_perM.loc[
                    df_reimPay_perM['Staff Name'] == f'{staff}', 'Payment Amount'].to_list()
                print('list_staffPay', list_staffPay)
                bk_perStaff_mapped = False
                for payment_amount in list_staffPay:
                    print(payment_amount)
                    mappedIndex_number = 0
                    mappedIndex = []
                    for bkIndex, bkValue in bankData_potentialTS['Credit/Debit amount'].to_dict().items():
                        if bkIndex in bk_mapped_index:
                            continue
                        if (payment_amount + bkValue) in list_bankCharge:
                            print('matched', payment_amount + bkValue)
                            # if common_data(mapped_bankIndex_reim, list(bkValue_list_reim.keys())):
                            mappedIndex_number += 1
                            mappedIndex.append(bkIndex)
                            print('mappedIndex', mappedIndex)
                            print('mappedIndex_number', mappedIndex_number)
                    if mappedIndex_number == 1:
                        bk_perStaff_mapped = True
                        bk_mapped_index = bk_mapped_index + mappedIndex
                        print('bk_mapped_index', bk_mapped_index)
                        bkMappedIndex_to_PIR = dict.fromkeys(mappedIndex, f'{staff}')
                        exactMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                    if mappedIndex_number > 1:
                        bk_perStaff_mapped = True
                        bkMappedIndex_to_PIR = dict.fromkeys(mappedIndex, f'{staff}')
                        valueMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                bk_mapped.append(bk_perStaff_mapped)

        if account_number != '001-221076-031':
            list_PIRnumber = set(df_reimPay_perM['PIR Number'].to_list())
            print(list_PIRnumber)
            bk_mapped = []
            bk_mapped_index = []
            valueMappedIndex_to_PIR = {}
            exactMappedIndex_to_PIR = {}
            for number_PIR in list_PIRnumber:
                print(number_PIR)
                sum_pay_perPIR = df_reimPay_perM.loc[
                    df_reimPay_perM['PIR Number'] == f'{number_PIR}', 'Payment Amount'].sum()
                print(sum_pay_perPIR)
                bkValue_list_reim = bankData_potentialTS.loc[
                    bankData_potentialTS['Credit/Debit amount'] == round(-sum_pay_perPIR,
                                                                         2), 'Credit/Debit amount'].to_dict()
                print(bkValue_list_reim)
                bk_perPIR_mapped = False
                if len(bkValue_list_reim) == 1:
                    # test: mapped_bankIndex_reim改成bk_mapped_index
                    if common_data(mapped_bankIndex_reim, list(bkValue_list_reim.keys())):
                        pass
                    else:
                        bk_perPIR_mapped = True
                        bkMappedIndex_to_PIR = dict.fromkeys(list(bkValue_list_reim.keys()), f'{number_PIR}')
                        exactMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                        bk_mapped_index = bk_mapped_index + list(bkValue_list_reim.keys())
                        print('bk_mapped_index', bk_mapped_index)
                        print('bkMappedIndex_to_PIR', bkMappedIndex_to_PIR)
                if len(bkValue_list_reim) >= 2:
                    index_in_bkTSBatch = set(bkValue_list_reim.keys()).intersection(
                        set(bankData_TSBatch.index.tolist()))
                    print(index_in_bkTSBatch)
                    if len(index_in_bkTSBatch) == 1:
                        bk_perPIR_mapped = True
                        bkMappedIndex_to_PIR = dict.fromkeys(list(bkValue_list_reim.keys()), f'{number_PIR}')
                        exactMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                        bk_mapped_index = bk_mapped_index + index_in_bkTSBatch
                        print('bk_mapped_index-choose from 2 index', bk_mapped_index)
                        print('bkMappedIndex_to_PIR-choose from 2 index', bkMappedIndex_to_PIR)
                    else:
                        bk_perPIR_mapped = True
                        bkMappedIndex_to_PIR = dict.fromkeys(list(bkValue_list_reim.keys()), f'{number_PIR}')
                        valueMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                        # bk_valueMapped_index = bk_valueMapped_index + list(bkValue_list_reim.keys())
                        print('bk_valueMapped_index', valueMappedIndex_to_PIR)
                bk_mapped.append(bk_perPIR_mapped)
        print('exactMappedIndex_to_PIR', exactMappedIndex_to_PIR)
        print('valueMappedIndex_to_PIR', valueMappedIndex_to_PIR)
        print('bk Condition:', False not in bk_mapped, bk_mapped)
        print('gl Condition:', False not in gl_mapped, gl_mapped)
        if (False not in bk_mapped) and (False not in gl_mapped):
            id_number_reim += 1
            # bankData.loc[bk_mapped_index, 'notes'] = f'reimbursement netoff {now} {month} {id_number_reim}'
            for key in exactMappedIndex_to_PIR.keys():
                bankData.loc[
                    key, 'notes'] = f'reimbursement netoff {now} {month} {id_number_reim} {exactMappedIndex_to_PIR[key]}'
            for key in valueMappedIndex_to_PIR.keys():
                bankData.loc[
                    key, 'notes'] = f'reimbursement netoff {now} {month} {id_number_reim} - value map {valueMappedIndex_to_PIR[key]}'
            # bankData.loc[list(valueMappedIndex_to_PIR.keys()), 'notes'] = f'reimbursement netoff {now} {id_number_reim} - value map {}'
            glData.loc[gl_mapped_index, 'notes'] = f'reimbursement netoff {now} {month} {id_number_reim}'
            mapped_bankIndex_reim = mapped_bankIndex_reim + bk_mapped_index
            mapped_glIndex_reim = mapped_glIndex_reim + gl_mapped_index
        if (False in bk_mapped) and (False not in gl_mapped):
            for key in exactMappedIndex_to_PIR.keys():
                bankData.loc[key, 'notes'] = f'reimbursement payment {now} {month} {exactMappedIndex_to_PIR[key]}'
            for key in valueMappedIndex_to_PIR.keys():
                bankData.loc[
                    key, 'notes'] = f'reimbursement payment {now} {month} {bkMappedIndex_to_PIR[key]} - value map'
            glData.loc[gl_mapped_index, 'notes'] = f'reimbursement payment {now} {month}'
            mapped_bankIndex_reim = mapped_bankIndex_reim + bk_mapped_index
            mapped_glIndex_reim = mapped_glIndex_reim + gl_mapped_index

    return mapped_bankIndex_reim, mapped_glIndex_reim, valueMappedIndex_to_PIR


# directory_BS = input("Please enter the folder directory of all the BS statements:")
# directory_GL = input("Please enter the folder directory of all the GL files:")
# directory_AP_Vendor = input("Please enter the file link of the AP_Vendor Mapping:")
# directory_AP_Employee = input("Please enter the file link of the AP_Employee Mapping:")
# directory_Commercial = input("Please enter the file link of the Commercial Mapping")
# month_period = input("Please enter the covered month periods:")


path_folder_BS = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Bank Statement'
path_folder_GL = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\GL'
path_folder_reimRegister = r'C:\Users\he kelly\Desktop\Alteryx & Python\reimbursement\New folder'
directory_AP_Vendor = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\AP Mapping.xlsx'
directory_AP_Employee = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\Employee mapping.xlsx'
directory_Commercial = r'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\Mapping\Cash receipt 2023.xlsx'
month_period = 'JAN FEB MAR'

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


#获取vendor mapping
map_vendor = pd.read_excel(directory_AP_Vendor, header=0)
map_vendor['Vendor Name'] = map_vendor['Vendor Name'].map(lambda x: x.upper())
map_employee = pd.read_excel(directory_AP_Employee, header=1)
accountNo_to_vendorSite = {'626-055784-001': 'Beijing OU', '622-512317-001': 'Shenzhen OU', '088-169370-011': 'China PRC OU', '001-221076-031': 'Taiwan OU'}


#获取employee mapping
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
accountNo_to_entity = {'626-055784-001': 'Beijing LE', '622-512317-001': 'BCG Shenzhen LE', '088-169370-011': 'China PRC LE', '001-221076-031': 'Taiwan LE'}

#读取Commercial mapping, 创建mapping dictionary
map_commercial = pd.read_excel(directory_Commercial, header=0)
map_commercial['Actual Receipt  Amount'].fillna(method='ffill', axis=0, inplace=True)
map_commercial['Receipt Dt'] = map_commercial['Receipt Dt'].astype('datetime64[ns]')
map_commercial['bank expense'] = map_commercial['bank expense'].astype('float')
map_commercial['Client Name'] = map_commercial['Client Name'].dropna().map(lambda x: x.upper())
tb_location = {'088-169370-011': 'PRC', '626-055784-001': 'Beijing', '622-512317-001': 'Shenzhen', '001-221076-031': 'Taipei'}

list_bankCharge = [0, 10, 20, 25, 30, 35, 40, 50, 60, 70, 80]

#PRC Section
bank_mapping_PRC = {'088-169370-011': '101244', '626-055784-001': '101001', '622-512317-001': '101135', '001-221076-031': '101245'}
# bank_mapping_PRC = {'626-055784-001': '101001', '622-512317-001': '101135', '088-169370-011': '101244'}
# bank_mapping_PRC = {'088-169370-011': '101244', '626-055784-001': '101001', '622-512317-001': '101135'}
# bank_mapping_PRC = {'622-512317-001': '101135'}
# bank_mapping_PRC = {'088-169370-011': '101244', '001-221076-031': '101245'}
for account_number, account_cd in bank_mapping_PRC.items():

    print('account_cd Start mapping', account_cd)

    #获取当前bank account的bank和gl数据
    bankData = df_bank[df_bank['Account number'] == f'{account_number}']
    glData = df_GL[df_GL['Account Cd'] == int(account_cd)]

    #处理无需mapping的type,并筛选需mapping的df
    if account_cd != '101245':
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
    else:
        bankData_charges = bankData[bankData['TRN type']=='Charges']
        bankData.loc[bankData_charges.index, 'notes'] = 'bank charges'
        bankData_interest = bankData[bankData['TRN type']=='Interest']
        bankData.loc[bankData_interest.index, 'notes'] = 'bank interest'
        index_filtered = list(set(bankData.index).difference(set(list(bankData_charges.index)+list(bankData_interest.index)))) #改名字 index_emptyNotes
        bankData_filtered = bankData.iloc[index_filtered] #改名字

    location = tb_location[bankData.iloc[1]['Account number']]

    #employee and vendor mapping
    #获取当前entity的employee and vendor mapping
    map_employee_byEntity = map_employee.loc[map_employee['Vendor Site OU'] == f'{accountNo_to_vendorSite[account_number]}']
    map_vendor_byEntity = map_vendor.loc[map_vendor['Vendor Site OU'] == f'{accountNo_to_vendorSite[account_number]}']
    excel_log.log(map_vendor_byEntity, 'map_vendor_byEntity')


    #commercial mapping
    bankData_commercial = bankData_filtered.loc[bankData_filtered['Credit/Debit amount']>0, :] #排除bk金额小于等于0.03的item
    glData_commercial = glData[glData['JE Headers Description'].str.contains('Cash Receipts')]

    mapped_bankIndex_commercial, mapped_glIndex_commercial = commercial_mapping(bankData_commercial, bankData, glData_commercial, glData, map_commercial, account_cd)

    #AP退款重付
    bankIndex_repayment_netoff = repayment_mapping(bankData_filtered, bankData, 'first')

    print('bankIndex_repayment_netoff', bankIndex_repayment_netoff)
    print('mapped_bankIndex_commercial', mapped_bankIndex_commercial)

    #AP Mapping
    #挖掉已匹上的commercial部分
    bankData_AP = bankData_filtered.loc[bankData_filtered.index.difference(mapped_bankIndex_commercial)]
    #挖掉退款重复的部分
    bankData_AP = bankData_AP.loc[bankData_AP.index.difference(bankIndex_repayment_netoff)]
    glData_AP = glData.loc[glData['View Source'].str.contains('Payables', regex=False, case=False, na=False)]
    glData_staff = glData[glData['Vendor Name'].str.contains('                              ', regex=False, case=False, na=False)]
    # excel_log.log(glData_staff, 'staff list')
    glData_reimbursement = pd.DataFrame()
    staff_invoice_indication = ['HLYERR', 'TB', 'RVCR', 'CM', 'CR']
    for item in staff_invoice_indication:
        df_staff = glData_AP[glData_AP['Invoice Number'].str.contains(f'{item}', regex=False, case=False, na=False)]
        glData_reimbursement = pd.concat([glData_reimbursement, df_staff])
    glData_reimbursement['Staff Name'] = glData_reimbursement['Vendor Name'].map(lambda x: x.split('      ')[0])
    glData_staff['Staff Name'] = glData_staff['Vendor Name'].map(lambda x: x.split('      ')[0])
    # excel_log.log(glData_reimbursement, 'reimbursement list')
    glData_vendor = glData_AP.loc[glData_AP.index.difference(glData_staff.index)]
    # excel_log.log(glData_vendor, 'vendor list')

    mapped_bankIndex_vendor1, mapped_glIndex_vendor1 = AP_mapping(bankData_AP, bankData, glData_vendor, glData, map_vendor_byEntity, account_cd, list_bankCharge, 'Name', 'vendor')

    bankData_AP_left = bankData_AP.loc[bankData_AP.index.difference(mapped_bankIndex_vendor1)]
    glData_vendor_left = glData_vendor.loc[glData_vendor.index.difference(mapped_glIndex_vendor1)]

    mapped_bankIndex_vendor2, mapped_glIndex_vendor2 = AP_mapping(bankData_AP_left, bankData, glData_vendor_left, glData, map_vendor_byEntity, account_cd, list_bankCharge, 'Num', 'vendor')

    bankData_AP_left3 = bankData_AP_left.loc[bankData_AP_left.index.difference(mapped_bankIndex_vendor2)]

    #挖出剩下的bankData里vendor的部分
    bankData_leftVendor = pd.DataFrame()
    for accountNum in map_vendor['Bank Account Num']:
        bankData_vendor = bankData_AP_left3[bankData_AP_left3['Narrative'].str.contains(f'{accountNum}', regex=False, case=False, na=False)]
        bankData_leftVendor = pd.concat([bankData_leftVendor, bankData_vendor])

    #挖出剩下的bankData里client的部分
    bankData_leftClient = pd.DataFrame()
    for clientName in map_commercial['Client Name in Chinese']:
        bankData_client = bankData_AP_left3[bankData_AP_left3['Narrative'].str.contains(f'{clientName}', regex=False, case=False, na=False)]
        bankData_leftClient = pd.concat([bankData_leftClient, bankData_client])

    #在bkData中filter出TS batch
    #修改点：增加台湾的bankData_TSBatch
    bankData_SBID = bankData_AP_left3[bankData_AP_left3['Bank reference'].str.contains('SBID', regex=False, case=False, na=False)]
    bankData_TSBatch = bankData_SBID
    bankData_TSBatch['Keyword'] = bankData_TSBatch['Narrative'].map(lambda x: x.split('/')[2])
    keyword_nonTS = ['COL', 'Intern', 'PTA', 'Payroll', 'Bonus', 'Cash advance']
    for keyword in keyword_nonTS:
        bankData_keyword = bankData_TSBatch[bankData_TSBatch['Keyword'].str.contains(f'{keyword}', regex=False, case=False, na=False)]
        bankData_TSBatch = bankData_TSBatch.loc[bankData_TSBatch.index.difference(bankData_keyword.index)]


    bankData_potentialTS = bankData_AP_left3.loc[bankData_AP_left3.index.difference(bankData_leftVendor.index)]
    bankData_potentialTS = bankData_potentialTS.loc[bankData_potentialTS.index.difference(bankData_leftClient.index)]

    mapped_bankIndex_reim, mapped_glIndex_reim, valueMappedIndex_to_PIR = reimbursement_mapping(bankData_potentialTS, bankData_TSBatch, bankData, glData_reimbursement, glData, df_reimPay, account_number, month_period)



    # Staff Mapping
    # 在bkData potentialTS里挖去已匹配的报销部分
    bankData_potentialStaff = bankData_potentialTS.loc[bankData_potentialTS.index.difference(mapped_bankIndex_reim)]
    # 在glData Staff里挖去已匹配的报销部分
    glData_staff = glData_staff.loc[glData_staff.index.difference(mapped_glIndex_reim)]
    # excel_log.log(bankData_potentialStaff, 'bankData_potentialStaff')
    # excel_log.log(glData_staff, 'glData_Staff')
    mapped_bankIndex_staff1, mapped_glIndex_staff1 = AP_mapping(bankData_potentialStaff, bankData, glData_staff, glData, map_employee_byEntity, account_cd, list_bankCharge, 'Name', 'staff')

    bankData_potentialStaff_left = bankData_potentialStaff.loc[bankData_potentialStaff.index.difference(mapped_bankIndex_staff1)]
    glData_staff_left = glData_staff.loc[glData_staff.index.difference(mapped_glIndex_staff1)]

    mapped_bankIndex_staff2, mapped_glIndex_staff2 = AP_mapping(bankData_potentialStaff_left, bankData, glData_staff_left, glData, map_employee_byEntity, account_cd, list_bankCharge, 'Num', 'staff')
    # # 设置初始值
    # id_number_staff = 0
    # mapped_glIndex_staff = []
    # mapped_bankIndex_staff = []
    #
    # # Staff Mapping 1 GL总和匹BK子集
    # for staff, df in glData_staff.groupby('Vendor Name'):
    #     print(staff)
    #     glSum_staff = df['Amount Func Cur'].sum()
    #     print(glSum_staff)
    #     bankAccountNum_staff = map_employee_byEntity.loc[map_employee_byEntity['Vendor Name'] == f'{staff}'.upper(), 'Bank Account Num']
    #     print('bankAccountNum_staff', bankAccountNum_staff)
    #     # 当一个staff name匹配一个银行账号名时
    #     if len(set(bankAccountNum_staff.to_list())) == 1:
    #         for num in set(bankAccountNum_staff.to_list()):
    #             pro_num = str(num).strip()
    #             bkValue_list_staff = bankData_potentialStaff.loc[bankData_potentialStaff["Narrative"].str.contains(f'{pro_num}', regex=False,
    #                                                                                     case=False), 'Credit/Debit amount'].to_dict()
    #             print('bkValue_list_staff', bkValue_list_staff)
    #             subsets_bkIndex_staff = get_sub_set(bkValue_list_staff.keys())
    #             for subset_bkIndex_staff in subsets_bkIndex_staff:
    #                 subset_bkValue_staff = key_to_value(subset_bkIndex_staff, bkValue_list_staff)
    #                 print('subset_bkValue_staff', subset_bkValue_staff)
    #                 if ((sum(subset_bkValue_staff) - glSum_staff) <= 0.03) & ((sum(subset_bkValue_staff) - glSum_staff) >= -0.03):
    #                     if common_data(subset_bkValue_staff, mapped_bankIndex_staff) or common_data(list(df.index), mapped_glIndex_staff):
    #                         pass
    #                     else:
    #                         id_number_staff = id_number_staff + 1
    #                         bankData.loc[subset_bkIndex_staff, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                         glData.loc[df.index, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                         mapped_glIndex_staff = mapped_glIndex_staff + list(df.index)
    #                         mapped_bankIndex_staff = mapped_bankIndex_staff + subset_bkIndex_staff
    #                         print('mapped_glIndex_staff', mapped_glIndex_staff)
    #                         print('mapped_bankIndex_staff', mapped_bankIndex_staff)
    #                         break
    #
    #     # 当一个vendor name匹配多个银行账号时
    #     if len(set(bankAccountNum_staff.to_list())) > 1:
    #         bkIndex_staff = []
    #         dic_bkValue_staff = {}
    #         for num in set(bankAccountNum_staff.to_list()):
    #             pro_num = str(num).strip()
    #             bkValue_list_staff = bankData_potentialStaff.loc[bankData_potentialStaff["Narrative"].str.contains(f'{pro_num}', regex=False,
    #                                                                                     case=False), 'Credit/Debit amount'].to_dict()
    #             bkIndex_staff = bkIndex_staff + list(bkValue_list_staff.keys())
    #             dic_bkValue_staff.update(bkValue_list_staff)
    #         print(bkIndex_staff)
    #         subsets_bkIndex_staff = get_sub_set(bkIndex_staff)
    #         for subset_bkIndex_staff in subsets_bkIndex_staff:
    #             subset_bkValue_staff = key_to_value(subset_bkIndex_staff, dic_bkValue_staff)
    #             if ((sum(subset_bkValue_staff) - glSum_staff) <= 0.03) & (
    #                     (sum(subset_bkValue_staff) - glSum_staff) >= -0.03):
    #                 if common_data(subset_bkValue_staff, mapped_bankIndex_staff) or common_data(list(df.index), mapped_glIndex_staff):
    #                     pass
    #                 else:
    #                     id_number_staff = id_number_staff + 1
    #                     bankData.loc[subset_bkIndex_staff, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                     glData.loc[df.index, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                     mapped_glIndex_staff = mapped_glIndex_staff + list(df.index)
    #                     mapped_bankIndex_staff = mapped_bankIndex_staff + subset_bkIndex_staff
    #                     break
    #
    # #Staff gl和bk挖去第一次匹配后的结果
    # bankData_potentialStaff_left1 = bankData_potentialStaff.loc[bankData_potentialStaff.index.difference(mapped_bankIndex_staff)]
    # glData_staff_left1 = glData_staff.loc[glData_staff.index.difference(mapped_glIndex_staff)]
    # excel_log.log(bankData_potentialStaff_left1, 'bankData_potentialStaff_left1')
    # excel_log.log(glData_staff_left1, 'glData_staff_left1')
    #
    # circle_number = 0
    # # Staff Mapping 2 bk总和匹gl子集
    # for staff, df in glData_staff_left1.groupby('Vendor Name'):
    #     circle_number += 1
    #     print('circle_number', circle_number)
    #     bankAccountNum_staff1 = map_employee_byEntity.loc[map_employee_byEntity['Vendor Name'] == f'{staff}'.upper(), 'Bank Account Num']
    #     print('bankAccountNum_staff1', bankAccountNum_staff1)
    #     dic_bkValue_staff1= {}
    #     if len(set(bankAccountNum_staff1.to_list())) >= 1:
    #         for num in set(bankAccountNum_staff1.to_list()):
    #             bkValue_list_staff1 = bankData_potentialStaff_left1.loc[bankData_potentialStaff_left1['Narrative'].str.contains(f'{str(num).strip()}', regex=False, case=False, na=False), 'Credit/Debit amount'].to_dict()
    #             print('bkValue_list_staff1', bkValue_list_staff1)
    #             dic_bkValue_staff1.update(bkValue_list_staff1)
    #     if len(dic_bkValue_staff1) >= 1:
    #         print('dic_bkValue_staff1', dic_bkValue_staff1)
    #         bk_sum = sum(dic_bkValue_staff1.values())
    #         print('bk_sum', bk_sum)
    #         bk_Index = list(dic_bkValue_staff1.keys())
    #         glValue_list_staff1 = df['Amount Func Cur'].to_dict()
    #         subsets_glIndex_staff1 = get_sub_set(glValue_list_staff1.keys())
    #         print('subsets_glIndex_staff1', subsets_glIndex_staff1)
    #         for subset_glIndex_staff1 in subsets_glIndex_staff1:
    #             subset_glValue_staff1 = key_to_value(subset_glIndex_staff1, glValue_list_staff1)
    #             print('subset_glValue_staff1', subset_glValue_staff1)
    #             if (sum(subset_glValue_staff1) - bk_sum <= 0.03) & (sum(subset_glValue_staff1) - bk_sum >= -0.03):
    #                 print('True')
    #                 if common_data(subset_glIndex_staff1, mapped_glIndex_staff) or common_data(bk_Index, mapped_bankIndex_staff):
    #                     pass
    #                 else:
    #                     print('Mapped')
    #                     id_number_staff = id_number_staff + 1
    #                     bankData.loc[bk_Index, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                     glData.loc[subset_glIndex_staff1, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                     mapped_glIndex_staff = mapped_glIndex_staff + list(subset_glIndex_staff1)
    #                     mapped_bankIndex_staff = mapped_bankIndex_staff + bk_Index
    #                     break
    # print('completed')
    # # staff gl和bk挖去第二次匹配后的结果
    # bankData_potentialStaff_left2 = bankData_potentialStaff_left1.loc[bankData_potentialStaff_left1.index.difference(mapped_bankIndex_staff)]
    # glData_staff_left2 = glData_staff_left1.loc[glData_staff_left1.index.difference(mapped_glIndex_staff)]
    # #
    # #staff gl子集和bk子集匹配
    # for staff, df in glData_staff_left2.groupby('Vendor Name'):
    #     bankAccountNum_staff2 = map_employee_byEntity.loc[map_employee_byEntity['Vendor Name'] == f'{staff}'.upper(), 'Bank Account Num']
    #     print('bankAccountName_staff2', bankAccountNum_staff2)
    #     dic_bkValue_staff2 = {}
    #     if len(set(bankAccountNum_staff2.to_list())) >= 1:
    #         for num in set(bankAccountNum_staff2.to_list()):
    #             bkValue_list_staff2 = bankData_potentialStaff_left2.loc[bankData_potentialStaff_left2['Narrative'].str.contains(f'{str(num).strip()}', regex=False, case=False, na=False), 'Credit/Debit amount'].to_dict()
    #             dic_bkValue_staff2.update(bkValue_list_staff2)
    #     if len(dic_bkValue_staff2):
    #         subsets_bkIndex_staff2 = get_sub_set(dic_bkValue_staff2.keys())
    #         glValue_list_staff2 = df['Amount Func Cur'].to_dict()
    #         subsets_glIndex_staff2 = get_sub_set(glValue_list_staff2.keys())
    #         for subset_bkIndex_staff2 in subsets_bkIndex_staff2:
    #             subset_bkValue_staff2 = key_to_value(subset_bkIndex_staff2, dic_bkValue_staff2)
    #             for subset_glIndex_staff2 in subsets_glIndex_staff2:
    #                 subset_glValue_staff2 = key_to_value(subset_glIndex_staff2, glValue_list_staff2)
    #                 if (sum(subset_glValue_staff2) - sum(subset_bkValue_staff2) <= 0.03) & (sum(subset_glValue_staff2) - sum(subset_bkValue_staff2) >= -0.03):
    #                     if common_data(subset_glIndex_staff2, mapped_glIndex_vendor) or common_data(
    #                             subset_bkIndex_staff2, mapped_bankIndex_vendor):
    #                         pass
    #                     else:
    #                         id_number_staff = id_number_staff + 1
    #                         bankData.loc[subset_bkIndex_staff2, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                         glData.loc[subset_glIndex_staff2, 'notes'] = f'staff netoff {now} {id_number_staff}'
    #                         mapped_glIndex_staff = mapped_glIndex_staff + list(subset_glIndex_staff2)
    #                         mapped_bankIndex_staff = mapped_bankIndex_staff + list(subset_bkIndex_staff2)
    #
    #
    # bankData_left = bankData_AP_left3.loc[bankData_AP_left3.index.difference(mapped_bankIndex_reim)]
    # # bankData_left = bankData_left.loc[bankData_left.index.difference(mapped_bankIndex_staff)]
    # excel_log.log(bankData_left, 'bankData_left')
    #
    # bankIndex_repayment_netoff2 = repayment_mapping(bankData_left, bankData, 'last')
    #
    #



    now_for_folder = now.replace(':', ' ')
    os.makedirs(rf'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\{now_for_folder}\{location}')
    bankData.to_excel(fr'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\{now_for_folder}\{location}\bank_{location}_{account_number}.xlsx')
    glData.to_excel(fr'C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\test\{now_for_folder}\{location}\gl_{location}_{account_cd}.xlsx')


















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














