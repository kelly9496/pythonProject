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
        dataframe.to_excel(rf'C:\Users\he kelly\Desktop\bank_reconciliation_py\Bank Rec Program\debug\{desc}_{now}.xlsx')
        self.log_number += 1

excel_log = ExcelLog(10)


#定义所需函数
def get_sub_set(nums):
    # print('nums', nums)
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
    print('1st round commercial mapping')
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
                                    index_list_gl = list(value_list_gl.keys())
                                    subsets_index_gl = get_sub_set(index_list_gl)
                                    for subset_index_gl in subsets_index_gl:
                                        subset_value_gl = key_to_value(subset_index_gl, value_list_gl)
                                        # 若筛出的glData的某个子集之和等于某一入账时间的project code对应的mapping表总和
                                        if sum(subset_value_gl) == sum_value_map:
                                            for index in subset_index_gl:
                                                if (index in glIndex_mappedToInd):
                                                    break
                                                else:
                                                    # print('recorded index:', index)
                                                    glIndex_mappedToInd.append(index)
                                                    break

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
    glData_commercial_left = glData_commercial_left.groupby('Vendor Name')

    mapped_glIndex_commercial_2 = []
    mapped_bankIndex_commercial_2 = []

    # 第二轮commercial mapping
    print('2nd round commercial mapping')
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
def commercial_mapping_TW(bankData_commercial, bankData, glData_commercial, glData, map_commercial):
    location = tb_location[bankData_commercial.iloc[1]['Account number']]
    # 处理commercial mapping表，筛出本entity RPApo账的部分，并按project ID分类
    map_commercial_RPA = map_commercial[map_commercial['location'].str.contains(f'{location}') & (map_commercial['Currency'] == 'TWD') & (map_commercial['Notification Email'] != '-')]
    map_commercial_RPA = map_commercial_RPA.groupby(['Receipt Dt', 'Actual Receipt  Amount'])

    # 设定初始值
    id_number_commercial = 0
    mapped_glIndex_commercial_1 = []
    mapped_bankIndex_commercial_1 = []

    print('1st round commercial mapping')
    # 第一轮commercial mapping
    for ind, row in bankData_commercial.iterrows():
        bank_value = row['Credit/Debit amount']
        bank_receipt_date = row['Value date']
        for group_condition, df_map in map_commercial_RPA:
            if group_condition[0] == bank_receipt_date:
                map_sum_byProject = df_map.groupby('Project ID').sum('AR in Office Currency')['AR in Office Currency'].to_dict()
                project_code_list = list(map_sum_byProject.keys())
                # 对加总值匹上的project code进行循环
                glIndex_mappedToInd = []
                for project_id in project_code_list:
                    df_map_grouped = df_map.groupby(['Notification Email', 'Project ID'])
                    for filter_condition, df in df_map_grouped:
                        if filter_condition[1] == project_id:
                            map_clear_date = df.iloc[0]['Notification Email']
                            # 获得某一入账时间的project code对应的mapping表总和，该值与GL的子集进行比对
                            sum_value_map = df['AR in Office Currency'].sum()
                            # 筛出还未mapping过的glData
                            glData_commercial_filtered = glData_commercial.loc[glData_commercial.index.difference(mapped_glIndex_commercial_1)]
                            # 对还未mapping上的glData用入账时间和project code进行初步筛选
                            glData_commercial_filtered = glData_commercial_filtered[(glData_commercial_filtered['JH Created Date'] < map_clear_date + dt.timedelta(days=8)) & (glData_commercial_filtered['JH Created Date'] > map_clear_date - dt.timedelta(days=8)) & (glData_commercial_filtered['Project Id'] == f'{project_id}')]
                            value_list_gl = glData_commercial_filtered['Amount Func Cur'].to_dict()
                            index_list_gl = list(value_list_gl.keys())
                            subsets_index_gl = get_sub_set(index_list_gl)
                            for subset_index_gl in subsets_index_gl:
                                subset_value_gl = key_to_value(subset_index_gl, value_list_gl)
                                # 若筛出的glData的某个子集之和等于某一入账时间的project code对应的mapping表总和
                                if sum(subset_value_gl) == sum_value_map:
                                    for index in subset_index_gl:
                                        if (index in glIndex_mappedToInd):
                                            break
                                        else:
                                            glIndex_mappedToInd.append(index)
                                            break

                        if len(glIndex_mappedToInd) and df_map['Actual Receipt  Amount'].sum() == bank_value:
                            if (ind in mapped_bankIndex_commercial_1) or common_data(glIndex_mappedToInd, mapped_glIndex_commercial_1):
                                pass
                            else:
                                id_number_commercial = id_number_commercial + 1
                                bankData.loc[ind, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                                glData.loc[glIndex_mappedToInd, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                                mapped_bankIndex_commercial_1.append(ind)
                                mapped_glIndex_commercial_1 = mapped_glIndex_commercial_1 + glIndex_mappedToInd

    # #获取第一轮mapping之后剩余部分的glData和bankData
    bankIndex_commercial_left = list(
        set(list(bankData_commercial.index)).difference(set(mapped_bankIndex_commercial_1)))
    glIndex_commercial_left = list(set(list(glData_commercial.index)).difference(set(mapped_glIndex_commercial_1)))

    glData_commercial_left = glData_commercial.loc[glIndex_commercial_left, :]
    bankData_commercial_left = bankData_commercial.loc[bankIndex_commercial_left, :]
    glData_commercial_left = glData_commercial_left.groupby('Vendor Name')

    mapped_glIndex_commercial_2 = []
    mapped_bankIndex_commercial_2 = []

    print('2nd round commercial mapping')
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
def commercial_mapping_HK(bankData_commercial, bankData, glData_commercial, glData, map_commercial, account_cd):
    location = tb_location[bankData_commercial.iloc[1]['Account number']]
    # 处理commercial mapping表，筛出本entity RPApo账的部分，并按project ID分类
    map_commercial_RPA = map_commercial[map_commercial['location'].str.contains(f'{location}') & (map_commercial['Currency'] == 'USD') & (map_commercial['Notification Email'] != '-')]
    # excel_log.log(map_commercial_RPA, 'map_commercial_PRC')
    map_commercial_RPA['AR with charges'] = map_commercial_RPA['AR in Office Currency'] - map_commercial_RPA['bank expense']
    map_commercial_RPA = map_commercial_RPA.groupby(['Receipt Dt', 'Client Name'])

    # 设定初始值
    id_number_commercial = 0
    mapped_index_commercial = []
    mapped_glIndex_commercial_1 = []
    mapped_bankIndex_commercial_1 = []
    print('1st round commercial mapping')
    # 第一轮commercial mapping
    for ind, row in bankData_commercial.iterrows():
        bank_value = row['Credit/Debit amount']
        bank_receipt_date = row['Value date']
        for group_condition, df_map in map_commercial_RPA:
            # if group_condition[1] != 'PRIMAVERA CAPITAL':
            #     continue
            if group_condition[0] == bank_receipt_date:
                print('group_condition', group_condition)
                print('bank index', ind)
                map_sum_byProject = df_map.groupby('Project ID').sum('AR with charges')['AR with charges'].to_dict()
                print(map_sum_byProject)
                project_code_list = list(map_sum_byProject.keys())
                project_code_subsets = get_sub_set(project_code_list)
                for subset in project_code_subsets:
                    print('project code subset', subset)
                    subset_value_map = key_to_value(subset, map_sum_byProject)
                    mapped_commercial1 = False
                    if (sum(subset_value_map) - bank_value <= 0.03) & (sum(subset_value_map) - bank_value >= -0.03):
                        mapped_commercial1 = True
                        print('mapped_commercial1', mapped_commercial1)
                        # 如果mapping表中receipt date匹上的project code的subset值的汇总和银行匹配上
                        glIndex_mappedToInd = []
                        # 对加总值匹上的project code进行循环
                        df_map_grouped = df_map.groupby(['Notification Email'])
                        for filter_condition, df in df_map_grouped:
                            print('filter_condition', filter_condition)
                            map_clear_date = df.iloc[0]['Notification Email']
                            # 获得某一入账时间的project code对应的mapping表总和，该值与GL的子集进行比对
                            sum_value_map = df['AR in Office Currency'].sum()
                            print('sum_value_map',sum_value_map)
                            # 筛出还未mapping过的glData
                            glData_commercial_filtered = glData_commercial.loc[glData_commercial.index.difference(mapped_glIndex_commercial_1)]
                            # 对还未mapping上的glData用入账时间和project code进行初步筛选
                            glData_commercial_filtered_AR = glData_commercial_filtered[(glData_commercial_filtered['Posted Date'] < map_clear_date + dt.timedelta(days=8)) & (glData_commercial_filtered['Posted Date'] > map_clear_date - dt.timedelta(days=8)) & (glData_commercial_filtered[r'Vendor/Client Name'].str.contains(f'{group_condition[1][:22]}', regex=False, case=False))]
                            print(group_condition[1][:22])
                            print(glData_commercial_filtered_AR)
                            value_list_gl = glData_commercial_filtered_AR['Amount'].to_dict()
                            index_list_gl = list(value_list_gl.keys())
                            subsets_index_gl = get_sub_set(index_list_gl)
                            for subset_index_gl in subsets_index_gl:
                                subset_value_gl = key_to_value(subset_index_gl, value_list_gl)
                                print(sum(subset_value_gl))
                                # 若筛出的glData的某个子集之和等于某一入账时间的project code对应的mapping表总和
                                if sum(subset_value_gl) == sum_value_map:
                                    print('GL Mapped')
                                    for index in subset_index_gl:
                                        print(index)
                                        if (index in glIndex_mappedToInd):
                                            break
                                        else:
                                            print('recorded index:', index)
                                            glIndex_mappedToInd.append(index)
                            for value in df['bank expense'].to_list():
                                if value != 0:
                                    glData_commercial_filtered_BC = glData_commercial_filtered[(glData_commercial_filtered['Posted Date'] < map_clear_date + dt.timedelta(days=8)) & (glData_commercial_filtered['Posted Date'] > map_clear_date - dt.timedelta(days=8)) & (glData_commercial_filtered['Amount'] == -value)]
                                    if len(glData_commercial_filtered_BC):
                                        list_index_mappedBC = list(set(glData_commercial_filtered_BC.index).difference(set(glIndex_mappedToInd)))
                                        glIndex_mappedToInd.append(list_index_mappedBC[0])



                        glData_sum_mappedToInd = glData_commercial.loc[glIndex_mappedToInd]['Amount'].sum()
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
    glData_commercial_left = glData_commercial_left.groupby('Vendor/Client Name')

    mapped_glIndex_commercial_2 = []
    mapped_bankIndex_commercial_2 = []

    # 第二轮commercial mapping
    print('2nd round commercial mapping')
    for client, df_left in glData_commercial_left:
        sum_gl = df_left['Amount'].sum()
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

def repayment_mapping(bankData_filtered, bankData, order, account_cd, list_bankCharge):

    # AP退款重付
    id_number_repayment = 0
    bankIndex_repayment_netoff = []
    # 找出含有RJCT的行
    if account_cd == '101245':
        bankData_RJCT = bankData_filtered[bankData_filtered['Narrative'].str.contains('退匯', regex=False, na=False)]
        # excel_log.log(bankData_RJCT, 'TW bankData_RJCT step1')
        bankData_RJCT['bankAccountName'] = bankData_RJCT['Narrative'].map(lambda x: x.split(' ')[1])
        # excel_log.log(bankData_RJCT, 'TW bankData_RJCT step2')
        bankData_RJCT['bankAccountName'] = bankData_RJCT['bankAccountName'].map(lambda x: x.replace('☆☆☆', ' '))
        # excel_log.log(bankData_RJCT, 'TW bankData_RJCT step3')
        # 删掉有return字眼的行
    else:
        bankData_RJCT = bankData_filtered[bankData_filtered['Narrative'].str.contains('RJCT', regex=False, na=False)]
        bankData_RJCT['bankAccountName'] = bankData_RJCT['Narrative'].map(lambda x: x.split('/')[2])
        # 删掉有return字眼的行
        bankData_RJCT.drop(bankData_RJCT[bankData_RJCT['bankAccountName'].str.contains('RETURN')].index, inplace=True)
        # excel_log.log(bankData_RJCT, 'return')
    for ind_RJCT, row_RJCT in bankData_RJCT.iterrows():
        condition_bkAccountName = bankData_filtered['Narrative'].str.contains(f'{row_RJCT["bankAccountName"]}')
        if order == 'first':
            condition_date = bankData_filtered['Value date'] < row_RJCT['Value date']
        if order == 'last':
            condition_date = bankData_filtered['Value date'] <= row_RJCT['Value date']
        if account_cd == '101245':
            condition_amount = bankData_filtered['Credit/Debit amount'] == -row_RJCT['Credit/Debit amount']
            for charge in list_bankCharge:
                if charge == 0:
                    continue
                amount_mapped = bankData_filtered['Credit/Debit amount'] == (-row_RJCT['Credit/Debit amount']-charge)
                condition_amount = condition_amount | amount_mapped
        else:
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

# def commercial_mapping(bankData_commercial, bankData, glData_commercial, glData, map_commercial, account_cd):
#     location = tb_location[bankData_commercial.iloc[1]['Account number']]
#     # 处理commercial mapping表，筛出本entity RPApo账的部分，并按project ID分类
#     map_commercial_RPA = map_commercial[map_commercial['location'].str.contains(f'{location}') & (map_commercial['Currency'] == 'CNY') & (map_commercial['Notification Email'] != '-')]
#     # excel_log.log(map_commercial_RPA, 'map_commercial_PRC')
#     if account_cd == '101245':
#         map_commercial_RPA = map_commercial_RPA.groupby(['Receipt Dt', 'Actual Receipt  Amount'])
#     else:
#         map_commercial_RPA = map_commercial_RPA.groupby(['Receipt Dt', 'Client Name'])
#
#     # 设定初始值
#     id_number_commercial = 0
#     mapped_index_commercial = []
#     mapped_glIndex_commercial_1 = []
#     mapped_bankIndex_commercial_1 = []
#
#     # 第一轮commercial mapping
#     for ind, row in bankData_commercial.iterrows():
#         bank_value = row['Credit/Debit amount']
#         bank_receipt_date = row['Value date']
#         for group_condition, df_map in map_commercial_RPA:
#             if group_condition[0] == bank_receipt_date:
#                 map_sum_byProject = df_map.groupby('Project ID').sum('AR in Office Currency')[
#                     'AR in Office Currency'].to_dict()
#                 project_code_list = list(map_sum_byProject.keys())
#                 project_code_subsets = get_sub_set(project_code_list)
#                 for subset in project_code_subsets:
#                     if ((account_cd == '101245') & (len(subset) == len(project_code_list))) or (account_cd != '101245'):
#                         subset_value_map = key_to_value(subset, map_sum_byProject)
#                         mapped_commercial1 = False
#                         if account_cd == '101245':
#                             # list_bankCharge = [0, 10, 20, 25, 30, 35, 40, 50, 60, 70, 80]
#                             # if sum(subset_value_map) - df_map['Actual Receipt  Amount'].sum() in list_bankCharge:
#                                 mapped_commercial1 = True
#                         else:
#                             if (sum(subset_value_map) - bank_value <= 0.03) & (sum(subset_value_map) - bank_value >= -0.03):
#                                 mapped_commercial1 = True
#                         # 如果mapping表中receipt date匹上的project code的subset值的汇总和银行匹配上
#                         if mapped_commercial1:
#                             glIndex_mappedToInd = []
#                             # 对加总值匹上的project code进行循环
#                             for project_id in subset:
#                                 df_map_grouped = df_map.groupby(['Notification Email', 'Project ID'])
#                                 for filter_condition, df in df_map_grouped:
#                                     if filter_condition[1] == project_id:
#                                         map_clear_date = df.iloc[0]['Notification Email']
#                                         # 获得某一入账时间的project code对应的mapping表总和，该值与GL的子集进行比对
#                                         sum_value_map = df['AR in Office Currency'].sum()
#                                         # print('sum_value_map', sum_value_map)
#                                         # 筛出还未mapping过的glData
#                                         glData_commercial_filtered = glData_commercial.loc[
#                                             glData_commercial.index.difference(mapped_glIndex_commercial_1)]
#                                         # 对还未mapping上的glData用入账时间和project code进行初步筛选
#                                         glData_commercial_filtered = glData_commercial_filtered[(glData_commercial_filtered[
#                                                                                                      'JH Created Date'] < map_clear_date + dt.timedelta(
#                                             days=8)) & (glData_commercial_filtered[
#                                                             'JH Created Date'] > map_clear_date - dt.timedelta(days=8)) & (
#                                                                                                         glData_commercial_filtered[
#                                                                                                             'Project Id'] == f'{project_id}')]
#                                         value_list_gl = glData_commercial_filtered['Amount Func Cur'].to_dict()
#                                         # print('value_list_gl', value_list_gl)
#                                         index_list_gl = list(value_list_gl.keys())
#                                         # print('index_list_gl', index_list_gl)
#                                         subsets_index_gl = get_sub_set(index_list_gl)
#                                         # print(subsets_index_gl)
#                                         for subset_index_gl in subsets_index_gl:
#                                             subset_value_gl = key_to_value(subset_index_gl, value_list_gl)
#                                             # 若筛出的glData的某个子集之和等于某一入账时间的project code对应的mapping表总和
#                                             if sum(subset_value_gl) == sum_value_map:
#                                                 # print(f'{subset_index_gl}', 'mapped')
#                                                 # id_number = id_number+1
#                                                 for index in subset_index_gl:
#                                                     if (index in glIndex_mappedToInd):
#                                                         # print(f'{index} previously mapped')
#                                                         # print('mapped_glIndex_commercial', mapped_glIndex_commercial_1)
#                                                         break
#                                                     else:
#                                                         # print('recorded index:', index)
#                                                         glIndex_mappedToInd.append(index)
#                                                         break
#
#                         # print('glIndex_mappedToInd', glIndex_mappedToInd)
#                         glData_sum_mappedToInd = glData_commercial.loc[glIndex_mappedToInd]['Amount Func Cur'].sum()
#                         check = glData_sum_mappedToInd - bank_value
#                         check_successful = False
#                         if account_cd == '101245':
#                             if df_map['Actual Receipt  Amount'].sum() == bank_value:
#                                 check_successful = True
#                         else:
#                             if (check <= 0.03) & (check >= -0.03):
#                                 check_successful = True
#                         if check_successful:
#                             if (ind in mapped_bankIndex_commercial_1) or common_data(glIndex_mappedToInd,
#                                                                                      mapped_glIndex_commercial_1):
#                                 pass
#                             else:
#                                 id_number_commercial = id_number_commercial + 1
#                                 # print('id_number', id_number_commercial)
#                                 bankData.loc[ind, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
#                                 glData.loc[
#                                     glIndex_mappedToInd, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
#                                 mapped_bankIndex_commercial_1.append(ind)
#                                 mapped_glIndex_commercial_1 = mapped_glIndex_commercial_1 + glIndex_mappedToInd
#
#     # #获取第一轮mapping之后剩余部分的glData和bankData
#     bankIndex_commercial_left = list(
#         set(list(bankData_commercial.index)).difference(set(mapped_bankIndex_commercial_1)))
#     glIndex_commercial_left = list(set(list(glData_commercial.index)).difference(set(mapped_glIndex_commercial_1)))
#
#     glData_commercial_left = glData_commercial.loc[glIndex_commercial_left, :]
#     bankData_commercial_left = bankData_commercial.loc[bankIndex_commercial_left, :]
#     # bankData_commercial_left['Narrative'] = bankData_commercial_left['Narrative'].map(lambda x: ''.join(line.strip() for line in x.splitlines()))
#     glData_commercial_left = glData_commercial_left.groupby('Vendor Name')
#
#     mapped_glIndex_commercial_2 = []
#     mapped_bankIndex_commercial_2 = []
#
#     # 第二轮commercial mapping
#     for client, df_left in glData_commercial_left:
#         sum_gl = df_left['Amount Func Cur'].sum()
#         bankAccountName_client = map_commercial.loc[
#             map_commercial['Client Name'] == f'{client}'.upper(), 'Client Name in Chinese']
#         if len(set(bankAccountName_client.to_list())):
#             for name in set(bankAccountName_client.to_list()):
#                 pro_name = name.strip()
#                 tf = bankData_commercial_left["Narrative"].str.contains(f'{pro_name}', regex=False, case=False)
#                 bkValue_list_commercial = bankData_commercial_left.loc[tf, 'Credit/Debit amount'].to_dict()
#                 subsets_bkIndex_commercial = get_sub_set(bkValue_list_commercial.keys())
#                 for subset_bkIndex_commercial in subsets_bkIndex_commercial:
#                     subset_value_bk = key_to_value(subset_bkIndex_commercial, bkValue_list_commercial)
#                     mapped_commercial2 = False
#                     if account_cd == '101245':
#                         list_bankCharge = [0, 10, 20, 25, 30, 35, 40, 50, 60, 70, 80]
#                         if (sum(subset_value_bk) - sum_gl) in list_bankCharge:
#                             mapped_commercial2 = True
#                     else:
#                         if ((sum(subset_value_bk) - sum_gl) <= 0.03) & ((sum(subset_value_bk) - sum_gl) >= -0.03):
#                             mapped_commercial2 = True
#                     if mapped_commercial2:
#                         if common_data(subset_bkIndex_commercial, mapped_bankIndex_commercial_2) or (
#                         common_data(list(df_left.index), mapped_glIndex_commercial_2)):
#                             pass
#                         else:
#                             id_number_commercial = id_number_commercial + 1
#                             bankData.loc[
#                                 subset_bkIndex_commercial, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
#                             glData.loc[df_left.index, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
#                             mapped_glIndex_commercial_2 = mapped_bankIndex_commercial_2 + list(df_left.index)
#                             mapped_bankIndex_commercial_2 = mapped_bankIndex_commercial_2 + subset_bkIndex_commercial
#
#     mapped_glIndex_commercial = mapped_glIndex_commercial_1 + mapped_glIndex_commercial_2
#     mapped_bankIndex_commercial = mapped_bankIndex_commercial_1 + mapped_bankIndex_commercial_2
#
#     return mapped_bankIndex_commercial, mapped_glIndex_commercial

def AP_mapping(bankData_AP, bankData, glData_vendor, glData, map_vendor_byEntity, account_cd, list_bankCharge, nameNum, vendorStaff):
    print(f'{vendorStaff} {nameNum} mapping')
    #修改点 编号不连续
    id_number_AP = 0
    mapped_glIndex_vendor = []
    mapped_bankIndex_vendor = []

    # AP Mapping 1 GL总和匹BK子集
    print('1 GL总和匹BK子集')
    for vendor, df in glData_vendor.groupby('Vendor Name'):
        glSum_vendor = df['Amount Func Cur'].sum()
        bankAccountNum_vendor = map_vendor_byEntity.loc[map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), f'Bank Account {nameNum}']

        # 当一个vendor name匹配多个银行账号时
        if len(set(bankAccountNum_vendor.to_list())):
            bkIndex_vendor = []
            dic_bkValue_AP = {}
            for num in set(bankAccountNum_vendor.to_list()):
                pro_num = str(num).strip()
                bkValue_list_AP = bankData_AP.loc[bankData_AP["Narrative"].str.contains(f'{pro_num}', regex=False, case=False), 'Credit/Debit amount'].to_dict()
                bkIndex_vendor = bkIndex_vendor + list(bkValue_list_AP.keys())
                dic_bkValue_AP.update(bkValue_list_AP)
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
    print('2 bk总和匹gl子集')
    for vendor, df in glData_vendor_left1.groupby('Vendor Name'):
        bankAccountName_vendor = map_vendor_byEntity.loc[map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), f'Bank Account {nameNum}']
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
    print('3 gl子集和bk子集匹配')
    for vendor, df in glData_vendor_left2.groupby('Vendor Name'):
        bankAccountName_vendor_2 = map_vendor_byEntity.loc[
            map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), f'Bank Account {nameNum}']
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
                        pass
                    else:
                        id_number_AP = id_number_AP + 1
                        bankData.loc[subset_bkIndex_vendor_3, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                        glData.loc[subset_glIndex_vendor_2, 'notes'] = f'AP {vendorStaff} netoff {now} {nameNum} {id_number_AP}'
                        mapped_glIndex_vendor = mapped_glIndex_vendor + list(subset_glIndex_vendor_2)
                        mapped_bankIndex_vendor = mapped_bankIndex_vendor + list(subset_bkIndex_vendor_3)

    return mapped_bankIndex_vendor, mapped_glIndex_vendor

def reimbursement_mapping(bankData_potentialTS, bankData_TSBatch, bankData, glData_reimbursement, glData, df_reimPay, account_number, month_period):
    # 设置初始值
    id_number_reim = 0
    mapped_glIndex_reim = []
    mapped_bankIndex_reim = []
    # 获取本entity下的报销mapping表
    entity = accountNo_to_entity[f'{account_number}']
    df_reimPay_filtered = df_reimPay[df_reimPay['Entity'].str.contains(f'{entity}')]
    print('employee mapping表信息和GL比对')
    # 将报销mapping表中的信息按月和GL匹配
    for month, df_reimPay_perM in df_reimPay_filtered.groupby('Month'):
        # 跳过不在month period里的月份
        if month not in month_period:
            continue
        list_staffName = set(df_reimPay_perM['Staff Name'].to_list())
        gl_mapped = []
        gl_mapped_index = []
        for staff in list_staffName:
            # 获取该月份mapping表中每一位员工的payment amount总和
            sum_pay = df_reimPay_perM.loc[df_reimPay_perM['Staff Name'] == f'{staff}', 'Payment Amount'].sum()
            # 将reimbursement payment info与gl进行比对
            gl_perStaff_mapped = False
            # 按月份和员工名筛出GL里的金额，形成带有index和payment amount的字典
            glData_reim_staffperM = glData_reimbursement.loc[(glData_reimbursement['Staff Name'] == f'{staff}') & (
                glData_reimbursement['JE Headers Description'].str.contains(f'{month}', case=False))]
            glValue_list_reim = glData_reim_staffperM['Amount Func Cur'].to_dict()
            # glValue_list_reim = glData_reimbursement.loc[(glData_reimbursement['Staff Name'] == f'{staff}') & (glData_reimbursement['JE Headers Description'].str.contains(f'{month}', case=False)), 'Amount Func Cur'].to_dict()
            sum_gl_reim = sum(glValue_list_reim.values())
            # 判断是否匹配上
            if sum_pay + sum_gl_reim < 0.02 and sum_pay + sum_gl_reim > -0.02:
                if common_data(mapped_glIndex_reim, glValue_list_reim.keys()):
                    pass
                else:
                    gl_perStaff_mapped = True
                    gl_staffMapped_index = list(glValue_list_reim.keys())
                    gl_mapped_index = gl_mapped_index + gl_staffMapped_index

            else:
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
        print('employee mapping表信息和Bank比对')
        # 将reimbursement payment info与bk进行比对
        if account_number == '001-221076-031':
            list_bankCharge = [0, -10, -20, -30, -40]
            bk_mapped = []
            bk_mapped_index = []
            valueMappedIndex_to_PIR = {}
            exactMappedIndex_to_PIR = {}
            for staff in list_staffName:
                list_staffPay = df_reimPay_perM.loc[
                    df_reimPay_perM['Staff Name'] == f'{staff}', 'Payment Amount'].to_list()
                bk_perStaff_mapped = False
                for payment_amount in list_staffPay:
                    mappedIndex_number = 0
                    mappedIndex = []
                    for bkIndex, bkValue in bankData_potentialTS['Credit/Debit amount'].to_dict().items():
                        if bkIndex in bk_mapped_index:
                            continue
                        if (payment_amount + bkValue) in list_bankCharge:
                            mappedIndex_number += 1
                            mappedIndex.append(bkIndex)
                    if mappedIndex_number == 1:
                        bk_perStaff_mapped = True
                        bk_mapped_index = bk_mapped_index + mappedIndex
                        bkMappedIndex_to_PIR = dict.fromkeys(mappedIndex, f'{staff}')
                        exactMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                    if mappedIndex_number > 1:
                        bk_perStaff_mapped = True
                        bkMappedIndex_to_PIR = dict.fromkeys(mappedIndex, f'{staff}')
                        valueMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                bk_mapped.append(bk_perStaff_mapped)

        if account_number != '001-221076-031':
            list_PIRnumber = set(df_reimPay_perM['PIR Number'].to_list())
            bk_mapped = []
            bk_mapped_index = []
            valueMappedIndex_to_PIR = {}
            exactMappedIndex_to_PIR = {}
            for number_PIR in list_PIRnumber:
                sum_pay_perPIR = df_reimPay_perM.loc[
                    df_reimPay_perM['PIR Number'] == f'{number_PIR}', 'Payment Amount'].sum()
                bkValue_list_reim = bankData_potentialTS.loc[
                    bankData_potentialTS['Credit/Debit amount'] == round(-sum_pay_perPIR,
                                                                         2), 'Credit/Debit amount'].to_dict()
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
                if len(bkValue_list_reim) >= 2:
                    index_in_bkTSBatch = set(bkValue_list_reim.keys()).intersection(
                        set(bankData_TSBatch.index.tolist()))
                    if len(index_in_bkTSBatch) == 1:
                        bk_perPIR_mapped = True
                        bkMappedIndex_to_PIR = dict.fromkeys(list(bkValue_list_reim.keys()), f'{number_PIR}')
                        exactMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                        bk_mapped_index = bk_mapped_index + index_in_bkTSBatch
                    else:
                        bk_perPIR_mapped = True
                        bkMappedIndex_to_PIR = dict.fromkeys(list(bkValue_list_reim.keys()), f'{number_PIR}')
                        valueMappedIndex_to_PIR.update(bkMappedIndex_to_PIR)
                        # bk_valueMapped_index = bk_valueMapped_index + list(bkValue_list_reim.keys())
                bk_mapped.append(bk_perPIR_mapped)
        if (False not in bk_mapped) and (False not in gl_mapped):
            id_number_reim += 1
            for key in exactMappedIndex_to_PIR.keys():
                bankData.loc[
                    key, 'notes'] = f'reimbursement netoff {now} {month} {id_number_reim} {exactMappedIndex_to_PIR[key]}'
            for key in valueMappedIndex_to_PIR.keys():
                bankData.loc[
                    key, 'notes'] = f'reimbursement netoff {now} {month} {id_number_reim} - value map {valueMappedIndex_to_PIR[key]}'
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

def payroll_mapping(bankData_left, bankData_SBID, bankData_TSBatch, bankData, glData_left, glData, account_cd):
    #设置初始值
    id_number_payroll = 0
    mapped_glIndex_payroll = []
    mapped_bankIndex_payroll = []
    if account_cd == '101245':
        dict_bk_payroll = bankData_left.loc[bankData_left['Narrative'].str.contains('Payroll', case=False, na=False), 'Credit/Debit amount'].to_dict()
    else:
        dict_bk_payroll = bankData_SBID.loc[bankData_SBID.index.difference(bankData_TSBatch.index), 'Credit/Debit amount'].to_dict()
    dict_gl_payroll = glData_left.loc[glData_left['Category Name'].str.contains('Payroll', case=False, na=False), 'Amount Func Cur'].to_dict()
    print('dict_bk_payroll', dict_bk_payroll)
    print('dict_gl_payroll', dict_gl_payroll)
    subsets_bkIndex_payroll = get_sub_set(dict_bk_payroll.keys())
    subsets_glIndex_payroll = get_sub_set(dict_gl_payroll.keys())
    for subset_bkIndex_payroll in subsets_bkIndex_payroll:
        subset_bkValue_payroll = key_to_value(subset_bkIndex_payroll, dict_bk_payroll)
        for subset_glIndex_payroll in subsets_glIndex_payroll:
            subset_glValue_payroll = key_to_value(subset_glIndex_payroll, dict_gl_payroll)
            if sum(subset_bkValue_payroll) == sum(subset_glValue_payroll):
                if common_data(subset_glIndex_payroll, mapped_glIndex_payroll) or common_data(subset_bkIndex_payroll, mapped_bankIndex_payroll):
                    pass
                else:
                    id_number_payroll = id_number_payroll + 1
                    bankData.loc[subset_bkIndex_payroll, 'notes'] = f'payroll netoff {now} {id_number_payroll}'
                    glData.loc[subset_glIndex_payroll, 'notes'] = f'payroll netoff {now} {id_number_payroll}'
                    mapped_glIndex_payroll = mapped_glIndex_payroll + subset_glIndex_payroll
                    mapped_bankIndex_payroll = mapped_bankIndex_payroll + subset_bkIndex_payroll
                    break
    return mapped_bankIndex_payroll, mapped_glIndex_payroll
#
def cashSettlement_mapping(bankData_left, bankData, glData_left, glData, account_cd):
    # 设置初始值
    id_number_settlement = 0
    mapped_glIndex_settlement = []
    mapped_bankIndex_settlement = []
    if account_cd == '101245':
        dict_bk_settlement = bankData_left.loc[bankData_left['Narrative'].str.contains('THE BOSTON CONSULTING GROUP UK', case=False, na=False), 'Credit/Debit amount'].to_dict()
    else:
        dict_bk_settlement = bankData_left.loc[bankData_left['Narrative'].str.contains('THE BOSTON CONSULTING GROUP, INC.', case=False, na=False), 'Credit/Debit amount'].to_dict()
    dict_gl_settlement = glData_left.loc[glData_left['Category Name'].str.contains('Cash Settlements', case=False, na=False), 'Amount Func Cur'].to_dict()
    print('dict_bk_settlement', dict_bk_settlement)
    print('dict_gl_settlement', dict_gl_settlement)
    subsets_bkIndex_settlement = get_sub_set(dict_bk_settlement.keys())
    subsets_glIndex_settlement = get_sub_set(dict_gl_settlement.keys())
    for subset_bkIndex_settlement in subsets_bkIndex_settlement:
        subset_bkValue_settlement = key_to_value(subset_bkIndex_settlement, dict_bk_settlement)
        for subset_glIndex_settlement in subsets_glIndex_settlement:
            subset_glValue_settlement = key_to_value(subset_glIndex_settlement, dict_gl_settlement)
            if sum(subset_bkValue_settlement) == sum(subset_glValue_settlement):
                if common_data(subset_glIndex_settlement, mapped_glIndex_settlement) or common_data(subset_bkIndex_settlement, mapped_bankIndex_settlement):
                    pass
                else:
                    id_number_settlement += 1
                    bankData.loc[subset_bkIndex_settlement, 'notes'] = f'settlement netoff {now} {id_number_settlement}'
                    glData.loc[subset_glIndex_settlement, 'notes'] = f'settlement netoff {now} {id_number_settlement}'
                    mapped_glIndex_settlement = mapped_glIndex_settlement + subset_glIndex_settlement
                    mapped_bankIndex_settlement = mapped_bankIndex_settlement + subset_bkIndex_settlement
                    break
    return mapped_bankIndex_settlement, mapped_glIndex_settlement

# path_folder_BS = input("Please enter the folder directory of all the BS statements:")
# path_folder_GL = input("Please enter the folder directory of all the GL files:")
# path_folder_reimRegister = input("Please enter the folder directory of all reimbursement files:")
# directory_AP_Vendor = input("Please enter the file link of the AP_Vendor Mapping:")
# directory_AP_Employee = input("Please enter the file link of the AP_Employee Mapping:")
# directory_Commercial = input("Please enter the file link of the Commercial Mapping:")
# month_period = input("Please enter the covered month periods:")
# path_folder_target = input("Please enter the folder directory where you want to store the results:")

def text_to_df(path_text):

    gl = open(rf'{path_text}', "r")
    lines = gl.readlines()
    period = []
    entity = []
    ru = []
    rc = []
    account = []
    location = []
    source = []
    dt_entry = []
    dt_post = []
    reference = []
    vendor_client = []
    num_transaction = []
    description = []
    amount = []
    for line in lines[40:]:
        period.append(line[0:6])
        entity.append(line[7:11])
        ru.append(line[12:15])
        rc.append(line[16:19])
        account.append(line[20:26])
        location.append(line[27:33])
        source.append(line[57:68])
        dt_entry.append(line[69:78])
        dt_post.append(line[79:88])
        reference.append(line[89:112])
        vendor_client.append(line[113:135])
        num_transaction.append(line[136:151])
        description.append(line[152:171])
        amount.append(line[173:193])


    df_gl = pd.DataFrame()
    df_gl['Period'] = period[2:]
    df_gl['Entity'] = entity[2:]
    df_gl['RU'] = ru[2:]
    df_gl['RC'] = rc[2:]
    df_gl['Account'] = account[2:]
    df_gl['Location'] = location[2:]
    df_gl['Journal Source'] = source[2:]
    df_gl['Entry Date'] = dt_entry[2:]
    df_gl['Posted Date'] = dt_post[2:]
    df_gl['Reference'] = reference[2:]
    df_gl[r'Vendor/Client Name'] = vendor_client[2:]
    df_gl['Transaction Number'] = num_transaction[2:]
    df_gl['Description'] = description[2:]
    df_gl['Amount'] = amount[2:]

    df_gl['Amount'] = df_gl['Amount'].map(lambda x: x.strip().replace(',', '').replace('(', '-').replace(')', ''))
    df_gl['Amount'] = pd.to_numeric(df_gl['Amount'], errors='coerce')
    df_gl.dropna(subset=['Amount'], how='all', inplace=True)
    df_gl=df_gl.reset_index()

    reference_blank = df_gl[df_gl['Reference'].str.isspace()]
    for ind, row in reference_blank.iterrows():
        if row['Posted Date'].isspace():
            df_gl.loc[ind, 'Reference'] = df_gl.loc[ind-1, 'Reference']

    period_blank = df_gl[df_gl['Period'].str.isspace()]
    for ind, row in period_blank.iterrows():
        if row['Period'].isspace():
            column_names = ['Period', 'Entity', 'RU', 'RC', 'Account', 'Location', 'Journal Source', 'Entry Date', 'Posted Date']
            for column in column_names:
                df_gl.loc[ind, f'{column}'] = df_gl.loc[ind-1, f'{column}']

    return df_gl

path_folder_BS = r'C:\Users\he kelly\Desktop\bank_reconciliation_py\Bank Rec Program\Bank Statement'
path_folder_GL = r'C:\Users\he kelly\Desktop\bank_reconciliation_py\Bank Rec Program\GL'
path_folder_reimRegister = r'C:\Users\he kelly\Desktop\bank_reconciliation_py\reimbursement\New folder'
directory_AP_Vendor = r'C:\Users\he kelly\Desktop\bank_reconciliation_py\Bank Rec Program\Mapping\AP Mapping.xlsx'
directory_AP_Employee = r'C:\Users\he kelly\Desktop\bank_reconciliation_py\Bank Rec Program\Mapping\Employee mapping.xlsx'
directory_Commercial = r'C:\Users\he kelly\Desktop\bank_reconciliation_py\Bank Rec Program\Mapping\Cash receipt 2023.xlsx'
month_period = 'JAN FEB MAR'
path_folder_target = r'C:\Users\he kelly\Desktop\bank_reconciliation_py\Bank Rec Program\test'

now = str(datetime.now()).split('.')[0]

print('获取所有bank信息ing')
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


print('获取所有GL信息ing')
#获取所有GL信息
files_GL = os.listdir(rf'{path_folder_GL}')
df_GL = pd.DataFrame()
df_GL_HK = pd.DataFrame()
for file_GL in files_GL:
    if file_GL.endswith('txt'):
        file_path_GL = os.path.join(path_folder_GL, file_GL)
        df_file_GL = text_to_df(file_path_GL)
        df_GL_HK = pd.concat([df_GL_HK, df_file_GL])
    else:
        file_path_GL = os.path.join(path_folder_GL, file_GL)
        df_file_GL = pd.read_excel(file_path_GL, header=1).reset_index()
        df_GL = pd.concat([df_GL, df_file_GL])

df_GL_HK['Posted Date'] = pd.to_datetime(df_GL_HK['Posted Date'])

print('获取vendor mapping中')
#获取vendor mapping
map_vendor = pd.read_excel(rf'{directory_AP_Vendor}', header=0)
map_vendor['Vendor Name'] = map_vendor['Vendor Name'].map(lambda x: x.upper())
map_employee = pd.read_excel(directory_AP_Employee, header=1)
accountNo_to_vendorSite = {'626-055784-001': 'Beijing OU', '622-512317-001': 'Shenzhen OU', '088-169370-011': 'China PRC OU', '001-221076-031': 'Taiwan OU'}

print('获取employee mapping中')
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

print('获取commercial mapping中')
#读取Commercial mapping, 创建mapping dictionary
map_commercial = pd.read_excel(rf'{directory_Commercial}', header=0)
map_commercial['Actual Receipt  Amount'].fillna(method='ffill', axis=0, inplace=True)
map_commercial['Receipt Dt'] = map_commercial['Receipt Dt'].astype('datetime64[ns]')
map_commercial['bank expense'] = map_commercial['bank expense'].astype('float')
map_commercial['Client Name'] = map_commercial['Client Name'].dropna().map(lambda x: x.upper())
tb_location = {'088-169370-011': 'PRC', '626-055784-001': 'Beijing', '622-512317-001': 'Shenzhen', '001-221076-031': 'Taipei', '500-422688-274': 'Hong Kong', '500-422688-001': 'Hong Kong', '500-422688-002': 'Hong Kong'}

list_bankCharge = [0, 10, 20, 25, 30, 35, 40, 50, 60, 70, 80]

# PRC Section
bank_mapping_PRC = {'088-169370-011': '101244', '626-055784-001': '101001', '622-512317-001': '101135', '001-221076-031': '101245'}
# bank_mapping_PRC = {'088-169370-011': '101244'}

for account_number, account_cd in bank_mapping_PRC.items():

    print('Start Mapping Account Code', account_cd)

    #获取当前bank account的bank和gl数据
    bankData = df_bank[df_bank['Account number'] == f'{account_number}']
    glData = df_GL[df_GL['Account Cd'] == int(account_cd)]

    print('处理无需mapping的bank data')

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
        bankData.loc[sweep_netoff, 'notes'] = 'sweep'
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

    print('Commercial Mapping')

    #commercial mapping
    bankData_commercial = bankData_filtered.loc[bankData_filtered['Credit/Debit amount']>0, :] #排除bk金额小于等于0.03的item
    glData_commercial = glData[glData['JE Headers Description'].str.contains('Cash Receipts')]

    if account_cd == '101245':
        mapped_bankIndex_commercial, mapped_glIndex_commercial = commercial_mapping_TW(bankData_commercial, bankData, glData_commercial, glData, map_commercial)
    else:
        mapped_bankIndex_commercial, mapped_glIndex_commercial = commercial_mapping(bankData_commercial, bankData, glData_commercial, glData, map_commercial, account_cd)

    print('first repayment mapping')
    #AP退款重付
    bankIndex_repayment_netoff = repayment_mapping(bankData_filtered, bankData, 'first', account_cd, list_bankCharge)


    #AP Mapping
    #挖掉已匹上的commercial部分
    bankData_AP = bankData_filtered.loc[bankData_filtered.index.difference(mapped_bankIndex_commercial)]
    #挖掉退款重复的部分
    bankData_AP = bankData_AP.loc[bankData_AP.index.difference(bankIndex_repayment_netoff)]
    glData_AP = glData.loc[glData['View Source'].str.contains('Payables', regex=False, case=False, na=False)]
    glData_staff = glData[glData['Vendor Name'].str.contains('                              ', regex=False, case=False, na=False)]
    glData_reimbursement = pd.DataFrame()
    staff_invoice_indication = ['HLYERR', 'TB', 'RVCR', 'CM', 'CR']
    for item in staff_invoice_indication:
        df_staff = glData_AP[glData_AP['Invoice Number'].str.contains(f'{item}', regex=False, case=False, na=False)]
        glData_reimbursement = pd.concat([glData_reimbursement, df_staff])
    glData_reimbursement['Staff Name'] = glData_reimbursement['Vendor Name'].map(lambda x: x.split('      ')[0])
    glData_staff['Staff Name'] = glData_staff['Vendor Name'].map(lambda x: x.split('      ')[0])
    glData_vendor = glData_AP.loc[glData_AP.index.difference(glData_staff.index)]

    print('AP vendor mapping')
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

    print('reimbursement mapping')
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


    print('AP Staff mapping')
    # Staff Mapping
    # 在bkData potentialTS里挖去已匹配的报销部分
    bankData_potentialStaff = bankData_potentialTS.loc[bankData_potentialTS.index.difference(mapped_bankIndex_reim)]
    # 在glData Staff里挖去已匹配的报销部分
    glData_staff = glData_staff.loc[glData_staff.index.difference(mapped_glIndex_reim)]
    mapped_bankIndex_staff1, mapped_glIndex_staff1 = AP_mapping(bankData_potentialStaff, bankData, glData_staff, glData, map_employee_byEntity, account_cd, list_bankCharge, 'Name', 'staff')

    bankData_potentialStaff_left = bankData_potentialStaff.loc[bankData_potentialStaff.index.difference(mapped_bankIndex_staff1)]
    glData_staff_left = glData_staff.loc[glData_staff.index.difference(mapped_glIndex_staff1)]

    mapped_bankIndex_staff2, mapped_glIndex_staff2 = AP_mapping(bankData_potentialStaff_left, bankData, glData_staff_left, glData, map_employee_byEntity, account_cd, list_bankCharge, 'Num', 'staff')

    mapped_bankIndex_staff = mapped_bankIndex_staff1 + mapped_bankIndex_staff2

    bankData_AP_left4 = bankData_AP_left3.loc[bankData_AP_left3.index.difference(mapped_bankIndex_reim)]
    bankData_AP_left4 = bankData_AP_left4.loc[bankData_AP_left4.index.difference(mapped_bankIndex_staff)]
    # excel_log.log(bankData_AP_left4, 'bankData_left')

    print('last repayment mapping')

    bankIndex_repayment_netoff2 = repayment_mapping(bankData_AP_left4, bankData, 'last', account_cd, list_bankCharge)

    bankData_left2 = bankData_AP_left4.loc[bankData_AP_left4.index.difference(bankIndex_repayment_netoff2)]

    print('payroll mapping')
    mapped_bankIndex_payroll, mapped_glIndex_payroll = payroll_mapping(bankData_left2, bankData_SBID, bankData_TSBatch, bankData, glData, glData, account_cd)

    print('cash settlement mapping')
    mapped_bankIndex_settlement, mapped_glIndex_settlement = cashSettlement_mapping(bankData_left2, bankData, glData, glData, account_cd)


    now_for_folder = now.replace(':', ' ')
    os.makedirs(rf'{path_folder_target}\{now_for_folder}\{location}')
    bankData.to_excel(fr'{path_folder_target}\{now_for_folder}\{location}\bank_{location}_{account_number}.xlsx')
    glData.to_excel(fr'{path_folder_target}\{now_for_folder}\{location}\gl_{location}_{account_cd}.xlsx')
    df_reimPay.to_excel(fr'{path_folder_target}\{now_for_folder}\reimbursement summary.xlsx')

#HK Section
bank_mapping_HK = {'500-422688-001': '101102', '500-422688-274': '101113', '500-422688-002': '101130'}
# bank_mapping_HK = {'500-422688-002': '101130'}
#
# for account_number, account_cd in bank_mapping_HK.items():
#
#     print('Start Mapping Account Code', account_cd)
#
#     #获取当前bank account的bank和gl数据
#     bankData = df_bank[df_bank['Account number'] == f'{account_number}']
#     glData = df_GL_HK[df_GL_HK['Account'] == f'{account_cd}']
#
#     print('处理无需mapping的bank data')
#     bankData_charges = bankData[(bankData['TRN type'] == 'Charges')|(bankData['Customer reference'].str.contains('MT940 MONTHLY CH'))|(bankData['Narrative'].str.contains('HSBCNET MONTHLY FEE'))]
#     bankData.loc[bankData_charges.index, 'notes'] = 'bank charges'
#     # bankData.loc[bankData['Customer reference'].str.contains('MT940 MONTHLY CH'), 'notes'] = 'bank charges'
#     # bankData.loc[bankData['Narrative'].str.contains('HSBCNET MONTHLY FEE'), 'notes'] = 'bank charges'
#     bankData_interest = bankData[bankData['TRN type'] == 'Interest']
#     bankData.loc[bankData_interest.index, 'notes'] = 'bank interest'
#     index_filtered = list(set(bankData.index).difference(set(list(bankData_charges.index)+list(bankData_interest.index)))) #改名字 index_emptyNotes
#     bankData_filtered = bankData.iloc[index_filtered] #改名字
#
#     location = tb_location[bankData.iloc[1]['Account number']]
#
# #     #employee and vendor mapping
# #     #获取当前entity的employee and vendor mapping
# #     map_employee_byEntity = map_employee.loc[map_employee['Vendor Site OU'] == f'{accountNo_to_vendorSite[account_number]}']
# #     map_vendor_byEntity = map_vendor.loc[map_vendor['Vendor Site OU'] == f'{accountNo_to_vendorSite[account_number]}']
# #
#     print('Commercial Mapping')
#
#     #commercial mapping
#     bankData_commercial = bankData_filtered.loc[bankData_filtered['Credit/Debit amount']>0, :] #排除bk金额小于等于0.03的item
#     glData_commercial = glData[glData['Reference'].str.contains('Cash Receipts')]
#     mapped_bankIndex_commercial, mapped_glIndex_commercial = commercial_mapping_HK(bankData_commercial, bankData, glData_commercial, glData, map_commercial, account_cd)
#     print(mapped_bankIndex_commercial)
#     print(mapped_glIndex_commercial)
#
#     bankData_AP = bankData_filtered.loc[bankData_filtered.index.difference(mapped_bankIndex_commercial)]
#     glData_AP = glData[glData['Journal Source'].str.contains('Payables')]
#
#
#     id_number_reim = 0
#     mapped_bankIndex_reim = []
#     mapped_glIndex_reim = []
#     glData_reim = glData_AP[glData_AP['Reference'].str.contains('HK_')]
#     bankData_reim = bankData_AP[bankData_AP['TRN type'].str.contains('Bulk')]
#     gl_reim_batch = glData_reim.groupby('Reference').sum('Amount')['Amount'].to_dict()
#     bk_reim_batch = bankData_reim['Credit/Debit amount'].to_dict()
#     for batch, gl_amount in gl_reim_batch.items():
#         for index, bk_amount in bk_reim_batch.items():
#             if (gl_amount - bk_amount) < 0.1 and (gl_amount - bk_amount) > -0.1:
#                 id_number_reim += 1
#                 gl_index = glData_reim[glData_reim['Reference'].str.contains(f'{batch}')].index.tolist()
#                 mapped_glIndex_reim = mapped_glIndex_reim + gl_index
#                 mapped_bankIndex_reim.append(index)
#                 bankData.loc[index, 'notes'] = f'reimbursement netoff {now} {id_number_reim}'
#                 glData.loc[gl_index, 'notes'] = f'reimbursement netoff {now} {id_number_reim}'
#
#     id_number_AP = 0
#     mapped_glIndex_AP = []
#     mapped_bankIndex_AP = []
#     glData_AP_left = glData_AP.loc[glData_AP.index.difference(mapped_glIndex_reim)]
#     bankData_AP_left = bankData_AP.loc[bankData_AP.index.difference(mapped_bankIndex_reim)]
#     glData_AP_left['Transaction Number'] = glData_AP_left['Transaction Number'].map(lambda x: x.strip())
#     transactionNo_gl = glData_AP_left['Transaction Number'].to_dict()
#     transactionNo_bk = bankData_AP_left['Customer reference'].to_dict()
#     for ind_gl, number_gl in transactionNo_gl.items():
#         for ind_bk, number_bk in transactionNo_bk.items():
#             glValue_a = glData_AP_left.loc[ind_gl, 'Amount']
#             bkValue_a = bankData_AP_left.loc[ind_bk, 'Credit/Debit amount']
#             if (number_gl in number_bk) and glValue_a == bkValue_a:
#                 id_number_AP += 1
#                 mapped_glIndex_AP.append(ind_gl)
#                 mapped_bankIndex_AP.append(ind_bk)
#                 bankData.loc[ind_bk, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                 glData.loc[ind_gl, 'notes'] = f'AP netoff {now} {id_number_AP}'
#             if (number_gl in number_bk) and ('/EXCH/' in bankData_AP_left.loc[ind_bk, 'Narrative']) and (glValue_a - bkValue_a < 400) and (glValue_a - bkValue_a > -400):
#                 id_number_AP += 1
#                 mapped_glIndex_AP.append(ind_gl)
#                 mapped_bankIndex_AP.append(ind_bk)
#                 bankData.loc[ind_bk, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                 bankData.loc[ind_bk, 'FX'] = bkValue_a - glValue_a
#                 glData.loc[ind_gl, 'notes'] = f'AP netoff {now} {id_number_AP}'
#                 glData.loc[ind_gl, 'FX'] = bkValue_a - glValue_a
#
#
#
#     glData_AP_left = glData_AP_left.loc[glData_AP_left.index.difference(mapped_glIndex_AP)]
#     bankData_AP_left = bankData_AP_left.loc[bankData_AP_left.index.difference(mapped_bankIndex_AP)]
#
#     customer_reference = bankData_AP_left['Customer reference'].to_dict()
#     for ind, reference in customer_reference.items():
#         if not bool(re.search(r'\d', f'{reference}')):
#             dict_gl_AP = glData_AP_left.loc[glData_AP_left['Vendor/Client Name'].str.contains(f'{reference}', case=False), 'Amount'].to_dict()
#             if len(dict_gl_AP):
#                 subsets_glIndex_AP = get_sub_set(list(dict_gl_AP.keys()))
#                 subsets_glIndex_AP.reverse()
#                 subsets_glIndex_AP = [x for x in subsets_glIndex_AP if x != []]
#                 for subset_glIndex_AP in subsets_glIndex_AP:
#                     subset_glValue_AP = key_to_value(subset_glIndex_AP, dict_gl_AP)
#                     glValue_b = sum(subset_glValue_AP)
#                     bkValue_b = bankData_AP_left.loc[ind, 'Credit/Debit amount']
#                     if glValue_b == bkValue_b:
#                         if common_data(subset_glIndex_AP, mapped_glIndex_AP):
#                             pass
#                         else:
#                             id_number_AP += 1
#                             mapped_glIndex_AP = mapped_glIndex_AP + subset_glIndex_AP
#                             mapped_bankIndex_AP.append(ind)
#                             bankData.loc[ind, 'notes'] = f'AP1 netoff {now} {id_number_AP}'
#                             glData.loc[subset_glIndex_AP, 'notes'] = f'AP1 netoff {now} {id_number_AP}'
#                             break
#                     if ('/EXCH/' in bankData_AP_left.loc[ind, 'Narrative']) and (glValue_b - bkValue_b < 400) and (glValue_b - bkValue_b > -400):
#                         if common_data(subset_glIndex_AP, mapped_glIndex_AP):
#                             pass
#                         else:
#                             id_number_AP += 1
#                             mapped_glIndex_AP = mapped_glIndex_AP + subset_glIndex_AP
#                             mapped_bankIndex_AP.append(ind)
#                             bankData.loc[ind, 'notes'] = f'AP1 netoff {now} {id_number_AP}'
#                             bankData.loc[ind, 'FX'] = bkValue_b - glValue_b
#                             glData.loc[subset_glIndex_AP, 'notes'] = f'AP1 netoff {now} {id_number_AP}'
#                             glData.loc[subset_glIndex_AP, 'FX'] = (bkValue_b - glValue_b)/len(subset_glIndex_AP)
#                             break
#
#     glData_AP_left = glData_AP_left.loc[glData_AP_left.index.difference(mapped_glIndex_AP)]
#     bankData_AP_left = bankData_AP_left.loc[bankData_AP_left.index.difference(mapped_bankIndex_AP)]
#
#
#     glData_AP_left['Vendor/Client Name'] = glData_AP_left['Vendor/Client Name'].map(lambda x: x.strip())
#
#
#     #1bk vs multi gl
#     for client in set(glData_AP_left['Vendor/Client Name'].to_list()):
#         # if client != 'PRESIDENT & FELLOWS OF':
#         #     continue
#         if client.endswith('.') or client.endswith(','):
#             client_bk = client[:-1]
#         if '&' in client:
#             client_bk = client.translate(str.maketrans({'&': 'AND'}))
#         else:
#             client_bk = client
#         dict_bk_AP = bankData_AP_left.loc[bankData_AP_left['Narrative'].str.contains(f'{client_bk}', regex=False, case=False), 'Credit/Debit amount'].to_dict()
#         if len(dict_bk_AP):
#             dict_gl_AP1 = glData_AP_left.loc[glData_AP_left['Vendor/Client Name'].str.contains(f'{client}', regex=False, case=False), 'Amount'].to_dict()
#             subsets_glIndex_AP1 = get_sub_set(list(dict_gl_AP1.keys()))
#             subsets_glIndex_AP1.reverse()
#             subsets_glIndex_AP1 = [x for x in subsets_glIndex_AP1 if x != []]
#             for ind, bk_value in dict_bk_AP.items():
#                 for subset_glIndex_AP1 in subsets_glIndex_AP1:
#                     subset_glValue_AP1 = key_to_value(subset_glIndex_AP1, dict_gl_AP1)
#                     glValue_c = sum(subset_glValue_AP1)
#                     bkValue_c = bk_value
#                     if bkValue_c == glValue_c:
#                         if ind in mapped_bankIndex_AP or common_data(subset_glIndex_AP1, mapped_glIndex_AP):
#                             pass
#                         else:
#                             id_number_AP += 1
#                             mapped_glIndex_AP = mapped_glIndex_AP + subset_glIndex_AP1
#                             mapped_bankIndex_AP.append(ind)
#                             bankData.loc[ind, 'notes'] = f'AP2 netoff {now} {id_number_AP}'
#                             glData.loc[subset_glIndex_AP1, 'notes'] = f'AP2 netoff {now} {id_number_AP}'
#                             break
#                     if ('/EXCH/' in bankData_AP_left.loc[ind, 'Narrative']) and (bkValue_c - glValue_c < 400) and (bkValue_c - glValue_c > -400):
#                         if ind in mapped_bankIndex_AP or common_data(subset_glIndex_AP1, mapped_glIndex_AP):
#                             pass
#                         else:
#                             id_number_AP += 1
#                             mapped_glIndex_AP = mapped_glIndex_AP + subset_glIndex_AP1
#                             mapped_bankIndex_AP.append(ind)
#                             bankData.loc[ind, 'notes'] = f'AP2 netoff {now} {id_number_AP}'
#                             bankData.loc[ind, 'FX'] = bkValue_c - glValue_c
#                             glData.loc[subset_glIndex_AP1, 'notes'] = f'AP2 netoff {now} {id_number_AP}'
#                             glData.loc[subset_glIndex_AP1, 'FX'] = bkValue_c - glValue_c
#                             break
#
#
#     glData_AP_left = glData_AP_left.loc[glData_AP_left.index.difference(mapped_glIndex_AP)]
#     bankData_AP_left = bankData_AP_left.loc[bankData_AP_left.index.difference(mapped_bankIndex_AP)]
#
#     #multi bk vs multi gl
#     for client in set(glData_AP_left['Vendor/Client Name'].to_list()):
#         # if client != 'PRESIDENT & FELLOWS OF':
#         #     continue
#         if client.endswith('.') or client.endswith(','):
#             client_bk = client[:-1]
#         if '&' in client:
#             client_bk = client.translate(str.maketrans({'&': 'AND'}))
#         else:
#             client_bk = client
#         dict_bk_AP = bankData_AP_left.loc[bankData_AP_left['Narrative'].str.contains(f'{client_bk}', regex=False, case=False), 'Credit/Debit amount'].to_dict()
#         subsets_bkIndex_AP = get_sub_set(list(dict_bk_AP.keys()))
#         subsets_bkIndex_AP.reverse()
#         subsets_bkIndex_AP = [x for x in subsets_bkIndex_AP if x != []]
#         if len(dict_bk_AP):
#             dict_gl_AP1 = glData_AP_left.loc[glData_AP_left['Vendor/Client Name'].str.contains(f'{client}', regex=False, case=False), 'Amount'].to_dict()
#             subsets_glIndex_AP1 = get_sub_set(list(dict_gl_AP1.keys()))
#             subsets_glIndex_AP1.reverse()
#             subsets_glIndex_AP1 = [x for x in subsets_glIndex_AP1 if x != []]
#             for subset_bkIndex_AP in subsets_bkIndex_AP:
#                 subset_bkValue_AP = key_to_value(subset_bkIndex_AP, dict_bk_AP)
#                 for subset_glIndex_AP1 in subsets_glIndex_AP1:
#                     subset_glValue_AP1 = key_to_value(subset_glIndex_AP1, dict_gl_AP1)
#                     bkValue_d = sum(subset_bkValue_AP)
#                     glValue_d = sum(subset_glValue_AP1)
#                     if bkValue_d == glValue_d:
#                         if common_data(subset_bkIndex_AP, mapped_bankIndex_AP) or common_data(subset_glIndex_AP1, mapped_glIndex_AP):
#                             pass
#                         else:
#                             id_number_AP += 1
#                             mapped_glIndex_AP = mapped_glIndex_AP + subset_glIndex_AP1
#                             mapped_bankIndex_AP = mapped_bankIndex_AP + subset_bkIndex_AP
#                             bankData.loc[subset_bkIndex_AP, 'notes'] = f'AP3 netoff {now} {id_number_AP}'
#                             glData.loc[subset_glIndex_AP1, 'notes'] = f'AP3 netoff {now} {id_number_AP}'
#                             break
#                     if ('/EXCH/' in bankData_AP_left.loc[subset_bkIndex_AP[0], 'Narrative']) and (bkValue_d - glValue_d < 400) and (bkValue_d - glValue_d > -400):
#                         if common_data(subset_bkIndex_AP, mapped_bankIndex_AP) or common_data(subset_glIndex_AP1, mapped_glIndex_AP):
#                             pass
#                         else:
#                             id_number_AP += 1
#                             mapped_glIndex_AP = mapped_glIndex_AP + subset_glIndex_AP1
#                             mapped_bankIndex_AP= mapped_bankIndex_AP + subset_bkIndex_AP
#                             bankData.loc[subset_bkIndex_AP, 'notes'] = f'AP3 netoff {now} {id_number_AP}'
#                             bankData.loc[subset_bkIndex_AP, 'FX'] = (bkValue_d - glValue_d)/len(subset_bkIndex_AP)
#                             glData.loc[subset_glIndex_AP1, 'notes'] = f'AP3 netoff {now} {id_number_AP}'
#                             glData.loc[subset_glIndex_AP1, 'FX'] = (bkValue_d - glValue_d)/len(subset_glIndex_AP1)
#                             break
#
#     print(mapped_bankIndex_AP)
#
#     glData_AP_left = glData_AP_left.loc[glData_AP_left.index.difference(set(mapped_glIndex_AP))]
#     bankData_AP_left = bankData_AP_left.loc[bankData_AP_left.index.difference(set(mapped_bankIndex_AP))]
#     print(glData_AP_left)
#
#     id_number_fund = 0
#     mapped_bankIndex_fund = []
#     mapped_glIndex_fund = []
#
#     dict_bk_fund = bankData_AP_left.loc[bankData_AP_left['Narrative'].str.contains('500422688')|bankData_AP_left['Narrative'].str.contains('THE BOSTON CONSULTING GRP INTL'), 'Credit/Debit amount'].to_dict()
#     dict_gl_fund = glData.loc[glData['Journal Source'].str.contains('Spreadsheet', case=False) & glData['Reference'].str.contains('fund trans', case=False), 'Amount'].to_dict()
#     print(dict_bk_fund)
#     print(dict_gl_fund)
#     subsets_glIndex_fund = get_sub_set(dict_gl_fund.keys())
#     print(subsets_glIndex_fund)
#     subsets_bkIndex_fund = get_sub_set(dict_bk_fund.keys())
#     print(subsets_bkIndex_fund)
#     for subset_glIndex_fund in subsets_glIndex_fund:
#         subset_glValue_fund = key_to_value(subset_glIndex_fund, dict_gl_fund)
#         print(subset_glValue_fund)
#         for subset_bkIndex_fund in subsets_bkIndex_fund:
#             subset_bkValue_fund = key_to_value(subset_bkIndex_fund, dict_bk_fund)
#             print(subset_bkValue_fund)
#             if sum(subset_glValue_fund) == sum(subset_bkValue_fund):
#                 if common_data(subset_glIndex_fund, mapped_glIndex_fund) or common_data(subset_bkIndex_fund, mapped_bankIndex_fund):
#                     pass
#                 else:
#                     id_number_fund += 1
#                     mapped_bankIndex_fund = mapped_bankIndex_fund + subset_bkIndex_fund
#                     mapped_glIndex_fund = mapped_glIndex_fund + subset_glIndex_fund
#                     bankData.loc[subset_bkIndex_fund, 'notes'] = f'fund netoff {now} {id_number_fund}'
#                     glData.loc[subset_glIndex_fund, 'notes'] = f'fund netoff {now} {id_number_fund}'
#                     break
#
#
#     bankData_left = bankData_AP_left.loc[bankData_AP_left.index.difference(set(mapped_bankIndex_fund))]
#
#     id_number_payroll = 0
#     mapped_bankIndex_payroll = []
#     mapped_glIndex_payroll = []
#
#     dict_gl_payroll = glData.loc[glData['Journal Source'].str.contains('Spreadsheet', case=False) & glData['Reference'].str.contains('Payroll', case=False), 'Amount'].to_dict()
#     print('dict_gl_payroll', dict_gl_payroll)
#     dict_bk_payroll = bankData_left.loc[bankData_left['Narrative'].str.contains('salary', case=False), 'Credit/Debit amount'].to_dict()
#     print('dict_bk_payroll', dict_bk_payroll)
#     subsets_glIndex_payroll = get_sub_set(dict_gl_payroll.keys())
#     subsets_bkIndex_payroll = get_sub_set(dict_bk_payroll.keys())
#     for subset_glIndex_payroll in subsets_glIndex_payroll:
#         subset_glValue_payroll = key_to_value(subset_glIndex_payroll, dict_gl_payroll)
#         for subset_bkIndex_payroll in subsets_bkIndex_payroll:
#             subset_bkValue_payroll = key_to_value(subset_bkIndex_payroll, dict_bk_payroll)
#             if sum(subset_glValue_payroll) == sum(subset_bkValue_payroll):
#                 if common_data(subset_bkIndex_payroll, mapped_glIndex_payroll) or common_data(subset_bkIndex_payroll, mapped_bankIndex_payroll):
#                     pass
#                 else:
#                     id_number_payroll += 1
#                     mapped_bankIndex_payroll = mapped_bankIndex_payroll + subset_bkIndex_payroll
#                     mapped_glIndex_payroll = mapped_glIndex_payroll + subset_glIndex_payroll
#                     bankData.loc[subset_bkIndex_payroll, 'notes'] = f'payroll netoff {now} {id_number_payroll}'
#                     glData.loc[subset_glIndex_payroll, 'notes'] = f'payroll netoff {now} {id_number_payroll}'
#                     break
#
#
#
#     now_for_folder = now.replace(':', ' ')
#     os.makedirs(rf'{path_folder_target}\{now_for_folder}\{location}_{account_cd}')
#     bankData.to_excel(fr'{path_folder_target}\{now_for_folder}\{location}_{account_cd}\bank_{location}_{account_number}.xlsx')
#     glData.to_excel(fr'{path_folder_target}\{now_for_folder}\{location}_{account_cd}\gl_{location}_{account_cd}.xlsx')


#
# input('Finished:')












