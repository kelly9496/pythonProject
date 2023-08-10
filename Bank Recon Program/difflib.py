# input 1: bankData_commercial
# input 2: map_commercial_RPA
# input 3: account_cd
# input 4: glData_commercial
# input 5: bankData
# input 6: glData
# input 7: map_commercial
# input 8: list_bankCharge

def commercial_mapping(bankData_commercial, bankData, glData_commercial, glData, map_commercial, account_cd):

    location = tb_location[bankData_commercial.iloc[1]['Account number']]
    # 处理commercial mapping表，筛出本entity RPApo账的部分，并按project ID分类
    map_commercial_RPA = map_commercial[map_commercial['location'].str.contains(f'{location}') & (map_commercial['Currency'] == 'CNY') & (map_commercial['Notification Email'] != '-')]
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
                            if (ind in mapped_bankIndex_commercial_1) or common_data(glIndex_mappedToInd, mapped_glIndex_commercial_1):
                                pass
                            else:
                                id_number_commercial = id_number_commercial + 1
                                # print('id_number', id_number_commercial)
                                bankData.loc[ind, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                                glData.loc[glIndex_mappedToInd, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                                mapped_bankIndex_commercial_1.append(ind)
                                mapped_glIndex_commercial_1 = mapped_glIndex_commercial_1 + glIndex_mappedToInd

    # #获取第一轮mapping之后剩余部分的glData和bankData
    bankIndex_commercial_left = list(set(list(bankData_commercial.index)).difference(set(mapped_bankIndex_commercial_1)))
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
        bankAccountName_client = map_commercial.loc[map_commercial['Client Name'] == f'{client}'.upper(), 'Client Name in Chinese']
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
                        if common_data(subset_bkIndex_commercial, mapped_bankIndex_commercial_2) or (common_data(list(df_left.index), mapped_glIndex_commercial_2)):
                            pass
                        else:
                            id_number_commercial = id_number_commercial + 1
                            bankData.loc[subset_bkIndex_commercial, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                            glData.loc[df_left.index, 'notes'] = f'commercial netoff {now} {id_number_commercial}'
                            mapped_glIndex_commercial_2 = mapped_bankIndex_commercial_2 + list(df_left.index)
                            mapped_bankIndex_commercial_2 = mapped_bankIndex_commercial_2 + subset_bkIndex_commercial

    mapped_glIndex_commercial = mapped_glIndex_commercial_1 + mapped_glIndex_commercial_2
    mapped_bankIndex_commercial = mapped_bankIndex_commercial_1 + mapped_bankIndex_commercial_2

    return mapped_bankIndex_commercial, mapped_glIndex_commercial
