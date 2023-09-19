def AP_mapping(bankData_AP, bankData, glData_vendor, glData, map_vendor_byEntity, account_cd, list_bankCharge, nameNum, vendorStaff):

    print(f'{vendorStaff} {nameNum} mapping')
    #修改点 编号不连续
    id_number_AP = 0
    mapped_glIndex_vendor = []
    mapped_bankIndex_vendor = []

    for vendor, df in glData_vendor.groupby('Vendor Name'):
        # if vendor == 'PricewaterhouseCoopers, Taiwan':
        #     print('vendor', vendor)
        bankAccountName_vendor = map_vendor_byEntity.loc[map_vendor_byEntity['Vendor Name'] == f'{vendor}'.upper(), f'Bank Account {nameNum}']
        bankAccountName_vendor = bankAccountName_vendor.dropna()
        dic_bkValue_AP = {}
        if len(set(bankAccountName_vendor.to_list())) >= 1:
            for name in set(bankAccountName_vendor.to_list()):
                bkValue_list_AP = bankData_AP_left2.loc[bankData_AP_left2['Narrative'].str.contains(f'{str(name).strip()}', regex=False, case=False,
                                                                na=False), 'Credit/Debit amount'].to_dict()
                dic_bkValue_AP.update(bkValue_list_AP)
        if len(dic_bkValue_AP) == 0:
            conti
        bkIndex_vendor = list(dic_bkValue_AP.keys())
        subsets_bkIndex_vendor = get_sub_set(bkIndex_vendor)
        bkSum_vendor = sum(dic_bkValue_AP.values())
        dic_glValue_vendor= df['Amount Func Cur'].to_dict()
        glSum_vendor = df['Amount Func Cur'].sum()
        subsets_glIndex_vendor = get_sub_set(dic_glValue_vendor.keys())
        for subset_bkIndex_vendor in subsets_bkIndex_vendor:
            if common_data(subset_bkIndex_vendor, mapped_bankIndex_vendor):
                continue
            subset_bkValue_vendor = key_to_value(subset_bkIndex_vendor, dic_bkValue_AP)
            mapped_first = False
            if (account_cd == '101245') and (glSum_vendor - sum(subset_bkValue_vendor) in list_bankCharge):
                mapped_first = True
            if abs(sum(subset_bkValue_vendor) - glSum_vendor) < 0.03:
                mapped_first = True
            if mapped_first:
                id_number_AP = id_number_AP + 1
                bankData.loc[subset_bkIndex_vendor, 'Result'] = f'netoff'
                bankData.loc[subset_bkIndex_vendor, 'Category'] = f'AP {vendorStaff}'
                bankData.loc[subset_bkIndex_vendor, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                glData.loc[df.index, 'Result'] = f'netoff'
                glData.loc[df.index, 'Category'] = f'AP {vendorStaff}'
                glData.loc[df.index, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                mapped_glIndex_vendor = mapped_glIndex_vendor + list(df.index)
                mapped_bankIndex_vendor = mapped_bankIndex_vendor + list(subset_bkIndex_vendor)

        subsets_glIndex_vendor = [x for x in subsets_glIndex_vendor if len(set(x) & set(mapped_glIndex_vendor)) == 0]
        subsets_bkIndex_vendor = [x for x in subsets_bkIndex_vendor if len(set(x) & set(mapped_bankIndex_vendor)) == 0]

        for subset_glIndex_vendor in subsets_glIndex_vendor:
            if common_data(subset_glIndex_vendor, mapped_glIndex_vendor):
                continue
            subset_glValue_vendor = key_to_value(subset_glIndex_vendor, dic_glValue_vendor)
            mapped_second = False
            if (account_cd == '101245') and (sum(subset_glValue_vendor) - bkSum_vendor in list_bankCharge):
                mapped_second = True
            if abs(sum(subset_glValue_vendor) - bkSum_vendor) < 0.03:
                mapped_second = True
            if mapped_second:
                id_number_AP = id_number_AP + 1
                bankData.loc[bkIndex_vendor, 'Result'] = f'netoff'
                bankData.loc[bkIndex_vendor, 'Category'] = f'AP {vendorStaff}'
                bankData.loc[bkIndex_vendor, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                glData.loc[subset_glIndex_vendor, 'Result'] = f'netoff'
                glData.loc[subset_glIndex_vendor, 'Category'] = f'AP {vendorStaff}'
                glData.loc[subset_glIndex_vendor, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                mapped_glIndex_vendor = mapped_glIndex_vendor + subset_glIndex_vendor
                mapped_bankIndex_vendor = mapped_bankIndex_vendor + bkIndex_vendor

        subsets_glIndex_vendor = [x for x in subsets_glIndex_vendor if len(set(x) & set(mapped_glIndex_vendor)) == 0]
        subsets_bkIndex_vendor = [x for x in subsets_bkIndex_vendor if len(set(x) & set(mapped_bankIndex_vendor)) == 0]

        for subset_glIndex_vendor in subsets_glIndex_vendor:
            if common_data(subset_glIndex_vendor, mapped_glIndex_vendor):
                continue
            subset_glValue_vendor = key_to_value(subset_glIndex_vendor, dic_glValue_vendor)
            for subset_bkIndex_vendor in subsets_bkIndex_vendor:
                if common_data(subset_bkIndex_vendor, mapped_bankIndex_vendor):
                    continue
                subset_bkValue_vendor = key_to_value(subset_bkIndex_vendor, dic_bkValue_AP)
                mapped_third = False
                if (account_cd == '101245') and (sum(subset_glValue_vendor) - sum(subset_bkValue_vendor)) in list_bankCharge):
                    mapped_third = True
                if abs(sum(subset_glValue_vendor) - sum(subset_bkValue_vendor)) < 0.03:
                    mapped_third = True
                if mapped_third:
                    id_number_AP = id_number_AP + 1
                    bankData.loc[subset_bkIndex_vendor, 'Result'] = f'netoff'
                    bankData.loc[subset_bkIndex_vendor, 'Category'] = f'AP {vendorStaff}'
                    bankData.loc[subset_bkIndex_vendor, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                    glData.loc[subset_glIndex_vendor, 'Result'] = f'netoff'
                    glData.loc[subset_glIndex_vendor, 'Category'] = f'AP {vendorStaff}'
                    glData.loc[subset_glIndex_vendor, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                    mapped_glIndex_vendor = mapped_glIndex_vendor + subset_glIndex_vendor
                    mapped_bankIndex_vendor = mapped_bankIndex_vendor + subset_bkIndex_vendor

    return mapped_bankIndex_vendor, mapped_glIndex_vendor








            for subset_glIndex_vendor_2 in subsets_glIndex_vendor_2:
                if common_data(subset_glIndex_vendor_2, mapped_glIndex_vendor3):
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
                if mapped_third == True and vendor == 'PricewaterhouseCoopers, Taiwan':
                    print('mapped_third', mapped_third)
                    print('subset_bkIndex_vendor_3', subset_bkIndex_vendor_3)
                    print('subset_glIndex_vendor_2', subset_glIndex_vendor_2)
                if mapped_third:
                    if vendor == 'PricewaterhouseCoopers, Taiwan':
                        print('duplicate')
                    if common_data(subset_glIndex_vendor_2, mapped_glIndex_vendor3) or common_data(
                            subset_bkIndex_vendor_3, mapped_bankIndex_vendor3):
                        pass
                    else:
                        if vendor == 'PricewaterhouseCoopers, Taiwan':
                            print('recorded')
                        id_number_AP = id_number_AP + 1
                        bankData.loc[subset_bkIndex_vendor_3, 'Result'] = f'netoff'
                        bankData.loc[subset_bkIndex_vendor_3, 'Category'] = f'AP {vendorStaff}'
                        bankData.loc[
                            subset_bkIndex_vendor_3, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                        glData.loc[subset_glIndex_vendor_2, 'Result'] = f'netoff'
                        glData.loc[subset_glIndex_vendor_2, 'Category'] = f'AP {vendorStaff}'
                        glData.loc[
                            subset_glIndex_vendor_2, 'Identification'] = f'(AP {vendorStaff} netoff) ({now}) ({nameNum} {id_number_AP})'
                        mapped_glIndex_vendor3 = mapped_glIndex_vendor3 + list(subset_glIndex_vendor_2)
                        mapped_bankIndex_vendor3 = mapped_bankIndex_vendor3 + list(subset_bkIndex_vendor_3)

    mapped_bankIndex_vendor = mapped_bankIndex_vendor1 + mapped_bankIndex_vendor2 + mapped_bankIndex_vendor3
    mapped_glIndex_vendor = mapped_glIndex_vendor1 + mapped_glIndex_vendor2 + mapped_glIndex_vendor3

    return mapped_bankIndex_vendor, mapped_glIndex_vendor
