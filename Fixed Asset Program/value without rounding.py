import pandas as pd

register_source = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\Value without rounding\PRC New FA Register.06.xlsx', sheet_name='List')
register_target_original = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\Value without rounding\leasehold and furniture v2.xlsx')

print(register_source.head())

register_target = register_target_original.loc[register_target_original['所属仓库(必填)'].str.contains('SHI')]
register_target = register_target.loc[register_target['资产类型(必填)'] == 'Furniture']
list_contract = set(register_target['VCP合同编号'].to_list())
print(list_contract)

for contract in list_contract:
    df_source = register_source[register_source['Mapping'].str.contains(f'{contract}', na=False)]
    df_target = register_target[register_target['VCP合同编号'] == f'{contract}']
    record_matched_source = []
    record_matched_target = []
    for index_s, value_s in df_source['资产金额'].items():
        for index_t, value_t in df_target['资产金额'].items():
            if value_s - value_t > -1 and value_s - value_t < 1:
                if index_s in record_matched_source or index_t in record_matched_target:
                    continue
                register_target_original.loc[index_t, '资产金额'] = value_s
                record_matched_target.append(index_t)
                record_matched_source.append(index_s)

register_target_original.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\Value without rounding\leasehold and furniture - without rounding v2.xlsx')


