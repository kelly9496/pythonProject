import pandas as pd

path_target = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\New FA Register 2023.07 - system data.xlsx'
path_source = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\New FA Register 2023.07 - depreciation rollover.xlsx'

category_list = ['Computer', 'Furniture', 'Office Equipment', 'Leasehold']
df_target_ori = pd.read_excel(f'{path_target}', sheet_name='资产', header=1)
df_target = df_target_ori[df_target_ori['资产类型(必填)'].isin(category_list)]
# print(df_target)

df_source = pd.read_excel(f'{path_source}', sheet_name='资产', header=1)
# df_source = df_source[df_source['资产类型(必填)'].isin(category_list)]
print(df_source.columns)
#
for ind, row in df_target.iterrows():
    asset_number = row['资产编码']
    remaining_years = df_source.loc[df_source['资产编码'] == f'{asset_number}', 'Remaining Years-Jul']
    print(remaining_years)
    if len(remaining_years)==1:
        # print(remaining_years.iloc[0])
        df_target_ori.loc[ind, 'Remaining useful life'] = remaining_years.iloc[0]
        print('done')
    depreciation_Jul = df_source.loc[df_source['资产编码'] == f'{asset_number}', 'JUL-23']
    if len(depreciation_Jul)==1:
        df_target_ori.loc[ind, 'JUL-23'] = depreciation_Jul.iloc[0]

#     df_target_ori.loc[ind, 'Remaining useful life'] = df_source.loc[df_source['资产编码']==f'{asset_number}', 'Remaining Years-Jul']
#     df_target_ori.loc[ind, 'JUL-23'] = df_source.loc[df_source['资产编码']==f'{asset_number}', 'JUL-23']
#
df_target_ori.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\depreciation.xlsx')