import pandas as pd


columns_replacement = ['记录ID(不可修改)', '资产编码', '资产名称(必填)', '备注', '拥有者(必填)', '资产照片', '资产型号', '使用人', '存放地点', '图片']

df_register = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\FA Reconciliation - 20240521\New FA Register 202404 (version 1).xlsx', sheet_name='资产', header=1)
df_system = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\FA Reconciliation - 20240521\System Data - 20240521 - Computer Notes Update.xlsx', sheet_name='资产', header=1)
df_replacement = df_system[columns_replacement]
print(df_replacement)

df_replacement = df_replacement.set_index('资产编码')
print(df_replacement)
df_register = df_register.set_index('资产编码')
print(df_register)
df_register.update(df_replacement, overwrite=True)

#
# for ind, row in df_register.iterrows():
#     asset_number = row['资产编码']
#     ind_system = df_system[df_system['资产编码'] == f'{asset_number}'].index
#     # print(ind, ind_system)
#     if len(ind_system):
#         print(ind, ind_system[0])
#     for column in columns_replacement:
#         print(column)
#         # df_register.loc[ind, f'{column}'] = df_system.loc[ind_system, f'{column}']
#
df_register.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\FA Reconciliation - 20240521\test\df_register_replaced3.xlsx')

