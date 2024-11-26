import pandas as pd

path_register = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\New FA Register 202312 - depreciation accrual  DepreciationADJ Reclass Expense Disposal updated - Computer Split Recon.xlsx'
df_register = pd.read_excel(path_register, sheet_name='资产', header=1)

filtered_cols = list(df_register.filter(like='折旧-').columns)
columns_modification = ['累计折旧金额', '资产净值', '本月折旧'] + filtered_cols

df_transfer_BJ = df_register[df_register['2023 YE ADJ'] == 'transfer to BJ']
serial_number_BJ = set(df_transfer_BJ['IT Serial Number'].to_list())

for serial_number in serial_number_BJ:
    df_serial_number = df_transfer_BJ[df_transfer_BJ['IT Serial Number'] == f'{serial_number}']
    df_serial_number_BJ = df_serial_number[df_serial_number['所属仓库(必填)'].str.contains('BEI', na=False)]
    df_serial_number_SH = df_serial_number[df_serial_number['所属仓库(必填)'].str.contains('SHI', na=False)]
    df_register.loc[df_serial_number_BJ.index, '资产金额'] = df_serial_number_BJ['资产金额'].iloc[0]*(1-0.770471464019851) + df_serial_number_SH['资产净值'].iloc[0]*0.770471464019851
    df_register.loc[df_serial_number_SH.index, '资产金额'] = 0
    for column in columns_modification:
        df_register.loc[df_serial_number_BJ.index, f'{column}'] = df_serial_number_BJ[f'{column}'].iloc[0]*(1-0.770471464019851)
        print(df_register.loc[df_serial_number_BJ.index, f'{column}'])
        df_register.loc[df_serial_number_SH.index, f'{column}'] = 0
        df_register.loc[df_serial_number_SH.index, '资产名称(必填)'] = 'transferred out and to be deleted'



df_transfer_SH = df_register[df_register['2023 YE ADJ'] == 'transfer to SH']
serial_number_SH = set(df_transfer_SH['IT Serial Number'].to_list())


for serial_number in serial_number_SH:
    df_serial_number = df_transfer_SH[df_transfer_SH['IT Serial Number'] == f'{serial_number}']
    df_serial_number_SH = df_serial_number[df_serial_number['所属仓库(必填)'].str.contains('SHI', na=False)]
    df_serial_number_BJ = df_serial_number[df_serial_number['所属仓库(必填)'].str.contains('BEI', na=False)]
    df_register.loc[df_serial_number_SH.index, '资产金额'] = df_serial_number_SH['资产金额'].iloc[0]*0.770471464019851 + df_serial_number_BJ['资产净值'].iloc[0]*(1-0.770471464019851)
    df_register.loc[df_serial_number_BJ.index, '资产金额'] = 0
    for column in columns_modification:
        df_register.loc[df_serial_number_SH.index, f'{column}'] = df_serial_number_SH[f'{column}'].iloc[0]*0.770471464019851
        print(df_register.loc[df_serial_number_SH.index, f'{column}'])
        df_register.loc[df_serial_number_BJ.index, f'{column}'] = 0
        df_register.loc[df_serial_number_BJ.index, '资产名称(必填)'] = 'transferred out and to be deleted'

df_register.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\test\SH BJ Transferv4.xlsx')
