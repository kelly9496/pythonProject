import pandas as pd
df_fx = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202308\Accrual\HK Accrual Tracking - June.xlsx', sheet_name='Contract Lead Sheet', header=0)
df_fx = df_fx[df_fx['Accrual.1']<0]
print(df_fx.columns)
df_register = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202308\New FA Register 2023.08 - system data.xlsx', sheet_name='资产', header=1)
for ind, row in df_fx.iterrows():
    contract_number = row['合同编号']
    print(contract_number)
    fx = row['Accrual.1']
    print(fx)
    df_byContract = df_register.loc[df_register['VCP合同编号'].str.contains(f'{contract_number}', na=False)]
    print(df_byContract)
    contractValue_register = df_byContract['资产金额'].sum()
    contractValue_fx = row['合同总金额-USD']
    print(contractValue_register)
    print(contractValue_fx)
    if abs(contractValue_fx - contractValue_register) < 1:
        df_byContract['资产金额'] = df_byContract['资产金额'].map(lambda x: x-x/contractValue_register*fx)
        dic_adjustedValue = df_byContract['资产金额'].to_dict()
        df_register.update(df_byContract['资产金额'])
        if abs(df_byContract['资产金额'].sum() - row['Invoice Total'])<1:
            df_fx.loc[ind, 'notes'] = 'done'

df_fx.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202308\FX\df_fx.xlsx')
df_register.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202308\FX\df_register.xlsx')



