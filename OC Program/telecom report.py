import pandas as pd

df_helios = pd.read_excel(r'C:\Users\he kelly\Desktop\OC\telecom report\Helios 20240307.xlsx', sheet_name='分摊报表')
df_gl = pd.read_excel(r'C:\Users\he kelly\Desktop\OC\telecom report\Telecom actual report 2024.1-2 0306.xlsx', sheet_name='Raw Data', header=1)

df_check = df_gl[df_gl['Check'] == 'Y']

df_helios['费用发生日期'] = pd.to_datetime(df_helios['费用发生日期'])

for ind, row in df_check.iterrows():
    invoice_no = row['Invoice Number']
    staff_id = row['Emp Id']
    df_mapped = df_helios[df_helios['报销单单号'] == f'{invoice_no}']
    amount2024 = df_mapped.loc[df_mapped['费用发生日期'].dt.year == 2024, '本位币金额'].sum()
    currency2024 = df_mapped.loc[df_mapped['费用发生日期'].dt.year == 2024, '本位币币种']
    if len(currency2024):
        df_gl.loc[ind, 'Amount for 2024'] = amount2024
        df_gl.loc[ind, 'Currency'] = currency2024.iloc[0]

df_gl.to_excel(r'C:\Users\he kelly\Desktop\OC\telecom report\df_gl.xlsx')


