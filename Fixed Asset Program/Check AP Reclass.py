import numpy as np
import pandas as pd

df_target = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\FA Data.xlsx', sheet_name='Sheet1', header=0)
# print(df_target)
path_register = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\New FA Register 202312.xlsx'

df_reclass = pd.read_excel(rf'{path_register}', sheet_name='Reclass&Accrual List', header=1)
df_reclass = df_reclass[df_reclass['Category'].str.contains('Reclass', case=False, na=False)]
print(df_reclass)

id = 0
for inv_no, df_inv in df_reclass.groupby('Invoice'):

    id += 1

    account_amount = {160200: df_inv[160200].sum(), 161200: df_inv[161200].sum(), 162200: df_inv[162200].sum(), 163200: df_inv[163200].sum()}
    print(account_amount)
    month_posted = df_inv['Posted Month'].to_list()[0]
    print('month', month_posted)
    df_target_inv = df_target.loc[df_target['Invoice Number'].str.contains(fr'{inv_no}', case=False, na=False)]
    print('df_target_inv', df_target_inv)
    check = True

    for account, amount in account_amount.items():
        if not pd.isna(amount):
            print('inv_no', inv_no)
            print('account', account)
            gl_amount = df_target_inv.loc[(df_target_inv['Account Cd'] == int(account)) & (df_target_inv['Period Name'].str.contains(f'{month_posted}', case=False, na=False)), 'Amount Func Cur'].sum()
            if abs(gl_amount-amount) > 0.04:
                check = False

    if check:
        df_target.loc[df_target_inv.index, 'AP Reclass Check'] = f'Done {id}'
        df_reclass.loc[df_inv.index, 'AP Reclass Check'] = f'Done {id}'
    else:
        df_target.loc[df_target_inv.index, 'AP Reclass Check'] = f'Wrong Adj {id}'
        df_reclass.loc[df_inv.index, 'AP Reclass Check'] = f'Wrong Adj {id}'

# for ind, row in df_reclass.iterrows():
#     inv_no = row['Invoice']
#     account_amount = {160200: row[160200], 161200: row[161200], 162200: row[162200], 163200: row[163200]}
#     print(account_amount)
#     month_posted = row['Posted Month']
#     print('month', month_posted)
#     df_target_inv = df_target.loc[df_target['Invoice Number'].str.contains(fr'{inv_no}', case=False, na=False)]
#     print('df_target_inv', df_target_inv)
#     check = True
#
#     for account, amount in account_amount.items():
#         if not pd.isna(amount):
#             print('inv_no', inv_no)
#             print('account', account)
#             gl_amount = df_target_inv.loc[(df_target_inv['Account Cd'] == int(account)) & (df_target_inv['Period Name'].str.contains(f'{month_posted}', case=False, na=False)), 'Amount Func Cur'].sum()
#             if abs(gl_amount-amount) > 0.04:
#                 check = False
#
#     if check:
#         df_target.loc[df_target_inv.index, 'AP Reclass Check'] = 'Done'
#         df_reclass.loc[ind, 'AP Reclass Check'] = 'Done'

df_target.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\test\df_target.xlsx')
df_reclass.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\test\df_reclass.xlsx')



