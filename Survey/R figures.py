import pandas as pd

path_TB1 = input('path_TB1 ')
month_TB1 = input('month_TB1 ')
path_TB2 = input('path_TB2 ')
month_TB2 = input('month_TB2 ')
path_TB3 = input('path_TB3 ')
month_TB3 = input('month_TB3 ')

target_path = fr'C:\Users\he kelly\Desktop\TW\TW Survey\{month_TB1} {month_TB2} {month_TB3}.xlsx'

df_TB1 = pd.read_excel(fr'{path_TB1}', sheet_name='TB Oracle Style', header=0)
df_TB2 = pd.read_excel(fr'{path_TB2}', sheet_name='TB Oracle Style', header=0)
df_TB3 = pd.read_excel(fr'{path_TB3}', sheet_name='TB Oracle Style', header=0)

list_account = [401001, 401002, 403000, 403001, 412000, 413002, 417000, 491001, 406000]

df_result = pd.DataFrame()

def TB_to_Result(df_TB, month, df_result):

    df_TB = df_TB[df_TB['Enity'] == 6001]
    account_summary = df_TB.groupby('Account')['Actual Net Activity PTD'].sum()
    account_summary = account_summary[list_account]
    df_summary = pd.DataFrame(account_summary)
    df_summary.rename(columns={'Actual Net Activity PTD': f'{month}'}, inplace=True)
    df_summary.loc['sum', f'{month}'] = df_summary[f'{month}'].sum()
    df_output = pd.concat([df_result, df_summary], axis=1)

    return df_output

df_result = TB_to_Result(df_TB1, month_TB1, df_result)
df_result = TB_to_Result(df_TB2, month_TB2, df_result)
df_result = TB_to_Result(df_TB3, month_TB3, df_result)

print(df_result)
df_result.to_excel(f'{target_path}')