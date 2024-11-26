import pandas as pd
from datetime import datetime
import datetime as dt

df_recon = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\New FA Register 202311.xlsx', sheet_name='资产', header=1)
df_new = df_recon[df_recon['记录ID(不可修改)'].isnull()]
df_new = df_recon[df_recon['资产编码'].isnull()]
# df_new.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\df_new_test.xlsx')

df_system = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\system full data 1228.xlsx', sheet_name='资产', header=1)
print(df_system['创建时间'])
df_system['创建时间'] = pd.to_datetime(df_system['创建时间'])
# df_system = df_system.loc[df_system['创建时间'] > datetime.now() - dt.timedelta(hours=5)]
df_system = df_system.loc[df_system['创建时间'] > datetime.now() - dt.timedelta(days=5)]
print(df_system)
# df_system.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\df_system_test.xlsx')
df_new['使用日期'] = df_new['使用日期'].astype(str)

merge_columns = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\system import template.xlsx', header=1).columns

df_new = df_new.merge(df_system, how='left', on=list(merge_columns))

df_new.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\df_new.xlsx')


