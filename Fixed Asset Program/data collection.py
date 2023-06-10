import pandas as pd
import os

folder_path = r'C:\Users\he kelly\Desktop\TB&GL\2023.5\0530'
target_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202305\Invoices by 2023.5.xlsx'
df_target=pd.read_excel(target_path,sheet_name='FA')
column_list = list(df_target.columns)

file_names = os.listdir(folder_path)
for file_name in file_names:
    file_path = os.path.join(folder_path, file_name)
    df = pd.read_excel(file_path, sheet_name='Drill', header=1)
    df_filtered = df.loc[df['Period Name'].str.contains('MAY-23') & df['Account Cd'].isin([160200,161200,162200,163200]),:]
    df_target = pd.concat([df_target, df_filtered], join='inner')

column_list_GL = list(df_target.columns)
column_list_palette = list(set(column_list).difference(set(column_list_GL)))+['Supplier InvoiceNo']



df_palette = pd.read_excel(target_path, sheet_name='Palette')
df_target = df_target.merge(df_palette[column_list_palette], how='left', left_on='Invoice Number',right_on='Supplier InvoiceNo')


df_target.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202305\FA Data.xlsx')