import shutil
import os
import re
from datetime import datetime
import pandas as pd
import xlwings as xw

class ExcelLog:
    def __init__(self, number):
        self.max_log_number = number | 10
        self.log_number = 0

    def log(self, dataframe, desc):
        if self.log_number >= self.max_log_number:
            return
        now = str(datetime.now()).replace(':', '_')
        dataframe.to_excel(rf'C:\Users\he kelly\Desktop\ME Working\202310\test\{desc}_{now}.xlsx')
        self.log_number += 1

excel_log = ExcelLog(10)

current_month = '202310'
# os.makedirs(fr'C:\Users\he kelly\Desktop\ME Working\{current_month}')

path_folder_ME = r'C:\Users\he kelly\Desktop\ME Working\202309'
files_ME = os.listdir(rf'{path_folder_ME}')
df_ADI_FA = pd.DataFrame()
df_ADI_OC = pd.DataFrame()
entity_to_path = dict()
for file_ME in files_ME:
    if 'WebADI' in file_ME:
        file_path_ME = os.path.join(path_folder_ME, file_ME)
        destination_path_ME = fr'C:\Users\he kelly\Desktop\ME Working\{current_month}\{file_ME}'
        # shutil.copyfile(file_path_ME, fr'{destination_path_ME}')
        re_Entity = re.compile(r'WebADI-(\w+\b).*')
        match_Entity = re_Entity.search(file_ME)
        if match_Entity:
            entity = match_Entity.group(1)
            print(entity)
        entity_to_path.update({f'{entity}': f'{destination_path_ME}'})
        df_entity_FA = pd.read_excel(fr'{destination_path_ME}', sheet_name='FA', header=10)
        df_entity_OC = pd.read_excel(fr'{destination_path_ME}', sheet_name='Other Cost', header=10)
        df_ADI_FA = pd.concat([df_ADI_FA, df_entity_FA])
        df_ADI_OC = pd.concat([df_ADI_OC, df_entity_OC])
        # print(entity_to_path)

# excel_log.log(df_ADI_FA, 'df_ADI_FA')
# excel_log.log(df_ADI_OC, 'df_ADI_OC')

code_to_entity = {1601: 'HK', 6001: 'TW', 2821: 'SZ', 2841: 'BJ', 2801: 'SH'}

path_register_current = fr'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\{current_month}\New FA Register {current_month}.xlsx'
df_ADI_Dep = pd.read_excel(f'{path_register_current}', sheet_name='Dep ADI', header=6)
df_ADI_RecAcc = pd.read_excel(f'{path_register_current}', sheet_name='Reclass&Accrual ADI', header=2)
entity_accrual = set(df_ADI_RecAcc.loc[df_ADI_RecAcc['Batch Name'].str.contains('FA accrual', na=False, case=False), 'Ent'].to_list())
print(entity_accrual)

excel_log.log(df_ADI_Dep, 'df_Dep')
columns_ADI = df_ADI_Dep.columns.intersection(df_ADI_FA.columns)
print(df_ADI_RecAcc)
# columns_ADI = columns_ADI.drop('Unnamed: 0')
# print(df_ADI_FA['Ent'][1])

for entity_cd in code_to_entity.keys():
    if entity_cd == 2841:
        continue
    if entity_cd != 6001:
        continue

    # Section1: FA Depreciation
    print(entity_cd)

    df_FA_Dep = df_ADI_FA.loc[(df_ADI_FA['Ent'] == int(entity_cd)) & (df_ADI_FA['Batch Name'] == 'FA Depreciation')]
    first_row_Dep = int(df_FA_Dep.iloc[0:1].index.values)
    last_row_Dep = int(df_FA_Dep.tail(1).index.values)
    print(first_row_Dep, last_row_Dep)
    df_Dep = df_ADI_Dep[columns_ADI].loc[(df_ADI_Dep['Ent'] == int(entity_cd)) & (df_ADI_Dep['Batch Name'] == 'FA Depreciation')]
    # print(df_Dep)
    entity = code_to_entity[entity_cd]
    # print(entity)
    path_current = entity_to_path[f'{entity}']
    # print(path_current)

    # Section2: FA Accrual
    accrual_FA = False
    df_FA_accrual = df_ADI_FA.loc[(df_ADI_FA['Ent'] == int(entity_cd)) & (df_ADI_FA['Batch Name'].str.contains('FA Accrual', na=False, case=False))]
    print(df_FA_accrual)
    if len(df_FA_accrual):
        accrual_FA = True
        first_row_accrual = int(df_FA_accrual.iloc[0:1].index.values)
        last_row_accrual = int(df_FA_accrual.tail(1).index.values)
        print(first_row_accrual, last_row_accrual)

    accrual_ADI = False
    if entity_cd in entity_accrual:
        df_ADI_accrual = df_ADI_RecAcc.loc[(df_ADI_RecAcc['Ent'] == int(entity_cd)) & (df_ADI_RecAcc['Batch Name'].str.contains('FA Accrual', na=False, case=False))]
        if len(df_ADI_accrual):
            accrual_ADI = True
        print(df_ADI_accrual)
        print(len(df_ADI_accrual))
        # first_row_ADIacc = int(df_ADI_accrual.iloc[0:1].index.values)
        # last_row_ADIacc = int(df_ADI_accrual.tail(1).index.values)
        # print(first_row_ADIacc, last_row_ADIacc)
    # else:
    #     df_ADI_accrual = df_FA_accrual
    #     if len(df_ADI_accrual):
    #         accrual_ADI = True
    #     line_descriptions = set(df_ADI_accrual['Line Description'].to_list())
    #     print(line_descriptions)
    #     for line in line_descriptions:
    #         accrual_index = df_ADI_accrual.loc[(df_ADI_accrual['Line Description'] == f'{line}') & (~df_ADI_accrual['Batch Name'].str.contains('reverse', na=False, case=False))].index
    #         accrual_debit =
    #         df_ADI_accrual.loc[(df_ADI_accrual['Line Description'] == f'{line}') & df_ADI_accrual['Batch Name'].str.contains('reverse', na=False, case=False), 'Debit'] = df_ADI_accrual.loc[(df_ADI_accrual['Line Description'] == f'{line}') & (~df_ADI_accrual['Batch Name'].str.contains('reverse', na=False, case=False)), 'Credit']
    #         df_ADI_accrual.loc[(df_ADI_accrual['Line Description'] == f'{line}') & df_ADI_accrual['Batch Name'].str.contains('reverse', na=False, case=False), 'Credit'] = df_ADI_accrual.loc[(df_ADI_accrual['Line Description'] == f'{line}') & (~df_ADI_accrual['Batch Name'].str.contains('reverse', na=False, case=False)), 'Debit']

    # Section3: FA Reclass
    df_FA_reclass = df_ADI_FA.loc[(df_ADI_FA['Ent'] == int(entity_cd)) & (df_ADI_FA['Batch Name'].str.contains('FA Reclass', na=False, case=False))]
    print(df_FA_reclass)
    reclass_FA = False
    if len(df_FA_reclass):
        reclass_FA = True
        first_row_reclass = int(df_FA_reclass.iloc[0:1].index.values)
        last_row_reclass = int(df_FA_reclass.tail(1).index.values)
        print(first_row_reclass, last_row_reclass)



    app = xw.App(visible=True, add_book=False)
    book = app.books.open(path_current)
    sheet = book.sheets["FA"]
    sheet[f'A{12 + first_row_Dep}'].options(index=False, header=False).value = df_Dep
    if accrual_FA:
        sheet.range(f'A{12 + first_row_accrual}:A{12 + last_row_accrual}').api.EntireRow.Delete()
    if reclass_FA:
        sheet.range(f'A{12 + first_row_reclass}:A{12 + last_row_reclass}').api.EntireRow.Delete()

    if accrual_ADI:
        sheet.range(f'{12 + last_row_Dep + 1}:{12 + last_row_Dep + len(df_ADI_accrual) + 2}').api.Insert()
        sheet[f'A{12 + last_row_Dep + 2}:A{12 + last_row_Dep + len(df_ADI_accrual) + 1}'].options(index=False, header=False).value = df_ADI_accrual
    book.save()







