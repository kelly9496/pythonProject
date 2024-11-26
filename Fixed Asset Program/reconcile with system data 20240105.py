import pandas as pd
from datetime import datetime
import datetime as dt
import xlwings as xw
import shutil
import os
from decimal import Decimal, ROUND_HALF_UP
from myFunction import my_rounding

def merge_common_columns(df_main, df_affix, common_columns):




    rest_columns = df_affix.columns.difference(common_columns)
    # rest_columns_main = df_main.columns.difference(common_columns)
    # rest_columns_affix = df_affix.columns.difference(common_columns)
    print(rest_columns)
    list_to_be_mapped = df_affix.index.values.tolist()

    for ind, row in df_main.iterrows():

        df = pd.DataFrame(columns=common_columns)
        df.loc[0] = row[common_columns]
        print('df', df)
        empty_columns = df.columns[df.isnull().any()].tolist()
        common_columns = list(set(common_columns).difference(set(empty_columns)))
        print('common_columns', common_columns)
        print('len_common_columns', len(common_columns))

        print(ind)
        df_bool = df_affix[common_columns] == row[common_columns]
        # df_bool.to_excel(rf'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\test\df_bool\{ind}.xlsx')
        ind_true = df_bool[df_bool.all(axis=1)].index
        print('ind_true', ind_true)
        if len(ind_true) == 1:
            if ind_true in list_to_be_mapped:
                print('ind_true in list_to_be_mapped', ind_true in list_to_be_mapped)
                list_to_be_mapped.remove(ind_true)
                ind_true = ind_true.values[0]
            else:
                continue
        elif len(ind_true) > 1:
            ind_to_be_mapped = [x for x in ind_true.values if x in list_to_be_mapped]
            print('ind_to_be_mapped', ind_to_be_mapped)
            if ind_to_be_mapped:
                ind_true = ind_to_be_mapped[0]
                list_to_be_mapped.remove(ind_true)
            else:
                continue
        else:
            continue

        print('final ind_true', ind_true)
        df_main.loc[ind, rest_columns] = df_affix.loc[ind_true, rest_columns].to_dict().values()

    return df_main

path_register = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\FA Reconciliation - 20240521\New FA Register 202404 - sync with system.xlsx'
path_system = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\FA Reconciliation - 20240521\System Data -20240607.xlsx'
df_register = pd.read_excel(path_register, sheet_name='资产', header=1)
df_system = pd.read_excel(path_system, sheet_name='资产', header=1)
df_mapping = pd.read_excel(path_register, sheet_name='Mapping', header=0)
id = 0

# print(df_system)


columns_system = ['记录ID(不可修改)', '数据标题(不可修改)', '修改时间', '资产照片', '资产名称(必填)', '资产编码',  '备注', '资产型号', '是否存在', '是否盘点', '资产状态', '使用人', '存放地点', '图片', '创建人', '创建时间', '所属城市(必填)', '拥有者(必填)']
# print(df_system[columns_system])
columns_register = df_register.columns.difference(columns_system)
# 需要合并汇总金额的列
# 筛出带有关键字的columns
columns_value = ['累计折旧金额', '资产净值', '资产金额', '本月折旧', '累计折旧-202307', '折旧-202308', '累计折旧-202308',  '折旧-202309', '累计折旧-202309', '折旧-202310', '累计折旧-202310',  '折旧-202311', '累计折旧-202311', '折旧-202312', '累计折旧-202312']
# 需要确认信息是否一致的列
# columns_identical = columns_register.difference(columns_value)
columns_identical = ['Capex Code', 'IT Serial Number', 'VCP合同编号', '使用日期', '厂商名称', '发票号', '所属仓库(必填)', '折旧年限', '购买部门', '剩余折旧年限-202308', '剩余折旧年限-202309', '剩余折旧年限-202310', '剩余折旧年限-202311', '剩余折旧年限-202312']



# # 香港Ops及其他replacement对齐
# register_mappedNo = []
# for assetNo_register in set(df_mapping['Register'].to_list()):
#     # print(assetNo_register)
#     if assetNo_register in register_mappedNo:
#         continue
#     df_assetRegister = df_mapping[df_mapping['Register'].str.contains(f'{assetNo_register}')]
#     assetNo_system = df_assetRegister['System']
#     df_assetSystem = df_mapping[df_mapping['System'].str.contains(f'{assetNo_system.iloc[0]}')]
#     # print(assetNo_system)
#     if len(df_assetRegister) == 1:
#         # df_assetSystem = df_mapping[df_mapping['System'].str.contains(f'{assetNo_system.iloc[0]}')]
#
#         # case2: system1对register1
#         if len(df_assetSystem) == 1:
#             # print('assetNo_register', assetNo_register)
#             ind_register = df_register[df_register['资产编码'].str.contains(f'{assetNo_register}', na=False)].index.values[0]
#             ind_system = df_system[df_system['资产编码'].str.contains(f'{assetNo_system.iloc[0]}')].index.values[0]
#             register_mappedNo.append(assetNo_register)
#             df_register.loc[ind_register, columns_system] = df_system.loc[ind_system, columns_system].to_dict().values()
#             # print(df_register.loc[ind_register])
#
#         # case1: system1对register多，对register信息进行合并
#         if len(df_assetSystem) > 1:
#             #获取需要合并的register资产编号
#             list_registerNo = df_assetSystem['Register']
#             df_register_mapped = df_register[df_register['资产编码'].isin(list_registerNo)]
#             #获取list里的第一个ind作为合并行
#             ind_combine = df_register_mapped.index.values[0]
#             ind_delete = df_register_mapped.index.values[1:]
#             df_duplicate_test = df_register.loc[df_register['资产编码'].isin(list_registerNo), columns_identical]
#             ind_system = df_system[df_system['资产编码'].str.contains(f'{assetNo_system.iloc[0]}')].index.values[0]
#             if df_duplicate_test.duplicated(keep=False).all():
#                 df_register.loc[ind_combine, columns_value] = [df_register_mapped[f'{column}'].sum() for column in columns_value]
#                 df_register.loc[ind_combine, columns_system] = df_system.loc[ind_system, columns_system].to_dict().values()
#                 register_mappedNo += list_registerNo.to_dict().values()
#                 df_register.drop(index=ind_delete, inplace=True)
#
#
#     # case3: register1对system多 如果register里的一个资产对应系统里的多个资产，则对register里的资产信息进行分拆
#     if len(df_assetRegister) > 1:
#         #获取需要拆分为的system资产编号及个数
#         list_systemNo = assetNo_system
#         qty_system = len(list_systemNo)
#         df_system_mapped = df_system[df_system['资产编码'].isin(list_systemNo)]
#         value_sum = df_system_mapped['资产金额'].sum()
#         print(value_sum)
#
#         # print(list_systemNo)
#         # print(qty_system)
#         ind_register = df_register[df_register['资产编码'].str.contains(f'{assetNo_register}', na=False)].index.values[0]
#         if abs(value_sum) < 0.01:
#             df_system_mapped[columns_value] = [df_register.loc[ind_register, f'{column}']/qty_system for column in columns_value]
#         else:
#             for ind, row in df_system_mapped.iterrows():
#                 df_system_mapped.loc[ind, columns_value] = [df_register.loc[ind_register, f'{column}'] * (row['资产金额']/value_sum) for column in columns_value]
#         df_system_mapped[columns_identical] = [df_register.loc[ind_register, f'{column}'] for column in columns_identical]
#         # print(df_system_mapped[columns_value])
#         # print(df_system_mapped[columns_identical])
#         df_register = pd.concat([df_register, df_system_mapped])
#         df_register.drop(index=ind_register, inplace=True)
#         register_mappedNo.append(ind_register)
#
#     id += 1
# print(id)
# print(len(register_mappedNo))


# #筛出带有asset number且需要merge的部分和系统匹配
# df_merge_assetNo = df_register[df_register['记录ID(不可修改)'].isnull() & df_register['资产编码'].notnull()]
#
# # print(df_merge_assetNo)
#
# for ind, row in df_merge_assetNo.iterrows():
#     assetNo = row['资产编码']
#     ind_system_mapped = df_system.loc[df_system['资产编码'].str.contains(f'{assetNo}', na=False)].index.values[0]
#     # print(ind_system_mapped)
#     df_register.loc[ind, columns_system] = df_system.loc[ind_system_mapped, columns_system].to_dict().values()

# #将剩下无资产编号的数据导入系统
# path_template_add = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202309\系统导入模板\资产2023_10_25_13_47_33.xlsx'
# #将这个文件复制到当前目录下
# shutil.copyfile(path_template_add, fr'{os.path.dirname(path_register)}\system import template.xlsx')
# df_new = df_register[df_register['记录ID(不可修改)'].isnull() & df_register['资产编码'].isnull()]
# df_template_new = pd.DataFrame(columns = pd.read_excel(f'{path_template_add}', sheet_name='资产', header=1).columns)
# df_new = pd.concat([df_template_new, df_new], join='inner')
# df_new.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\df_new.xlsx')
# df_new_originalIndex = df_new.reset_index()
#
# app = xw.App(visible=True, add_book=False)
# book = app.books.open(fr'{os.path.dirname(path_register)}\system import template.xlsx')
# sheet = book.sheets["资产"]
# row_num = sheet['A2'].current_region.last_cell.row
# print(row_num)
# sheet['A{}'.format(row_num+1)].options(index=False, header=False).value=df_new
# book.save()


#筛出无asset number且需要merge的部分和系统匹配
df_merge_columns = df_register[df_register['记录ID(不可修改)'].isnull() & df_register['资产编码'].isnull()]
# df_merge_columns = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202312\test\df_merge_columnsv3.xlsx')
# df_merge_columns = df_merge_columns[df_merge_columns['记录ID(不可修改)'].isnull()]
df_system['创建时间'] = pd.to_datetime(df_system['创建时间'])
df_system = df_system.loc[df_system['创建时间'] > datetime.now() - dt.timedelta(days=30)]
# print(df_merge_columns)
common_columns = ['资产名称(必填)', '所属仓库(必填)', '资产类型(必填)', '是否存在', '是否盘点', '资产状态', '使用日期', '购买部门', 'VCP合同编号', '厂商名称', 'Capex Code', '资产金额', '发票号', '折旧年限', '累计折旧金额', '资产净值']
# print(common_columns)
#对齐两个dataframe的数据类型（金额、名称、时间）
df_system['使用日期'] = pd.to_datetime(df_system['使用日期'])
df_merge_columns['使用日期'] = pd.to_datetime(df_merge_columns['使用日期'])
df_system['资产名称(必填)'] = df_system['资产名称(必填)'].map(lambda x: x.strip())
df_merge_columns['资产名称(必填)'] = df_merge_columns['资产名称(必填)'].map(lambda x: x.strip())
df_system['资产金额'] = df_system['资产金额'].apply(lambda x: my_rounding(x))
df_system['累计折旧金额'] = df_system['累计折旧金额'].apply(lambda x: my_rounding(x))
df_system['资产净值'] = df_system['资产净值'].apply(lambda x: my_rounding(x))
df_merge_columns['资产金额'] = df_merge_columns['资产金额'].apply(lambda x: my_rounding(x))
df_merge_columns['累计折旧金额'] = df_merge_columns['累计折旧金额'].apply(lambda x: my_rounding(x))
df_merge_columns['资产净值'] = df_merge_columns['资产净值'].apply(lambda x: my_rounding(x))

# print(df_system)

df_merge_columns = merge_common_columns(df_merge_columns, df_system, common_columns)




#empty_columns = df.columns[df.isnull().all()]
df_register.update(df_merge_columns)
df_register.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\FA Reconciliation - 20240521\test.xlsx')
