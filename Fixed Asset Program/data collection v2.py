import pandas as pd
from datetime import datetime
import datetime as dt
import xlwings as xw
import shutil
import os




account_to_category = {160200: 'Computer', 161200: 'Office Equipment', 162200: 'Furniture', 163200: 'Leasehold'}
entity_to_endDate = {2821: '2028/1/31', 2841: '2025/3/31', 6001: '2029/9/30', 2801: '2028/2/29', 1601: '2031/1/31'}

#input1: GC gl path
gl_path = r'C:\Users\he kelly\Desktop\TB&GL\2024\0530\GL Dump query Jun.16 _ with RU and Currency Conversion and Category (27).xlsx'
#input2: VCP List path
invoice_path = r'C:\Users\he kelly\The Boston Consulting Group, Inc\Greater China Finance Team - Tax & Accounts Reporting\ME Working\GC VCP Accrual supporting\202405 amortization & accrual list.xlsx'
#input3: month period
month = 'MAY-24'
#input4: destination path
destination_path = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202405\FA Data.xlsx'
#input7: register path
path_register = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202405\New FA Register 202404 - sync with system.xlsx'
month_dt = datetime.strptime(f'{month}', '%b-%y')
month_str = datetime.strftime(month_dt, '%y%m')


df_invoice = pd.read_excel(invoice_path, sheet_name='Recorded invoice list', header=0)
df_gl = pd.read_excel(gl_path, sheet_name='Drill', header=1)
df_FA = df_gl[(df_gl['Period Name'] == f'{month}') & (df_gl['Account Cd'].isin([160200, 161200, 162200, 163200]))]
df_FA_AP = df_FA[df_FA['View Source'] == 'Payables']
df_FA_Ss = df_FA[df_FA['View Source'] == 'Spreadsheet']
df_target = df_FA_AP.merge(df_invoice[['Buyer', 'ContractNo', 'Supplier InvoiceNo']], how='left', left_on='Invoice Number', right_on='Supplier InvoiceNo')
df_target['资产类型(必填)'] = df_target['Account Cd'].map({160200: 'Computer', 161200: 'Office Equipment', 162200: 'Furniture', 163200: 'Leasehold'})

map_office = {2821: 'BCG SHZ Office', 2801: 'BCG SHI Office', 2841: 'BCG BEI Office', 1601: 'BCG HKG Office', 6001: 'BCG TAI Office'}
map_category = {160200: 'IT', 162200: 'Ops', 163200: 'Leasehold'}
df_target['所属仓库(必填)'] = df_target['Entity Cd'].map(map_office) + ' - ' + df_target['Account Cd'].map(map_category)

def depreciation_term(row):

    accountCd = row['Account Cd']
    invoiceDate = row['Invoice Date']
    entityCd = row['Entity Cd']

    if accountCd == 160200:
        term = 36
    if accountCd == 161200:
        term = 36

    end_date = entity_to_endDate[entityCd]
    month_end = datetime.strptime(end_date, '%Y/%m/%d').month
    year_end = datetime.strptime(end_date, '%Y/%m/%d').year
    month_invoice = invoiceDate.month
    year_invoice = invoiceDate.year
    interval = (year_end - year_invoice) * 12 + (month_end - month_invoice)

    if accountCd == 162200:
        term = max(60, interval)

    if accountCd == 163200:
        term = interval

    return term

df_target['折旧年限'] = df_target.apply(depreciation_term, axis=1)

df_target['是否存在'] = '是'
df_target['是否盘点'] = '否'
df_target['资产状态'] = '正常'
df_target['累计折旧金额'] = 0
df_target['资产净值'] = df_target['Amount Func Cur'] + df_target['累计折旧金额']

df_target['所属城市(必填)'] = df_target['Entity Cd'].map({2801: '上海', 2821: '深圳', 2841: '北京', 1601: '香港', 6001: '台湾'})
df_target['Qty'] = pd.Series()
df_target['Accrual/ADJ/Reclass'] = pd.Series()
df_target['To AP'] = pd.Series()
df_target['资产型号'] = pd.Series()
df_target['单价'] = pd.Series()
df_target['designation account'] = pd.Series()
df_target['购买部门'] = df_target['Account Cd'].map({160200: 'IT', 162200: 'Ops'})
df_target = df_target[df_target['Amount Func Cur'] != 0]


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

app = xw.App(visible=True, add_book=False)
book = app.books.open(path_register)
sheet = book.sheets["Reclass&Accrual List"]
sheet['A3'].options(index=False, header=False).value = df_reclass
book.save()
book.close()
app.quit()

df_target.loc[df_target['AP Reclass Check'].notnull(), 'Accrual/ADJ/Reclass'] = 'AP Reclass'

df_target.rename(columns={'Vendor Name': '厂商名称', '单价': '资产金额', 'JE Lines Desc': '资产名称(必填)', 'ContractNo': 'VCP合同编号', 'Invoice Number': '发票号', 'Invoice Date': '使用日期'}, inplace=True)

#
# #Output1: collected data
# df_target.to_excel(fr'{destination_path}')



#input5: processed data
input('Has the collected data been manually processed? ')

df_target = pd.read_excel(fr'{destination_path}', sheet_name='Sheet1')

#把这部分移至上一步
# df_target.rename(columns={'Vendor Name': '厂商名称', '单价': '资产金额', 'JE Lines Desc': '资产名称(必填)', 'ContractNo': 'VCP合同编号', 'Invoice Number': '发票号', 'Invoice Date': '使用日期'}, inplace=True)

#筛出需要新增进list的部分
df_transfer = df_target[df_target['Accrual/ADJ/Reclass'].isnull() | (df_target['Accrual/ADJ/Reclass'].str.contains('reclass', case=False) & df_target['designation account'].isin([160200, 161200, 162200, 163200]))]
print(df_transfer)
# df_transfer.to_excel(fr'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202309\test.xlsx')
#若有reclass，更新reclass后的资产类型
df_transfer.loc[df_transfer['designation account'].isin([160200, 161200, 162200, 163200]), '资产类型(必填)'] = df_transfer.loc[df_transfer['designation account'].isin([160200, 161200, 162200, 163200]), 'designation account'].map({160200: 'Computer', 161200: 'Office Equipment', 162200: 'Furniture', 163200: 'Leasehold'})
#若为leasehold类别，自动更新Qty和资产金额
df_transfer.loc[df_transfer['资产类型(必填)'] == 'Leasehold', '资产金额'] = df_transfer.loc[df_transfer['资产类型(必填)'] == 'Leasehold', 'Amount Func Cur']
df_transfer.loc[df_transfer['资产类型(必填)'] == 'Leasehold', 'Qty'] = 1

#若一个发票下有多个item, 自动填充空行
df_fill = df_transfer[df_transfer['Supplier InvoiceNo'].isnull()]
columns_fill = list(set(df_transfer.columns).difference({'资产名称(必填)', '资产金额', 'Qty', '资产型号'}))
for ind in sorted(df_fill.index):
    df_transfer.loc[ind, columns_fill] = df_transfer.loc[ind-1, columns_fill]

#填充所属仓库列
map_category = {'Computer': 'IT', 'Furniture': 'Ops', 'Leasehold': 'Leasehold'}
df_transfer.loc[df_transfer['所属仓库(必填)'].isnull(), '所属仓库(必填)'] = df_transfer.loc[df_transfer['所属仓库(必填)'].isnull(), 'Entity Cd'].map(map_office) + ' - ' + df_transfer.loc[df_transfer['所属仓库(必填)'].isnull(), '资产类型(必填)'].map(map_category)
df_transfer.loc[df_transfer['所属仓库(必填)'].isnull(), '所属仓库(必填)'] = df_transfer.loc[df_transfer['所属仓库(必填)'].isnull(), 'Entity Cd'].map(map_office) + ' - ' + df_transfer.loc[df_transfer['所属仓库(必填)'].isnull(), '购买部门']

def system_department(row):

    entity_code = row['Entity Cd']
    code_name = {2801: '上海', 2821: '深圳', 2841: '北京', 1601: '香港', 6001: '台湾'}
    code_abbreviation = {2801: 'SH', 2821: 'SZ', 2841: 'BJ', 1601: 'HK', 6001: 'TW'}
    entity_name = code_name[entity_code]
    entity_abbreviation = code_abbreviation[entity_code]

    category = row['资产类型(必填)']
    if category == 'Computer':
        department = f'{entity_name}' + 'IT'
    if category == 'Office Equipment':
        department = f'{entity_name}' + f'{row["购买部门"]}'
    if category == 'Furniture':
        department = f'{entity_name}' + f'{row["购买部门"]}'
    if category == 'Leasehold':
        department = f'{entity_name}' + f'{entity_abbreviation}'

    return department

df_transfer['所属城市(必填)'] = df_transfer.apply(system_department, axis=1)

#Check Error
for inv, df_inv in df_transfer.groupby('发票号'):
    sum_all = sum(set(df_inv['Amount Func Cur']))
    sum_detail = sum(df_inv['资产金额']*df_inv['Qty'])
    difference = abs(sum_all-sum_detail)
    if difference<0.1:
        pass
    else:
        print(f'difference found: {inv} {difference}')

#input6
modification = input('Do you want to check or made modification to df_transfer based on the difference found? (Y/N): ')
if modification == 'Y':
    df_transfer.to_excel(rf'{os.path.dirname(destination_path)}\FA Data - modification.xlsx')

    input('Check/Modification Finished?')
    df_transfer = pd.read_excel(rf'{os.path.dirname(destination_path)}\FA Data - modification.xlsx')


df_transfer['资产净值'] = df_transfer['资产金额'] - df_transfer['累计折旧金额']


#将df_transfer写入FA Register
df_register = pd.read_excel(rf'{path_register}', sheet_name='资产', header=1)
df_added = pd.DataFrame(columns=df_register.columns)
df_concat = pd.DataFrame(columns=df_register.columns)
for i in range(len(df_transfer)):
    qty = df_transfer.iloc[i]['Qty']
    print(qty)
    df = df_transfer[i:i+1]
    # print(df)
    for item in range(int(qty)):
        df_concat = pd.concat([df_concat, df], join='inner')
        print(df_concat)
df_added = pd.concat([df_added, df_concat], join='outer')

df_added.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202405\df_added.xlsx')

input('df_added printed?')

app = xw.App(visible=True, add_book=False)
book = app.books.open(path_register)
sheet = book.sheets["资产"]
row_num = sheet['A3'].current_region.last_cell.row
print(row_num)
sheet['A{}'.format(row_num+1)].options(index=False, header=False).value = df_added
book.save()
book.close()
app.quit()

#计算折旧并更新信息
register = pd.read_excel(f'{path_register}', sheet_name='资产', header=1)

#这里的current_month和上面的link起来
print('month_dt', month_dt)
current_month = datetime.strftime(month_dt, '%Y%m')
print('current_month', current_month)
# current_month = '202312'
# current_month_dt = datetime.strptime(f'{current_month}', '%Y%m')
current_M = month_dt.month
current_Y = month_dt.year
print('current_M', current_M)
print('current_Y', current_Y)
register['使用日期'] = pd.to_datetime(register['使用日期'])
register['资产净值'] = register['资产金额'] - register['累计折旧金额']

def remaining_term(row):

    if abs(row[-3]) > 0:
        term = row[-3]-1
    else:
        register_M = row['使用日期'].month
        register_Y = row['使用日期'].year
        interval = ((current_Y - register_Y)*12 + (current_M - register_M))
        term = row['折旧年限'] - interval + 1

    return term

def depreciation(row):

    if row[f'剩余折旧年限-{current_month}'] <= 0:
        depreciation = 0
    else:
        try:
            depreciation = row['资产净值']/row[f'剩余折旧年限-{current_month}']
        except:
            depreciation = 0

    return depreciation

register[f'剩余折旧年限-{current_month}'] = register.apply(remaining_term, axis=1)
# register[f'剩余折旧年限-{current_month}'].to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\剩余折旧年限-202310.xlsx')
print(register[f'剩余折旧年限-{current_month}'])

register[f'折旧-{current_month}'] = register.apply(depreciation, axis=1)
print(register[f'折旧-{current_month}'])

# # register[f'折旧-{current_month}'].to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\折旧-202310.xlsx')
register[f'累计折旧-{current_month}'] = register['累计折旧金额'] + register[f'折旧-{current_month}']
print(register[f'累计折旧-{current_month}'])
#
# generate Dep ADI
register['本月折旧'] = register[f'折旧-{current_month}']
print(register['本月折旧'])
#
# roll the register forward to current month
register['累计折旧金额'] = register[f'累计折旧-{current_month}']
register['资产净值'] = register['资产金额'] - register['累计折旧金额']
#
print(register)
path_register_current = fr'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\{current_month}\New FA Register {current_month} - depreciation updated.xlsx'
shutil.copyfile(path_register, path_register_current)
#
app = xw.App(visible=True, add_book=False)
book = app.books.open(path_register_current)
sheet = book.sheets["资产"]
sheet['A2'].options(index=False, header=True).value = register
book.save()
book.close()
app.quit()




#input 8
input("Reconciliation finished? ")
#input 9
path_template_add = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202309\系统导入模板\资产2023_10_25_13_47_33.xlsx'
#将这个文件复制到当前目录下
shutil.copyfile(path_template_add, fr'{os.path.dirname(path_register)}\system import template.xlsx')
df_recon = pd.read_excel(f'{path_register}', sheet_name='资产', header=1)
df_new = df_recon[df_recon['记录ID(不可修改)'].isnull()]
df_template_new = pd.DataFrame(columns = pd.read_excel(f'{path_template_add}', sheet_name='资产', header=1).columns)
df_new = pd.concat([df_template_new, df_new], join='inner')
df_new.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\df_new.xlsx')
df_new_originalIndex = df_new.reset_index()

app = xw.App(visible=True, add_book=False)
book = app.books.open(fr'{os.path.dirname(path_register)}\system import template.xlsx')
sheet = book.sheets["资产"]
row_num = sheet['A2'].current_region.last_cell.row
print(row_num)
sheet['A{}'.format(row_num+1)].options(index=False, header=False).value=df_new
book.save()

# # refresh系统里的累计折旧信息
# path_template_modification = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202309\资产2023_11_08_18_59_06.xlsx'
# #将这个文件复制到当前目录下
# shutil.copyfile(path_template_modification, fr'{os.path.dirname(path_register)}\system modification template.xlsx')
# df_recon = pd.read_excel(f'{path_register}', sheet_name='资产', header=1)
# df_old = df_recon[df_recon['记录ID(不可修改)'].notnull()]
# df_template_old = pd.DataFrame(columns = pd.read_excel(f'{path_template_modification}', sheet_name='资产', header=1).columns)
# print(df_template_old)
# print(df_old)
# df_old = pd.concat([df_template_old, df_old], join='inner')
# df_old.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\df_old.xlsx')
# if len(df_old) < 10000:
#     app = xw.App(visible=True, add_book=False)
#     book = app.books.open(fr'{os.path.dirname(path_register)}\system modification template.xlsx')
#     sheet = book.sheets["资产"]
#     row_num = sheet['A2'].current_region.last_cell.row
#     print(row_num)
#     sheet['A{}'.format(row_num + 1)].options(index=False, header=False).value = df_old
#     book.save()
# if (len(df_old) >= 10000) and (len(df_old) < 20000):
#     print('int(len(df_old)/2)', int(len(df_old)/2))
#     df_old1 = df_old.iloc[0:int(len(df_old)/2), :]
#     df_old2 = df_old.iloc[int(len(df_old)/2):, :]
#     df_old1.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\df_old1.xlsx')
#     df_old2.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\df_old2.xlsx')
#     shutil.copyfile(fr'{os.path.dirname(path_register)}\system modification template.xlsx', fr'{os.path.dirname(path_register)}\system modification template1.xlsx')
#     os.rename(fr'{os.path.dirname(path_register)}\system modification template.xlsx', fr'{os.path.dirname(path_register)}\system modification template2.xlsx')
#     app = xw.App(visible=True, add_book=False)
#     book = app.books.open(fr'{os.path.dirname(path_register)}\system modification template1.xlsx')
#     sheet = book.sheets["资产"]
#     row_num = sheet['A2'].current_region.last_cell.row
#     print(row_num)
#     sheet['A{}'.format(row_num + 1)].options(index=False, header=False).value = df_old1
#     book.save()
#     book = app.books.open(fr'{os.path.dirname(path_register)}\system modification template2.xlsx')
#     sheet = book.sheets["资产"]
#     row_num = sheet['A2'].current_region.last_cell.row
#     print(row_num)
#     sheet['A{}'.format(row_num + 1)].options(index=False, header=False).value = df_old2
#     book.save()
#
#
#



#将系统信息更新到空白处
# input 10
name_system = input('Please input the system file name: ')
path_system = fr'{os.path.dirname(path_register)}\{name_system}.xlsx'
df_system = pd.read_excel(rf'{path_system}', sheet_name='资产', header=1)
print(df_system['创建时间'])
df_system['创建时间'] = pd.to_datetime(df_system['创建时间'])
# df_system = df_system.loc[df_system['创建时间'] > datetime.now() - dt.timedelta(hours=5)]
df_system = df_system.loc[df_system['创建时间'] > datetime.now() - dt.timedelta(days=5)]
print(df_system)
df_system.to_excel(fr'{os.path.dirname(path_register)}\df_system_test.xlsx')
df_new['使用日期'] = df_new['使用日期'].astype(str)

df_new = df_new.merge(df_system, how='left', on=list(df_new.columns))
df_new = df_new.merge(df_new_originalIndex, how='left', on=list(df_new.columns))

df_new.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202310\df_new.xlsx')





