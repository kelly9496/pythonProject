import pandas as pd
import re
import numpy as np

file_path_target = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\HK Acc Dep\HK New FA Register.xlsx'
file_path_source = r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\HK Acc Dep\HK Fixed Assets Register 2023.07.xlsx'
register_new = pd.read_excel(file_path_target, header=0, sheet_name='Sheet1')
print(register_new.columns)
register_source = pd.read_excel(file_path_source, sheet_name='Computer', header=4)
print(register_source)
register_target = register_new[register_new[r"资产类型(必填)"] == 'Computer']


list_mapping = set(register_target['VCP合同编号'].to_list())
print(list_mapping)
list_pass =[ ]

# #以contract为比对依据
# for mapping in list_mapping:
#     if mapping in list_pass:
#         pass
#     else:
#         print(mapping)
#         accDep = register_source.loc[register_source['Mapping'] == f'{mapping}', 'Acc Dep till 2023.07'].sum()
#         cost_source = register_source.loc[register_source['Mapping'] == f'{mapping}', 'Cost'].sum()
#         df_target = register_target.loc[register_target['VCP合同编号'] == f'{mapping}']
#         cost_target = df_target['资产金额'].sum()
#         # print(cost_target, cost_target)
#         if ((float(cost_target) - float(cost_source)) < 2) & ((float(cost_target) - float(cost_source)) > -2):
#             def accDep_calculation(x, cost, depreciation):
#                 result = x / cost * depreciation
#                 return result
#             for index, row in df_target.iterrows():
#                 accDep_calculated = accDep_calculation(row['资产金额'], cost_target, accDep)
#                 register_new.loc[index, '累计折旧金额'] = accDep_calculated


# 以invoice为比对依据
register_target_invoice = register_target
list_invoice = set(register_target_invoice['发票号'].to_list())
print(list_invoice)
list_invoiceNumber = []
for invoice in list_invoice:
    invoice = str(invoice)
    print(invoice)
    invoice_number = re.findall(r'\d{6,}[-_]?\d*\b', f'{invoice}')
    if len(invoice_number) == 0:
        continue
    else:
        print(invoice_number)
        invoice_number=invoice_number[0]
        list_invoiceNumber.append(invoice_number)
set_invoiceNumber = set(list_invoiceNumber)
print(set_invoiceNumber)
for invoice in set_invoiceNumber:
    df_source_invoice = register_source.loc[register_source['Invoice No'].str.contains(f'{invoice}', na=False)]
    print(df_source_invoice)
    accDep_invoice = df_source_invoice['Acc Dep till 2023.07'].sum()
    print(accDep_invoice)
    cost_source_invoice = df_source_invoice['Cost'].sum()
    df_target_invoice = register_target_invoice[register_target_invoice['发票号'].str.contains(f'{invoice}', na=False)]
    cost_target_invoice = df_target_invoice['资产金额'].sum()
    print(cost_source_invoice, cost_target_invoice)
    if ((float(cost_target_invoice) - float(cost_source_invoice)) < 2) & ((float(cost_target_invoice) - float(cost_source_invoice)) > -2):
        def accDep_calculation(x, cost, depreciation):
            result = x / cost * depreciation
            return result
        for index, row in df_target_invoice.iterrows():
            accDep_calculated = accDep_calculation(row['资产金额'], cost_target_invoice, accDep_invoice)
            register_new.loc[index, '累计折旧金额'] = accDep_calculated

register_new.to_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202307\HK Acc Dep\HK New FA Register.xlsx')





