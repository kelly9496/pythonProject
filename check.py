# import re
# from collections import namedtuple
# import pandas as pd
# import pdfplumber
# import os
# import datetime as dt
#
# pdf_path = r'C:\Users\he kelly\Desktop\Alteryx & Python\reimbursement\New folder\BJ TS 2023\Apr\1\BJ.pdf'
# with pdfplumber.open(pdf_path) as pdf:
#     text = ''
#     for page in pdf.pages:
#         text = text + page.extract_text()
#
# list_staffName = []
# list_staffNo = []
# list_amount = []
# list_date = []
# totalAmount = []
# entityName = []
#
# for line in text.splitlines():
#
#     # 抓取员工人名
#     re_staffName = re.compile(r'Trading Partner (.+) Processing')
#     match_staffName = re_staffName.search(line)
#     if match_staffName:
#         staffName = match_staffName.group(1)
#         list_staffName.append(staffName)
#
#     # 抓取员工工号
#     re_staffNo = re.compile(r'^Number (\d+)$')
#     match_staffNo = re_staffNo.search(line)
#     if match_staffNo:
#         staffNo = match_staffNo.group(1)
#         list_staffNo.append(staffNo)
#
#     # 抓取付款金额
#     re_amount = re.compile(r'Payment Amount (.+) Supplier Number')
#     match_amount = re_amount.search(line)
#     if match_amount:
#         amount = match_amount.group(1)
#         list_amount.append(amount)
#
#     # 抓取付款日期
#     re_date = re.compile(r'Payment Date (.+) Payment Method')
#     match_date = re_date.search(line)
#     if match_date:
#         date = match_date.group(1)
#         list_date.append(date)
#
#     # 抓取batch付款总金额
#     re_sum = re.compile(r'Total (.*[0-9]+[.][0-9]{2})$')
#     match_sum = re_sum.search(line)
#     if match_sum:
#         total = match_sum.group(1)
#         totalAmount = total
#
#     # 抓取entity
#     re_entity = re.compile(r'Legal Entity (.+)$')
#     match_entity = re_entity.search(line)
#     if match_entity:
#         entity = match_entity.group(1)
#         entityName = entity
#
# print(totalAmount)
# print('list_staffName', list_staffName)
# print('list_staffNo', list_staffNo)
# print('list_amount', list_amount)
# print('list_date', list_date)
# print('totalAmount', totalAmount)
# print('entityName', entityName)
# print(pd.DataFrame(list_staffName))
#
# df_reimPayment = pd.DataFrame()
# if len(list_staffName):
#     df_reimPayment['Staff Name'] = pd.DataFrame(list_staffName)
# if len(list_staffNo):
#     df_reimPayment['Staff No'] = pd.DataFrame(list_staffNo)
# if len(list_amount):
#     df_reimPayment['Payment Amount'] = pd.DataFrame(list_amount)
# if len(list_date):
#     df_reimPayment['Payment Date'] = pd.DataFrame(list_date)
#     df_reimPayment['Payment Date'] = pd.to_datetime(df_reimPayment['Payment Date'])
#     df_reimPayment['Month'] = df_reimPayment['Payment Date'].dt.month
#     month_conversion = {1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUN', 7: 'JUL', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'}
#     df_reimPayment['Month'] = df_reimPayment['Month'].map(lambda x: month_conversion[x])
# if len(totalAmount):
#     df_reimPayment['Batch Amount'] = totalAmount
# if len(entityName):
#     df_reimPayment['Entity'] = entityName
#
# print(df_reimPayment)
#
#
# # return df_reimPayment

a