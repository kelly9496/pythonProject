import xlwings as xw
from datetime import datetime, timedelta

import os
import pandas as pd

weekly_tracker_path = r'C:\Users\he kelly\Desktop\FRC Tracker\Final version every week'

all_files = os.listdir(weekly_tracker_path)

weekly_trackers = [f for f in all_files if f.endswith('.xlsx')]
columns_extract = ['Request Number', 'Proposal Code', 'Proposal Start', 'Proposal End']
combined_trackers = pd.DataFrame(columns=columns_extract)


for tracker in weekly_trackers:
    tracker_path = os.path.join(weekly_tracker_path, tracker)
    df_tracker = pd.read_excel(tracker_path, sheet_name='Request in Proposal Ease', header=0)
    column_start = df_tracker.filter(regex='Proposal Start').columns.values
    column_end = df_tracker.filter(regex='Proposal End').columns.values
    df_tracker = df_tracker.loc[:, ['Request Number', 'Proposal Code', column_start[0], column_end[0]]]
    # df_tracker = df_tracker.dropna()
    df_tracker = df_tracker.rename(columns={f'{column_start[0]}': 'Proposal Start Day', f'{column_end[0]}': 'Proposal End Day'})
    combined_trackers = pd.concat([combined_trackers, df_tracker])

print(combined_trackers)
# combined_trackers.to_excel(r'C:\Users\he kelly\Desktop\FRC Tracker\test_combined1.xlsx')

combined_trackers = combined_trackers.set_index('Request Number')
print(combined_trackers)

today = datetime.today()
last_friday = today - timedelta(today.weekday()+3)


app = xw.App(visible=True)

source_wb = app.books.open(r'C:\Users\he kelly\Desktop\FRC Tracker\Case & Proposal request 0308_Overall.xlsx')
target_wb= app.books.open(r'C:\Users\he kelly\Desktop\FRC Tracker\CPC Request Tracker 20240308.xlsx')

source_sheet_case = source_wb.sheets['Case']
target_sheet_case = target_wb.sheets['Case']

source_sheet_prop = source_wb.sheets['Proposal Ranking']
target_sheet_prop = target_wb.sheets['Proposal ']
target_sheet_inv = target_wb.sheets['Investment calculation']

#case section

#定义复制源区域和目标区域
source_row_num_case = source_sheet_case['A1'].current_region.last_cell.row
target_row_num_case = target_sheet_case['A1'].current_region.last_cell.row
source_range_case1 = source_sheet_case.range(f'A2:Z{source_row_num_case}')
print(target_row_num_case)
target_range_case1 = target_sheet_case.range(f'A{target_row_num_case + 1}')
source_range_case2 = source_sheet_case.range(f'AA2:AF{source_row_num_case}')
print(target_row_num_case)
target_range_case2 = target_sheet_case.range(f'AA{target_row_num_case + 1}')

#复制源区域的数据
source_range_case1.copy(target_range_case1)
source_range_case2.copy(target_range_case2)

#proposal section

#定义复制源区域和目标区域
source_row_num_prop = source_sheet_prop['A1'].current_region.last_cell.row
target_row_num_prop = target_sheet_prop['A1'].current_region.last_cell.row
target_row_num_inv = target_sheet_inv['A1'].current_region.last_cell.row

source_range_prop = source_sheet_prop.range(f'A2:T{source_row_num_prop}')
# print(target_row_num_prop)
target_range_prop = target_sheet_prop.range(f'A{target_row_num_prop + 1}')
target_range_inv = target_sheet_inv.range(f'A{target_row_num_inv + 1}')

df_id = pd.DataFrame(source_sheet_prop.range(f'C2:C{source_row_num_prop}').value)
df_id = df_id.set_axis(['ID'], axis=1)
df_id['Date'] = last_friday
df_id['Proposal Start Day'] = df_id['ID'].map(lambda x: combined_trackers.loc[x, 'Proposal Start Day'])
df_id['Proposal End Day'] = df_id['ID'].map(lambda x: combined_trackers.loc[x, 'Proposal End Day'])
# df_id['Project Code'] = df_id['ID'].map(lambda x: combined_trackers.loc[x, 'Proposal Code'])
df_id['Date'] = df_id['Date'].dt.strftime('%Y-%m-%d')
df_id['Proposal Start Day'] = df_id['Proposal Start Day'].dt.strftime('%Y-%m-%d')
df_id['Proposal End Day'] = df_id['Proposal End Day'].dt.strftime('%Y-%m-%d')
print(df_id)

#复制源区域的数据
source_range_prop.copy(target_range_prop)
source_range_prop.copy(target_range_inv)

target_sheet_prop['U{}'.format(target_row_num_prop + 1)].options(index=False, header=False).value = df_id.loc[:, ['Date', 'Proposal Start Day', 'Proposal End Day']]
target_sheet_inv['U{}'.format(target_row_num_inv + 1)].options(index=False, header=False).value = df_id.loc[:, 'Date']


target_wb.save()

df_inv = target_sheet_inv.range('A1').expand('table').options(pd.DataFrame, header=1, index=False).value
print(df_inv)
# df_inv.to_excel(r'C:\Users\he kelly\Desktop\FRC Tracker\df_inv.xlsx')
print(df_inv['Date'])
ind_delete = df_inv.loc[df_inv['''Week X of Total Proposal duration
(To be provided from next week)'''].notnull() & ((df_inv['Date'] > last_friday - timedelta(days=1))&(df_inv['Date'] < last_friday + timedelta(days=1)))].index
df_inv.loc[ind_delete, '''Investment amount
In USD'M'''] = ''
print(ind_delete)
# df_inv.to_excel(r'C:\Users\he kelly\Desktop\FRC Tracker\df_inv1.xlsx')
df_inv_merged = pd.merge(df_inv, combined_trackers['Proposal Code'], how='left', left_on='Request IT', right_index=True)
df_inv_merged.to_excel(r'C:\Users\he kelly\Desktop\FRC Tracker\df_inv_merged.xlsx')
# combined_trackers.to_excel(r'C:\Users\he kelly\Desktop\FRC Tracker\combined_tracker.xlsx')
# df_inv.to_excel(r'C:\Users\he kelly\Desktop\FRC Tracker\df_inv5.xlsx')

target_sheet_inv['P2'].options(index=False, header=False).value = df_inv.loc[:, "Investment amount\nIn USD'M"]
# target_sheet_inv['V2'].options(index=False, header=False).value = df_inv_merged.loc[:, "Proposal Code"]

target_wb.save()
target_wb.close()
source_wb.close()

app.quit()


