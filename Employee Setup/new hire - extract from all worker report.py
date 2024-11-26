import pandas as pd

new_hire = pd.read_excel(r'C:\Users\he kelly\Desktop\TB Setup\BCG_-_All_Workers (76).xlsx', sheet_name='new hire 0902')
all_worker = pd.read_excel(r'C:\Users\he kelly\Desktop\TB Setup\BCG_-_All_Workers (76).xlsx', sheet_name='BCG - All Workers')

column_list = ['Employee ID', 'Worker', 'Business Title', 'Hire Date', 'Location', 'Job Family Group', 'Manager', 'Primary Allocation', 'Email - Primary Work']

new_hire = new_hire.merge(all_worker[column_list], how='left', on='Employee ID')
new_hire.to_excel(r'C:\Users\he kelly\Desktop\TB Setup\test.xlsx')