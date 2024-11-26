import pandas as pd

# # all_workers = pd.read_excel(r'C:\Users\he kelly\Desktop\TB Setup\weekly report\1008\BCG_-_All_Workers (85).xlsx', sheet_name='BCG - All Workers')
# new_hires = pd.read_excel(r'C:\Users\he kelly\Desktop\TB Setup\weekly report\1008\BCG_-_New_Hires_for_Distribution (58).xlsx')
all_workers = pd.read_excel(r'C:\Users\he kelly\Desktop\TB Setup\New Hires in Batch\1104\BCG_-_All_Workers (88).xlsx', sheet_name='BCG - All Workers')
new_hires = pd.read_excel(r'C:\Users\he kelly\Desktop\TB Setup\New Hires in Batch\1104\New Hires 1104.xlsx')
column_list = ['Employee ID', 'Worker', 'Business Title', 'Hire Date', 'Location', 'Job Family Group', 'Manager', 'Primary Allocation', 'Email - Primary Work']
new_hires = new_hires.merge(all_workers[column_list], how='left', on='Employee ID')
# print(new_hires.columns)

contingent_workers = pd.read_excel(r'C:\Users\he kelly\Downloads\BCG_-_All_Contingent_Workers (28).xlsx')
path_mapping = r'C:\Users\he kelly\Desktop\TB Setup\RPA\New Hires.xlsx'
mapping_BR = pd.read_excel(f'{path_mapping}', sheet_name='1 Title & Billing Rate Per Hour')
mapping_category = pd.read_excel(f'{path_mapping}', sheet_name='5 Category Map').set_index("Job Family Group")
mapping_PA = pd.read_excel(f'{path_mapping}', sheet_name='3 Primary Allocation')
mapping_approval = pd.read_excel(f'{path_mapping}', sheet_name='2 Approval Limits')
mapping_other = pd.read_excel(f'{path_mapping}', sheet_name='4 Other System Set up')

#map contingent worker type
new_hires = new_hires.merge(contingent_workers[["Employee ID", "Contingent Worker Type"]], on='Employee ID', how='left')
job_to_category = mapping_category['Job Category'].to_dict()

#map Job Category
new_hires['Job Category'] = new_hires['Job Family Group'].map(job_to_category)

#map business and cohort
new_hires = new_hires.merge(all_workers[["Employee ID", "Business"]], on='Employee ID', how='left')
new_hires = new_hires.merge(all_workers[["Employee ID", "Cohort/ Cohort Step"]], on='Employee ID', how='left')

#replicate workday iD
new_hires['Workday ID'] = new_hires['Employee ID']

#transform worker name
new_hires['Last Name']=new_hires['Worker'].str.split(' ').str.get(1)
new_hires['First Name']=new_hires['Worker'].str.split(' ').str.get(0)

#set Return ET
new_hires.loc[new_hires['Cohort/ Cohort Step']=='MDP', 'Return ET'] = 'N'

#transform start date
trans_month = {'01': 'JAN', '02': 'FEB', '03': 'MAR', '04': 'APR', '05': 'MAY', '06': 'JUN', '07': 'JUL', '08': 'AUG', '09': 'SEP', '10': 'OCT', '11': 'NOV', '12': 'DEC'}
date = new_hires['Hire Date'].astype(str).str.split('-').str.get(2)
month = new_hires['Hire Date'].astype(str).str.split('-').str.get(1).map(trans_month)
year = new_hires['Hire Date'].astype(str).str.split('-').str.get(0)
new_hires['Start Date'] = date + '-' + month + '-' + year

#map title
mapping_title = mapping_BR[['WD Title', 'TB Title']]
wd_to_tbTitle = mapping_title.loc[mapping_title['WD Title'].notnull()].set_index('WD Title')['TB Title'].to_dict()
mapping_cohortBR = mapping_BR[['Cohort Step', 'TB Title']]
cohort_to_br = mapping_cohortBR[mapping_cohortBR['Cohort Step'].notnull()].set_index('Cohort Step')['TB Title'].to_dict()
new_hires.loc[new_hires['Cohort/ Cohort Step'].notnull(), 'Title'] = new_hires['Cohort/ Cohort Step'].map(cohort_to_br)
new_hires.loc[new_hires['Cohort/ Cohort Step'].isnull(), 'Title'] = new_hires['Cohort/ Cohort Step'].map(wd_to_tbTitle)

#set base hours
new_hires.loc[new_hires['Business Title'] == 'Senior Advisor', 'Base Hours'] = 0
new_hires.loc[new_hires['Business Title'] != 'Senior Advisor', 'Base Hours'] = 40

#set email address
new_hires['Email Address'] = new_hires['Email - Primary Work']

#set authorization limit
cohort_to_authCNY = mapping_approval.set_index('WD Cohort Step')['CNY'].to_dict()
cohort_to_authUSD = mapping_approval.set_index('WD Cohort Step')['USD'].to_dict()
cohort_to_authTWD = mapping_approval.set_index('WD Cohort Step')['TWD'].to_dict()
print(cohort_to_authTWD)
office_CNY = ['Shanghai', 'Shenzhen', 'Beijing']
office_USD = ['Hong Kong']
office_TWD = ['Taipei']
for office in office_CNY:
    new_hires.loc[new_hires['Location'].str.contains(f'{office}'), 'Auth. Limit'] = new_hires.loc[new_hires['Location'].str.contains(f'{office}'), 'Cohort/ Cohort Step'].map(cohort_to_authCNY)
for office in office_USD:
    new_hires.loc[new_hires['Location'].str.contains(f'{office}'), 'Auth. Limit'] = new_hires.loc[new_hires['Location'].str.contains(f'{office}'), 'Cohort/ Cohort Step'].map(cohort_to_authUSD)
for office in office_TWD:
    new_hires.loc[new_hires['Location'].str.contains(f'{office}'), 'Auth. Limit'] = new_hires.loc[new_hires['Location'].str.contains(f'{office}'), 'Cohort/ Cohort Step'].map(cohort_to_authTWD)

#set TS Approval
new_hires.loc[new_hires['Job Category']=='BST', 'TS Approval'] = 'Always'

#set Billing Rate

for currency in ['CNY', 'USD', 'TWD']:
    mapping_titleBR = mapping_BR[['WD Title', f'{currency}']]
    wdTitle_to_br = mapping_titleBR.loc[mapping_titleBR['WD Title'].notnull()].set_index('WD Title')[f'{currency}'].to_dict()
    mapping_cohortBR = mapping_BR[['Cohort Step', f'{currency}']]
    cohort_to_br = mapping_cohortBR[mapping_cohortBR['Cohort Step'].notnull()].set_index('Cohort Step')[f'{currency}'].to_dict()
    mapping_paBR = mapping_BR[['PA', f'{currency}']]
    pa_to_br = mapping_paBR[mapping_paBR['PA'].notnull()].set_index('PA')[f'{currency}'].to_dict()
    if currency == 'CNY':
        list_office = office_CNY
    if currency == 'USD':
        list_office = office_USD
    if currency == 'TWD':
        list_office = office_TWD
    for office in list_office:
        new_hires.loc[(new_hires['Cohort/ Cohort Step'].notnull()) & (new_hires['Location'].str.contains(f'{office}')), 'Billing Rate'] = new_hires.loc[(new_hires['Cohort/ Cohort Step'].notnull()) & (new_hires['Location'].str.contains(f'{office}')), 'Cohort/ Cohort Step'].map(cohort_to_br)
        new_hires.loc[(new_hires['Cohort/ Cohort Step'].isnull()) & (new_hires['Location'].str.contains(f'{office}')), 'Billing Rate'] = new_hires.loc[(new_hires['Cohort/ Cohort Step'].isnull()) & (new_hires['Location'].str.contains(f'{office}')), 'Cohort/ Cohort Step'].map(wdTitle_to_br)
        new_hires.loc[(new_hires['Primary Allocation'].isin(['Data and Research Services_AP', 'GLB Design Studio_AP', 'Language Services_AP'])) & (new_hires['Location'].str.contains(f'{office}')), 'Billing Rate'] = new_hires.loc[(new_hires['Cohort/ Cohort Step'].isin(['Data and Research Services_AP', 'GLB Design Studio_AP', 'Language Services_AP'])) & (new_hires['Location'].str.contains(f'{office}')), 'Primary Allocation'].map(wdTitle_to_br)

#set PA
mapping_PA['CS Theo Cap&CS Rev'] = mapping_PA['CS Theo Cap&CS Rev'].str.strip()
mapping_PA = mapping_PA.set_index('Primary Allocation')
pa_to_tb = mapping_PA['CS Theo Cap&CS Rev'].to_dict()
pa_to_op = mapping_PA['Office/PA'].to_dict()
new_hires['CS Theo Cap&CS Rev'] = new_hires['Primary Allocation'].map(pa_to_tb)
new_hires['Office/PA'] = new_hires['Primary Allocation'].map(pa_to_op)
print(new_hires)

new_hires.to_excel(r'C:\Users\he kelly\Desktop\TB Setup\New Hires in Batch\1104\new hires result 1104.xlsx')
# new_hires.to_excel(r'C:\Users\he kelly\Desktop\TB Setup\weekly report\1025\new hires result 1025.xlsx')
