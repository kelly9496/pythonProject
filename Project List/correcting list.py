import pandas as pd
project_list = pd.read_excel(r'C:\Users\he kelly\Desktop\commercial\Project List\Code List 20240527 - kn description updated_filtered - original.xlsx', sheet_name='Sheet1', header=0)
# print(project_list)
pla_list = project_list[project_list['Project Name'].str.contains('PLA ', na=False, case=False) & (project_list['Digital'] == 1)]
print(pla_list)
count = 0
for ind, row in pla_list.iterrows():
    project_name = row['Project Name'][4:]
    client_ID = row['Project ID'][0:6]
    # print(client_ID)
    mapped_project = project_list[project_list['Project ID'].str.contains(f'{client_ID}') & project_list['Project Name'].str.contains(f'{project_name}')]
    print(mapped_project)
    if len(mapped_project) == 2:
        sum_NCC = mapped_project['Net Client Charges (Total)'].sum()
        sum_conValue = mapped_project['Total Contract Amount incl. Expense'].sum()
        # sum_expense = mapped_project['Contracted Expense Amount - Current'].sum()
        mapped_project = mapped_project.drop(ind)
        # print(mapped_project.index)
        project_list.loc[mapped_project.index, 'Net Client Charges (Total)'] = sum_NCC
        project_list.loc[mapped_project.index, 'Total Contract Amount incl. Expense'] = sum_conValue
        # project_list.loc[mapped_project.index, 'Contracted Expense Amount - Current'] = sum_expense
        count += 1
        project_list.loc[mapped_project.index, 'Notes'] = f'{count}'
        project_list.loc[ind, 'Notes'] = f'{count}'
        continue

    if len(mapped_project) == 1:
        for i in range(7, 20):
            project_name = row['Project Name'][4:i]
            mapped_project1 = project_list[project_list['Project ID'].str.contains(f'{client_ID}') & project_list['Project Name'].str.contains(f'{project_name}', case=False, na=False)]
            print(mapped_project1)
            if len(mapped_project1) == 2:
                sum_NCC1 = mapped_project1['Net Client Charges (Total)'].sum()
                sum_conValue1 = mapped_project1['Total Contract Amount incl. Expense'].sum()
                # sum_expense1 = mapped_project1['Contracted Expense Amount - Current'].sum()
                mapped_project1 = mapped_project1.drop(ind)
                # print(mapped_project.index)
                project_list.loc[mapped_project1.index, 'Net Client Charges (Total)'] = sum_NCC1
                project_list.loc[mapped_project1.index, 'Total Contract Amount incl. Expense'] = sum_conValue1
                # project_list.loc[mapped_project1.index, 'Contracted Expense Amount - Current'] = sum_expense
                count += 1
                project_list.loc[mapped_project1.index, 'Notes'] = f'{count}'
                project_list.loc[ind, 'Notes'] = f'{count}'
                break

    if len(mapped_project) == 3:
        print(mapped_project)

print(count)

project_list.to_excel(r'C:\Users\he kelly\Desktop\commercial\Project List\test8.xlsx')