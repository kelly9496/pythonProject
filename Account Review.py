import pandas as pd

mdp_list = ['DING, JOSH', 'LUI, VINCENT', 'LIAO, CAROL', 'ZHANG, ALLEN', 'GUO, RICHARD', 'HU, ROGER', 'HAO, CRYSTAL']

list_project = pd.read_excel(r'C:\Users\he kelly\Desktop\Account Review\Report 0423\Code List 0131 + CCO (8).xlsx', sheet_name='Report 1', header=3)
list_AR = pd.read_excel(r'C:\Users\he kelly\Desktop\Account Review\Report 0423\AR Aging.xlsx', sheet_name='AR Aging', header=3)
list_WIP = pd.read_excel(r'C:\Users\he kelly\Desktop\Account Review\Report 0423\WIP 515 for checkingV3 _ 41587105 _ 43573526.xlsx', sheet_name='WIP 515', header=3)
list_invoice = pd.read_excel(r'C:\Users\he kelly\Desktop\Account Review\Report 0423\AR Invoice Listing.xlsx', sheet_name='Total Invoice', header=3)
list_invoice.dropna(subset=['Proj ID'], inplace=True)
list_invoice = list_invoice[list_invoice['Proj ID'] != 'Proj ID']
list_invoice['Total Invoice without VAT'] = list_invoice['Total Invoice'] - list_invoice['VAT']
fx_USD = {'TWD': 32.6137, 'CNY': 7.246, 'USD': 1, 'HKD': 7.8326, 'GBP': 0.8034, 'KRW': 1374.8161, 'EUR': 0.9353}
def trans_to_USD(row):
    inv_currency = row['Inv Crncy Cd']
    inv_amount = row['Total Invoice without VAT']
    amount_USD = inv_amount/fx_USD[f'{inv_currency}']
    return amount_USD

list_invoice['Total Invoice without VAT - USD'] = list_invoice.apply(trans_to_USD, axis=1)
print(list_invoice)
# list_invoice.to_excel(r'C:\Users\he kelly\Desktop\Account Review\test\list_invoice2.xlsx')
# print(list_WIP)
# print(project_list)
selected_projects = pd.DataFrame()
for mdp in mdp_list:
    # if mdp != 'DING, JOSH':
    #     continue
    mdp_projects = list_project.loc[list_project['Billing MDP Name T&B'] == f'{mdp}']
    for ind, row in mdp_projects.iterrows():
        project_id = mdp_projects.loc[ind, 'Project ID']
        contract_value = mdp_projects.loc[ind, 'Total Contract Amount incl. Expenses - Current']
        amount_AR = list_AR.loc[list_AR['Project ID'] == f'{project_id}', 'A R in Office Currency'].sum()
        amount_WIP = list_WIP.loc[list_WIP['Project ID'] == f'{project_id}', 'Total WIP Balance'].sum()
        # amount_billed = list_invoice.loc[list_invoice['Proj ID'] == f'{project_id}', 'Total Invoice without VAT - USD'].sum()
        # pending_billing = contract_value - amount_billed
        print(mdp, project_id, amount_AR, amount_WIP)
        # print('AR', amount_AR)
        # print('WIP', amount_WIP)
        if (amount_AR == 0) and (amount_WIP <= 10):
            mdp_projects.drop(ind, inplace=True)
    selected_projects = pd.concat([selected_projects, mdp_projects])

# selected_projects.to_excel(r'C:\Users\he kelly\Desktop\Account Review\test\selected_projects.xlsx')

print(mdp_projects)