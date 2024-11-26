import pandas as pd

df_mapping = pd.read_excel(r'C:\Users\he kelly\Desktop\bidding\0320\Mapping.xlsx')
df_project = pd.read_excel(r'C:\Users\he kelly\Desktop\bidding\0320\Code List 1220 - client category check - code uncombined.xlsx', header=3)


#filter by code
dic_code = df_mapping.loc[df_mapping['Mapping Category'] == 'Code', 'Mapping'].to_dict()
print(dic_code)
df_byClient = pd.DataFrame()
for ind, code in dic_code.items():
    df = df_project[df_project['Project ID'].str.contains(f'{code}', na=False)]
    df['客户中文名'] = df_mapping.loc[ind, 'CN']
    df_byClient = pd.concat([df_byClient, df])

#filter by name
dic_name = df_mapping.loc[df_mapping['Mapping Category'] == 'Name', 'Mapping'].to_dict()
for ind, name in dic_name.items():
    df = df_project[df_project['Client Name-Current'].str.contains(f'{name}', na=False)]
    df['客户中文名'] = df_mapping.loc[ind, 'CN']
    df_byClient = pd.concat([df_byClient, df])

# print(df_byClient)
# df_byClient.to_excel(r'C:\Users\he kelly\Desktop\bidding\0320\Client Project List.xlsx')

df_client_selected =  df_byClient[(df_byClient['Year'] >= 2019) & (df_byClient['Total Contract Amount incl. Expenses - Current'] > 5000000/7.199)]
df_client_selected.loc[:, '金额'] = '大于500万'
df_client_selected.loc[df_client_selected['Total Contract Amount incl. Expenses - Current'] > 8000000/7.199, '金额'] = '大于800万'
columns_show = ['Year', 'Project ID', 'Project Name', 'Kn Description', 'Billing MDP Name T&B', 'Project Host Office Name', 'Project Actual Start Date', 'Client Name-Current', 'Client Category', 'Client City Name - Current', 'Client Country Name - Current', 'Total Contract Amount incl. Expenses - Current', '客户中文名', '金额']
df_client_selected[columns_show].to_excel(r'C:\Users\he kelly\Desktop\bidding\0320\Selected Client Project List.xlsx')

#filter by MDP
list_MDP = ['ZHOU, HAN', 'CHAN, TED', 'ZHU, HUI', 'HU, LISA', 'WEI, WALES']
df_byMDP = pd.DataFrame()
for mdp in list_MDP:
    df = df_project[df_project['Billing MDP Name T&B'] == f'{mdp}']
    df_byMDP = pd.concat([df_byMDP, df])
df_mdp_selected = df_byMDP[(df_byMDP['Year'] >= 2019) & (df_byMDP['Total Contract Amount incl. Expenses - Current'] > 5000000/7.199)]
df_mdp_selected.loc[:, '金额'] = '大于500万'
df_mdp_selected.loc[df_mdp_selected['Total Contract Amount incl. Expenses - Current'] > 8000000/7.199, '金额'] = '大于800万'
columns_show = ['Year', 'Project ID', 'Project Name', 'Kn Description', 'Billing MDP Name T&B', 'Project Host Office Name', 'Project Actual Start Date', 'Client Name-Current', 'Client Category', 'Client City Name - Current', 'Client Country Name - Current', 'Total Contract Amount incl. Expenses - Current', '金额']
df_mdp_selected[columns_show].to_excel(r'C:\Users\he kelly\Desktop\bidding\0320\Selected MDP Project List.xlsx')

# df_byMDP.to_excel(r'C:\Users\he kelly\Desktop\bidding\0320\MDP Project List.xlsx')