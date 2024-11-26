import difflib
import pandas as pd

def string_similar(s1, s2):
    return difflib.SequenceMatcher(None, s1, s2).quick_ratio()

print(string_similar('PNC Bank Corporation', 'Galaxy Entertainment'))

path_file = r'C:\Users\he kelly\Desktop\commercial\Search for CCO.xlsx'
df_project = pd.read_excel(f'{path_file}', sheet_name='Report 1', header=3)
df_client = pd.read_excel(f'{path_file}', sheet_name='Sheet1', header=0)


for ind, row in df_client.iterrows():
    client = row['Company']
    if client != 'Galaxy Entertainment':
        continue
    df_project['Ratio'] = df_project['Client Name-Current'].map(lambda x: string_similar(x, client))
    max_rows = df_project[df_project['Ratio'] == df_project['Ratio'].max()]
    # print(max_rows)

