import pandas as pd

df_1 = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\df_new.xlsx', sheet_name='Sheet2',
                        header=0)

df_2 = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202311\df_new.xlsx', sheet_name='Sheet3',
                         header=0)
print('df_1', df_1)
print('df_2', df_2)


common_columns = ['B', 'C', 'D']

def merge_common_columns(df_main, df_affix, common_columns):

    rest_columns = df_main.columns.difference(common_columns)
    print(rest_columns)
    list_to_be_mapped = df_affix.index.values.tolist()

    for ind, row in df_main.iterrows():

        print(ind)


        df_bool = df_affix[common_columns] == row[common_columns]
        ind_true = df_bool[df_bool.all(axis=1)].index
        print(ind_true)
        if len(ind_true) == 1:
            if ind_true in list_to_be_mapped:
                print('ind_true in list_to_be_mapped', ind_true in list_to_be_mapped)
                list_to_be_mapped.remove(ind_true)
                ind_true = ind_true.values[0]
            else:
                continue
        elif len(ind_true) > 1:
            ind_to_be_mapped = [x for x in ind_true.values if x in list_to_be_mapped]
            print('ind_to_be_mapped', ind_to_be_mapped)
            if ind_to_be_mapped:
                ind_true = ind_to_be_mapped[0]
                list_to_be_mapped.remove(ind_true)
            else:
                continue
        else:
            continue

        print('final ind_true', ind_true)
        df_main.loc[ind, rest_columns] = df_affix.loc[ind_true, rest_columns].to_dict().values()

    return df_main


df = merge_common_columns(df_1, df_2, common_columns)

print(df)
