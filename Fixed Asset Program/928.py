import pandas as pd

df = pd.DataFrame({'A': [0, 1, 2, 3, 4],
                   'B': [5, 6, 7, 8, 9],
                   'C': ['a', 'b', 'c', 'd', 'e']})

df_new = pd.DataFrame({'C': ['a', 'a', 'd']})
dic_new = {0: 'aa', 2: 'bb', 4: 'cc'}
print(dic_new)
print(df['C'])
df.update(df_new)
# df['C'] = pd.DataFrame.from_dict(dic_new, orient='index')
print(df)


