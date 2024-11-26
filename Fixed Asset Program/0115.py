import pandas as pd

df=pd.DataFrame(columns = ['A', 'B', 'C', 'D'])
df.loc[0] = [None, None, 1, 2]
print(df)
empty_columns = df.columns[df.isnull().any()].tolist()
print(empty_columns)
