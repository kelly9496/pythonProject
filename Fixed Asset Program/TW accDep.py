import pandas as pd

register = pd.read_excel(r'C:\Users\he kelly\Desktop\Fixed Assets\FA Register\202305\TW List\累计折旧计算.xlsx', sheet_name='Sheet1', header=0)
# register['累计折旧金额'] = register['资产金额'].astype(int).map(lambda x: x/6515863*620712.47)
print(register['资产金额'].astype('int'))
