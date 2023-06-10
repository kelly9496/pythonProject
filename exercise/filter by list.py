import pandas as pd
file_path_changeVAT = r"C:\Users\he kelly\Desktop\TW\TW Tax\TAI VAT Adjustment.xlsx"
changeVAT = pd.read_excel(file_path_changeVAT, sheet_name = "Change tax amount")
invoiceNo = changeVAT["其他发票/付款单单号"].dropna().tolist()
invoiceNo = invoiceNo[1:]

file_path_GL = r"C:\Users\he kelly\Desktop\OC\2023\2023.3\TW.xlsx"
gl = pd.read_excel(file_path_GL, sheet_name="Data")
gl.columns=gl.loc[0].tolist()
gl=gl.drop(0)
filteredGL=gl.loc[gl["Invoice Number"].isin(invoiceNo)]
filteredGL.to_excel(r"C:\Users\he kelly\Desktop\TW\TW Tax\filtered gl.xlsx")



# Map = pd.read_excel(file_path, sheet_name = "Map")
# # BSTMap = pd.read_excel(file_path, sheet_name = "BST Map")
# invoice = Map["Invoice Number"].dropna().tolist()
# amount = Map["Amount Avg Rate"].dropna().tolist()
# account = Map["Account Cd"].dropna().tolist
#
# list_account_filtered = pd.DataFrame(columns=list.columns)
# for i in account:
#     filteredList = list.loc[list["Account Cd"] == i, :]
#     list_account_filtered = pd.concat([list_account_filtered, filteredList], axis=0)
#
# list_accInv_filtered = pd.DataFrame(columns=list.columns)
#
# for i in invoice:
#     filteredList = list_account_filtered.loc[list_account_filtered["Invoice Number"] == i, :]
#     list_accInv_filtered = pd.concat([list_accInv_filtered, filteredList], axis=0)
#
#
# list_accInvValue_filtered = pd.DataFrame(columns=list.columns)
# for i in amount:
#     filteredList = list_account_filtered.loc[list_account_filtered["Amount Avg Rate"] == i, :]
#     list_accInvValue_filtered = pd.concat([list_accInvValue_filtered, filteredList], axis=0)
#
# list_accInvValue_filtered.to_excel(r"C:\Users\he kelly\Desktop\TW\TW Audit\2023\Annual audit\TB Invoice Detail v2.xlsx")
#
#
#
