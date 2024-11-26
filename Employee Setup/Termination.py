import pandas as pd

termination = pd.read_excel(r'C:\Users\he kelly\Desktop\TB Setup\weekly report\1008\BCG_-_Terminations_for_Distribution (56).xlsx')


#transform worker name
termination['Last Name'] = termination['Worker'].str.split(' ').str.get(1)
termination['First Name'] = termination['Worker'].str.split(' ').str.get(0)
termination['Full Name'] = termination['Last Name'] + ',' + termination['First Name']


#transform start date
trans_month = {'01': 'JAN', '02': 'FEB', '03': 'MAR', '04': 'APR', '05': 'MAY', '06': 'JUN', '07': 'JUL', '08': 'AUG', '09': 'SEP', '10': 'OCT', '11': 'NOV', '12': 'DEC'}
date = termination['Termination Date'].astype(str).str.split('-').str.get(2)
month = termination['Termination Date'].astype(str).str.split('-').str.get(1).map(trans_month)
year = termination['Termination Date'].astype(str).str.split('-').str.get(0)
termination['Term Date'] = date + '-' + month + '-' + year

print(termination['Full Name'])

termination.to_excel(r'C:\Users\he kelly\Desktop\TB Setup\weekly report\1008\BCG_-_Terminations_for_Distribution.xlsx')