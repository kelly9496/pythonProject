import pdfplumber
import pandas as pd

pdf = pdfplumber.open(r"C:\Users\he kelly\Desktop\Fixed Assets\Fixed Asset List\SZ\Nanpeng AV 2.pdf")
pages = pdf.pages

#print(pages)

if len(pages)>=1:
    tables = []
    for each in pages:
        print('111',each)
        table = each.extract_table()
        if table:
            print('find table', table)
            tables.extend(table)
else:
    tables = each.extract_table

data = pd.DataFrame(tables[1:],columns=tables[0])
data
data.to_excel(r"C:\Users\he kelly\Desktop\Fixed Assets\Fixed Asset List\SZ\Nanpeng AV 2.xlsx")

