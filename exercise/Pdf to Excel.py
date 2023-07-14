import pdfplumber
import pandas as pd

pdf = pdfplumber.open(r"C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\SH JAN Reimbursement.pdf")
pages = pdf.pages

#print(pages)

if len(pages)>=1:
    text_all = ''
    for each in pages:
        print('111',each)
        text = each.extract_text()
        if text:
            print('find table', type(text))
            text_all = text_all + text

print(text_all)
# else:
#     tables = each.extract_table

#
#
# tables.to_excel(r"C:\Users\he kelly\Desktop\Alteryx & Python\Bank Rec Program\SH JAN Reimbursement.pdf")

