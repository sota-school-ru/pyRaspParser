import pandas as pd
import openpyxl as xl
path = r'source\\test.xlsx'
wb = xl.load_workbook(path, data_only=True)
n_sheets = len(wb.worksheets)
print(n_sheets)

# !!! After using wb.worksheets return kast sheet
sheet1 = wb.active
sheet1 = wb.worksheets[0]
print(sheet1)
sheet2 = wb.worksheets[1]
print(sheet2)

max_col = sheet1.max_column
print(max_col)
max_row = sheet1.max_row
print(max_row)

for i in range(1, max_row+1):
    for j in range(1, max_col+1):
        cell = sheet1.cell(row=i, column=j )
        print(cell.value, end=' ')
    print()


df = pd.read_excel(path, engine='openpyxl')
print(df.info)

# pip install openpyxl-stubs

'''import xlwings as xw
# connect to book
wb = xw.Book(r'source\test.xlsx') # !!! use raw string - see first letter 'r'

#work with sheets
num_sheets = wb.sheets.count
print(num_sheets)'''
