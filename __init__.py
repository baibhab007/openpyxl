import openpyxl
print(openpyxl.__version__)
wb = openpyxl.load_workbook('2019 count.xlsx')
print(type(wb))
print(wb.sheetnames)
sheet1 = wb["Sheet1"]
print(type(sheet1))
print(sheet1['A1'].value)

sheet1['A1'].value = 'count total'
wb.save('2019 count updated.xlsx')
