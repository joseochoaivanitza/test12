##

import openpyxl
wb = openpyxl.reader.excel.load_workbook(filename="Pairwise.xlsx")
print(wb.sheetnames)
wb.active = 0
sheet = wb.active
#print(shet['A1].value)
for i in range(1,38):
    print(sheet['A'+str(i)].value,sheet['B'+str(i)].value,sheet['C'+str(i)].value,sheet['D'+str(i)].value,sheet['E'+str(i)].value,sheet['F'+str(i)].value,sheet['G'+str(i)].value,sheet['H'+str(i)].value)
    
