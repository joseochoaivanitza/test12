##

import openpyxl
wb = openpyxl.reader.excel.load_workbook(filename="sample.xlsx")
print(wb.sheetnames)
wb.active = 1
sheet = wb.active
#print(shet['A1].value)
for i in range(1,12):
    print(sheet['A'+str(i)].value,sheet['B'+str(i)].value,sheet['C'+str(i)].value)
    
