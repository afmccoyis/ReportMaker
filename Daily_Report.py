import openpyxl
import os

wb = openpyxl.load_workbook('Domestic_FITBook_Ver_5.0.7_17FEB2021.xlsm', data_only=True, keep_vba=True)
print(wb.get_sheet_names())

#ActiveSheet = wb.active         #works
#print(ActiveSheet['C9'].value)   #works


#print(os.getcwd())

ExeSumm = wb.worksheets[2]
Exe5GSumm = wb.worksheets[3]
ExeCAT1Summ = wb.worksheets[4]
IssueList = wb.worksheets[5]

print(ExeSumm['C9'].value)
