import openpyxl
import os

wb = openpyxl.load_workbook('Domestic_FITBook_Ver_5.0.7_17FEB2021.xlsm', data_only=True, keep_vba=True)
#print(wb.get_sheet_names())

#ActiveSheet = wb.active         #works
#print(ActiveSheet['C9'].value)   #works

def WhatsNext(x):
    if x == 'PASS':
        print("Hell Yeah")
    elif x == 'FAIL':
        print("Hell No")
    else:
        print("Ok")

#print(os.getcwd())

ExeSumm = wb.worksheets[2]
#Exe5GSumm = wb.worksheets[3]
#ExeCAT1Summ = wb.worksheets[4]
IssueList = wb.worksheets[5]
#print(ExeSumm['C9'].value)

#Marginal VCP
Marg_MO = ExeSumm['C9'].value
WhatsNext(Marg_MO)
Marg_MT = ExeSumm['C10'].value
WhatsNext(Marg_MT)
Marg_LC = ExeSumm['C11'].value
WhatsNext(Marg_LC)

#Mixed VCP
Mixed_MO = ExeSumm['H9'].value
Mixed_MT = ExeSumm['H10'].value
Mixed_LC = ExeSumm['H11'].value

#Data Mobility
Ping_MO = ExeSumm['H15'].value
Data_LC = ExeSumm['H17'].value

