import openpyxl
import os

#print(os.getcwd())

wb = openpyxl.load_workbook('Domestic_FITBook_Ver_5.1.3_14JUN2021.xlsm', data_only=True, keep_vba=True)
#print(wb.get_sheet_names())
#Tracker_Name = []
#Tracker_Value = []
#y = 0


def WhatsNext(x):
    if x == 'PASS':
        print(f"Hell Yeah")
    elif x == 'FAIL':
        print(f"Hell No")
    else:
        print("Ok")


def Info_Page(ExeSumm):
    #Vendor Name
    VendorName = ExeSumm['B3'].value
    #Test Location
    Test_Location = ExeSumm['F3'].value
    #Model Name
    ModelName = ExeSumm['B4'].value
    #Software
    SoftwareName = ExeSumm['F5'].value

#print(os.getcwd())
def Tab_4G_Sum(ExeSumm):
    # Marginal VCP
    Marg_MO = ExeSumm['C9'].value
    WhatsNext(Marg_MO)
    #y = y+1
    Marg_MT = ExeSumm['C10'].value
    WhatsNext(Marg_MT)
    #y = y+1
    Marg_LC = ExeSumm['C11'].value
    WhatsNext(Marg_LC)
    #y = y+1
    # Mixed VCP
    Mixed_MO = ExeSumm['C15'].value
    Mixed_MT = ExeSumm['C16'].value
    Mixed_LC = ExeSumm['C17'].value
    # Data Mobility
    Ping_MO = ExeSumm['C21'].value
    Data_LC = ExeSumm['C23'].value
    # LTE Data Performance
    B13_USB_DL = ExeSumm['G9'].value
    B13_USB_UL = ExeSumm['G10'].value
    B13_MHS_DL = ExeSumm['G11'].value
    B13_MHS_UL = ExeSumm['G12'].value
    B04_USB_DL = ExeSumm['H9'].value
    B04_USB_UL = ExeSumm['H10'].value
    B04_MHS_DL = ExeSumm['H11'].value
    B04_MHS_UL = ExeSumm['H12'].value
    B02_USB_DL = ExeSumm['I9'].value
    B02_USB_UL = ExeSumm['I10'].value
    B02_MHS_DL = ExeSumm['I11'].value
    B02_MHS_UL = ExeSumm['I12'].value
    # Video VCP
    Video_MO = ExeSumm['H28'].value
    Video_MT = ExeSumm['H29'].value
    Video_LC = ExeSumm['H30'].value
    # Intermarket VCP
    Int_Video_MO = ExeSumm['J28'].value
    Int_Video_MT = ExeSumm['J29'].value
    Int_Voice_MO = ExeSumm['J31'].value
    Int_Voice_MT = ExeSumm['J32'].value
    #4G Legacy Features
    Provision4G = ExeSumm['C26'].value
    Sel_Reg4G = ExeSumm['C27'].value
    Call_Feat = ExeSumm['C28'].value
    Sms = ExeSumm['C29'].value
    Data_Serv = ExeSumm['C30'].value
    Concurrent_Ser4G = ExeSumm['C31'].value
    Domestic_Roam = ExeSumm['C32'].value
    #4G VoLTE Features
    InitSetup = ExeSumm['H20'].value
    VvCallEst = ExeSumm['H21'].value
    SuppFea = ExeSumm['H23'].value
    InterTestCase = ExeSumm['H24'].value

def Tab_5G_Sum(Exe5GSumm):
    #5G Features
    Provision5G_FR1 = Exe5GSumm['B9'].value
    Sel_Reg5G_FR1 = Exe5GSumm['B10'].value
    Concurrent_Ser5G_FR1 = Exe5GSumm['B11'].value
    Streaming5G_FR1 = Exe5GSumm['B12'].value
    Interoperability5G_FR1 = Exe5GSumm['B13'].value
    Provision5G_FR2 = Exe5GSumm['C9'].value
    Sel_Reg5G_FR2 = Exe5GSumm['C10'].value
    Concurrent_Ser5G_FR2 = Exe5GSumm['C11'].value
    Streaming5G_FR2 = Exe5GSumm['C12'].value
    Interoperability5G_FR2 = Exe5GSumm['C13'].value
    Dss_sec6_5G_FR2 = Exe5GSumm['C14'].value
    #5G Data Performance
    Dss_5G_DL = Exe5GSumm['C18'].value
    Dss_5G_UL = Exe5GSumm['C19'].value
    Fr2_5G_DL = Exe5GSumm['H18'].value
    Fr2_5G_UL = Exe5GSumm['H19'].value
    Fr2_5G_bidir = Exe5GSumm['H20'].value
    #5G Video Performance
    Voice_Orig5G = Exe5GSumm['H9'].value
    Voice_Term5G = Exe5GSumm['H10'].value
    Voice_Lc5G = Exe5GSumm['H11'].value
    Video_Orig5G = Exe5GSumm['H12'].value
    Video_Term5G = Exe5GSumm['H13'].value
    Video_Lc5G = Exe5GSumm['H14'].value

def Tab_CAT_Sum(ExeCAT1Summ):
    #CAT-M1 Features
    CATM1_GSMA = ExeCAT1Summ['C9'].value
    CATM1_Perform = ExeCAT1Summ['C10'].value
    CATM1_UICC = ExeCAT1Summ['C11'].value
    CATM1_SMS = ExeCAT1Summ['C12'].value
    CATM1_Pwr = ExeCAT1Summ['C13'].value
    #Data Performance
    CATM1_DL = ExeCAT1Summ['B27'].value   #2.14
    CATM1_UL = ExeCAT1Summ['B28'].value   #2.15
    #Mixed LTE Calls
    CATM1_Origs = ExeCAT1Summ['C32'].value  # 2.18
    CATM1_Terms = ExeCAT1Summ['C33'].value  # 2.19
    CATM1_LC = ExeCAT1Summ['C34'].value  # 2.20

def Tab_NB_Sum(ExeNBSumm):
    #CAT-M1 Features
    NB_GSMA = ExeNBSumm['H9'].value
    NB_Perform = ExeNBSumm['H10'].value
    NB_UICC = ExeNBSumm['H11'].value
    NB_SMS = ExeNBSumm['H12'].value
    NB_Pwr = ExeNBSumm['H13'].value
    #Data Performance
    NB_DL = ExeNBSumm['G27'].value   #2.14
    NB_UL = ExeNBSumm['G28'].value   #2.15
    #Mixed LTE Calls
    NB_Origs = ExeNBSumm['H32'].value  # 2.18
    NB_Terms = ExeNBSumm['H33'].value  # 2.19
    NB_LC = ExeNBSumm['H34'].value  # 2.20

Exe4GSumm2 = wb.worksheets[3]
Tab_4G_Sum(Exe4GSumm2)
Exe5GSumm2 = wb.worksheets[4]
Tab_5G_Sum(Exe5GSumm2)
ExeCAT1Summ2 = wb.worksheets[5]
Tab_CAT_Sum(ExeCAT1Summ2)
Tab_NB_Sum(ExeCAT1Summ2)