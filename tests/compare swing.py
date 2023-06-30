#data processing in pSSE with IDK gen bus and InterState bus as Swing bus

import os,sys
os.environ["KMP_DUPLICATE_LIB_OK"]="TRUE"

PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 

CASE = r"D:\Dev\PSSE Automation\KERALA_sld_B"
CASE1 = r"D:\Dev\PSSE Automation\KERALA_sld_100_B.sav"
CASE2 = r"D:\Dev\PSSE Automation\KERALA_sld_400_B.sav"
CASE3 = r"D:\Dev\PSSE Automation\KERALA_sld_700_B.sav"

import psspy
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt

psspy.psseinit(2000)

# df = pd.read_csv('D:\Dev\PSSE Automation\KL24hr_data.csv')

psspy.case(CASE)
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )

psspy.case(CASE1)
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )

psspy.case(CASE2)
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
psspy.case(CASE3)
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
