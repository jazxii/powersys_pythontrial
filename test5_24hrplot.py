
import os,sys
os.environ["KMP_DUPLICATE_LIB_OK"]="TRUE"

PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 

CASE = r"C:\Users\jassi\Downloads\KERALA_sld.sav"
global p_val 
p_val = []

import psspy
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt

psspy.psseinit(2000)
psspy.case(CASE)
df = pd.read_csv('D:\Dev\PSSE Automation\KL24hr_data.csv')
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )

kl_demand = df['KL_DEMAND'].to_numpy()
kl_hydro = df['KL_HYDRO_GEN'].to_numpy()
timestamp = df['timestamp'].to_numpy()

_i = psspy.getdefaultint()


for i,j in zip(kl_demand,kl_hydro):
    psspy.scal_4(0,1,0,[_i,_i,_i,1,0,1],[i,j,0.0,-.0,0.0,-.0,0.33*i])
    psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
    ierr, p_vall = psspy.brnmsc(127,133,'1','P')
    p_val.append(p_vall)
    print(p_vall)

df['P'] = p_val
df.set_index(timestamp)

print(df) 
df[['P','KL_DEMAND']].plot(use_index=True)
plt.show()