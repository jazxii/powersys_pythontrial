
import os,sys
os.environ["KMP_DUPLICATE_LIB_OK"]="TRUE"

PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 

CASE = r"C:\Users\jassi\Downloads\KERALA_sld.sav"
CASE1 = r"C:\Users\jassi\Downloads\KERALA_sld_100.sav"
CASE2 = r"C:\Users\jassi\Downloads\KERALA_sld_400.sav"
CASE3 = r"C:\Users\jassi\Downloads\KERALA_sld_700.sav"
global p_val 
p_val = []

global p_val1
p_val1 = []

global p_val2 
p_val2 = []

global p_val3 
p_val3 = []
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
    
psspy.case(CASE1)
for i,j in zip(kl_demand,kl_hydro):
    psspy.scal_4(0,1,0,[_i,_i,_i,1,0,1],[i,j,0.0,-.0,0.0,-.0,0.33*i])
    psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
    ierr, p_vall1 = psspy.brnmsc(127,133,'1','P')
    p_val1.append(p_vall1)
    
    
psspy.case(CASE2)
for i,j in zip(kl_demand,kl_hydro):
    psspy.scal_4(0,1,0,[_i,_i,_i,1,0,1],[i,j,0.0,-.0,0.0,-.0,0.33*i])
    psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
    ierr, p_vall2 = psspy.brnmsc(127,133,'1','P')
    p_val2.append(p_vall2)
    
psspy.case(CASE3)
for i,j in zip(kl_demand,kl_hydro):
    psspy.scal_4(0,1,0,[_i,_i,_i,1,0,1],[i,j,0.0,-.0,0.0,-.0,0.33*i])
    psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
    ierr, p_vall3 = psspy.brnmsc(127,133,'1','P')
    p_val3.append(p_vall3)


df['P'] = p_val
df['P(+100MW)'] = p_val1
df['P(+400MW)'] = p_val2
df['P(+700MW)'] = p_val3


df.to_csv('D:\Dev\PSSE Automation\data_P_comp.csv',mode='a',index=True,header=True)

print(df) 

# fig, ax = plt.subplots(figsize=(12,5))
# df[['KL_DEMAND','KL_HYDRO_GEN']].plot(use_index=True)
# df[['P','P(+100MW)','P(+400MW)','P(+700MW)','KL_DEMAND']].plot(use_index=True,secondary_y=True)
# plt.show() 

