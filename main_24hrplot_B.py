
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

global pload
global qload
global pgen
global qgen
pload = []
qload = []
pgen = []
qgen = []

global pload1
global qload1
global pgen1
global qgen1
pload1 = []
qload1 = []
pgen1 = []
qgen1 = []

global pload2
global qload2
global pgen2
global qgen2
pload2 = []
qload2 = []
pgen2 = []
qgen2 = []

global pload3
global qload3
global pgen3
global qgen3
pload3 = []
qload3 = []
pgen3 = []
qgen3 = []

global p_val 
p_val = []

global p_val1
p_val1 = []

global p_val2 
p_val2 = []

global p_val3 
p_val3 = []

global pL 
pL = []

global pL1
pL1 = []

global pL2 
pL2 = []

global pL3 
pL3 = []

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
    ierr, pLL = psspy.brnmsc(127,133,'1','PLOS')
    pL.append(pLL)
    p_val.append(p_vall)
    ierr, [tot_pload] = psspy.aareareal(-1,1,'PLOAD')
    ierr, [tot_qload] = psspy.aareareal(-1,1,'QLOAD') 
    ierr, [tot_pgen] = psspy.aareareal(-1,1,'PGEN')
    ierr, [tot_qgen] = psspy.aareareal(-1,1,'QGEN') 
    pload.append(tot_pload[0])
    qload.append(tot_qload[0])
    pgen.append(tot_pgen[0])
    qgen.append(tot_qgen[0])
    
    
psspy.case(CASE1)
for i,j in zip(kl_demand,kl_hydro):
    psspy.scal_4(0,1,0,[_i,_i,_i,1,0,1],[i,j,0.0,-.0,0.0,-.0,0.33*i])
    psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
    ierr, p_vall1 = psspy.brnmsc(127,133,'1','P')
    ierr, pLL1 = psspy.brnmsc(127,133,'1','PLOS')
    pL1.append(pLL1)
    p_val1.append(p_vall1)
    ierr, [tot_pload1] = psspy.aareareal(-1,1,'PLOAD')
    ierr, [tot_qload1] = psspy.aareareal(-1,1,'QLOAD') 
    ierr, [tot_pgen1] = psspy.aareareal(-1,1,'PGEN')
    ierr, [tot_qgen1] = psspy.aareareal(-1,1,'QGEN') 
    pload1.append(tot_pload1[0])
    qload1.append(tot_qload1[0])
    pgen1.append(tot_pgen1[0])
    qgen1.append(tot_qgen1[0])
    
    
psspy.case(CASE2)
for i,j in zip(kl_demand,kl_hydro):
    psspy.scal_4(0,1,0,[_i,_i,_i,1,0,1],[i,j,0.0,-.0,0.0,-.0,0.33*i])
    psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
    ierr, p_vall2 = psspy.brnmsc(127,133,'1','P')
    ierr, pLL2 = psspy.brnmsc(127,133,'1','PLOS')
    pL2.append(pLL2)
    p_val2.append(p_vall2)
    ierr, [tot_pload2] = psspy.aareareal(-1,1,'PLOAD')
    ierr, [tot_qload2] = psspy.aareareal(-1,1,'QLOAD') 
    ierr, [tot_pgen2] = psspy.aareareal(-1,1,'PGEN')
    ierr, [tot_qgen2] = psspy.aareareal(-1,1,'QGEN') 
    pload2.append(tot_pload2[0])
    qload2.append(tot_qload2[0])
    pgen2.append(tot_pgen2[0])
    qgen2.append(tot_qgen2[0])
    
psspy.case(CASE3)
for i,j in zip(kl_demand,kl_hydro):
    psspy.scal_4(0,1,0,[_i,_i,_i,1,0,1],[i,j,0.0,-.0,0.0,-.0,0.33*i])
    psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
    ierr, p_vall3 = psspy.brnmsc(127,133,'1','P')
    ierr, pLL3 = psspy.brnmsc(127,133,'1','PLOS')
    pL3.append(pLL3)
    p_val3.append(p_vall3)
    ierr, [tot_pload3] = psspy.aareareal(-1,1,'PLOAD')
    ierr, [tot_qload3] = psspy.aareareal(-1,1,'QLOAD') 
    ierr, [tot_pgen3] = psspy.aareareal(-1,1,'PGEN')
    ierr, [tot_qgen3] = psspy.aareareal(-1,1,'QGEN') 
    pload3.append(tot_pload3[0])
    qload3.append(tot_qload3[0])
    pgen3.append(tot_pgen3[0])
    qgen3.append(tot_qgen3[0])


df['P'] = p_val
df['P(+100MW)'] = p_val1
df['P(+400MW)'] = p_val2
df['P(+700MW)'] = p_val3

df['PLOS'] = pL
df['PLOS(+100MW)'] = pL1
df['PLOS(+400MW)'] = pL2
df['PLOS(+700MW)'] = pL3

df['PGEN'] = pgen
df['PGEN(+100MW)'] = pgen1
df['PGEN(+400MW)'] = pgen2
df['PGEN(+700MW)'] = pgen3

df['PLOAD'] = pload
df['PLOAD(+100MW)'] = pload1
df['PLOAD(+400MW)'] = pload2
df['PLOAD(+700MW)'] = pload3

df.to_csv('D:\Dev\PSSE Automation\data_P_comp.csv',mode='a',index=False,header=True)

print(df) 
