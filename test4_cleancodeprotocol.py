
        
global pload
global qload
global pgen
global qgen
pload = []
qload = []
pgen = []
qgen = []

import os,sys
os.environ["KMP_DUPLICATE_LIB_OK"]="TRUE"

PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 

import psspy
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
#--------------------------------
# PSS/E Saved case

def process_data(x):
      ierr, count = psspy.abrncount(-1, 1, 1, 1, 1)
      ierr, [frombus] = psspy.abrnint(-1, 1, 1, 1, 1,'FROMNUMBER')
      ierr, [fromname] = psspy.abrnchar(-1, 1, 1, 1, 1,'FROMNAME')
      ierr, [tobus] = psspy.abrnint(-1, 1, 1, 1, 1,'TONUMBER')
      ierr, [toname] = psspy.abrnchar(-1, 1, 1, 1, 1,'TONAME')
      ierr, [Len] = psspy.abrnreal(-1, 1, 1, 1, 1, 'LENGTH')
      ierr, [p_val] = psspy.abrnreal(-1, 1, 1, 1, 1, 'P')
      ierr, [q_val] = psspy.abrnreal(-1, 1, 1, 1, 1, 'Q')
      ierr, [mva_val] = psspy.abrnreal(-1, 1, 1, 1, 1, 'MVA') 
      ierr, [tot_pload] = psspy.aareareal(-1,1,'PLOAD')
      ierr, [tot_qload] = psspy.aareareal(-1,1,'QLOAD') 
      ierr, [tot_pgen] = psspy.aareareal(-1,1,'PGEN')
      ierr, [tot_qgen] = psspy.aareareal(-1,1,'QGEN') 
      pload.append(tot_pload[0])
      qload.append(tot_qload[0])
      pgen.append(tot_pgen[0])
      qgen.append(tot_qgen[0])
      x['FROM BUS'] = frombus
      x['FROMNAME'] = fromname
      x['TO BUS'] = tobus
      x['TONAME'] = toname
      x['Length'] = Len
      x['P'] = p_val
      x['Q'] = q_val
      x['MVA'] = mva_val
      print(x)
      print("total P load:",tot_pload[0])
      print("total Q load:",tot_qload[0])
      print("total P GEN:",tot_pgen[0])
      print("total Q GEN:",tot_qgen[0])
      

CASE = r"C:\Users\jassi\Downloads\KERALA_sld.sav"
CASE1 = r"C:\Users\jassi\Downloads\KERALA_sld_100.sav"
CASE2 = r"C:\Users\jassi\Downloads\KERALA_sld_400.sav"
CASE3 = r"C:\Users\jassi\Downloads\KERALA_sld_700.sav"

psspy.psseinit(2000)
psspy.case(CASE)
df = pd.DataFrame()
process_data(df)

psspy.case(CASE1)
df1 = pd.DataFrame()
process_data(df1)

psspy.case(CASE2)
df2 = pd.DataFrame()
process_data(df2)

psspy.case(CASE3)
df3 = pd.DataFrame()
process_data(df3)

with pd.ExcelWriter('D:\Dev\PSSE Automation\RE_data.xlsx.xlsx') as writer:  
    df.to_excel(writer, sheet_name='SYS_normal')
    df1.to_excel(writer, sheet_name='SYS+100MW')
    df2.to_excel(writer, sheet_name='SYS+400MW')
    df3.to_excel(writer, sheet_name='SYS+700MW')
    

dff = pd.DataFrame()

dff['PGEN'] = pgen
dff['QGEN'] = qgen
dff['PLOAD'] = pload
dff['QLOAD'] = qload
index = pd.Index(['normal','+100MW - IDK','+400MW - halfPLKD','+700MW - halfPLKD'])
dff.set_index(index)
print(dff)
dff.plot()
# df2["P"].plot(kind = 'hist')
plt.show()
