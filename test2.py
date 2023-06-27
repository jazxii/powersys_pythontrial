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

CASE = r"C:\Users\jassi\Downloads\KERALA_sld.sav"
CASE1 = r"C:\Users\jassi\Downloads\KERALA_sld_100.sav"
CASE2 = r"C:\Users\jassi\Downloads\KERALA_sld_400.sav"
CASE3 = r"C:\Users\jassi\Downloads\KERALA_sld_700.sav"

if __name__ == '__main__':
  psspy.psseinit(2000)
  psspy.case(CASE)
  psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
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
pload =[]
qload = []
pgen = []
qgen = []

pload.append(float(tot_pload[0]))
qload.append(tot_qload[0])
pgen.append(tot_pgen[0])
qgen.append(tot_qgen[0])



df = pd.DataFrame()
df['FROM BUS'] = fromname
df['FROMNAME'] = fromname
df['TO BUS'] = tobus
df['TONAME'] = toname
df['Length'] = Len
df['P'] = p_val
df['Q'] = q_val
df['MVA'] = mva_val
df.to_csv('D:\Dev\PSSE Automation\data.csv',mode='a',index=False,header=False)
# df.to_excel('D:\Dev\PSSE Automation\RE_data.xlsx',sheet_name='SYS_normal')
print(df)



print("total P load:",tot_pload[0])
print("total Q load:",tot_qload[0])
print("total P GEN:",tot_pgen[0])
print("total Q GEN:",tot_qgen[0])


psspy.case(CASE1)
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
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


df1 = pd.DataFrame()
df1['FROM BUS'] = fromname
df1['FROMNAME'] = fromname
df1['TO BUS'] = tobus
df1['TONAME'] = toname
df1['Length'] = Len
df1['P'] = p_val
df1['Q'] = q_val
df1['MVA'] = mva_val
df1.to_csv('D:\Dev\PSSE Automation\data1.csv',mode='a',index=False,header=False)
# with pd.ExcelWriter('D:\Dev\PSSE Automation\RE_data.xlsx',mode='a') as writer:  
#     df1.to_excel(writer, sheet_name='SYS+100MW')
print(df1)

print("total P load:",tot_pload[0])
print("total Q load:",tot_qload[0])
print("total P GEN:",tot_pgen[0])
print("total Q GEN:",tot_qgen[0])


psspy.case(CASE2)
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
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


df2 = pd.DataFrame()
df2['FROM BUS'] = fromname
df2['FROMNAME'] = fromname
df2['TO BUS'] = tobus
df2['TONAME'] = toname
df2['Length'] = Len
df2['P'] = p_val
df2['Q'] = q_val
df2['MVA'] = mva_val
df2.to_csv('D:\Dev\PSSE Automation\data2.csv',mode='a',index=False,header=False)
# with pd.ExcelWriter('D:\Dev\PSSE Automation\RE_data.xlsx',mode='a') as writer:  
#     df2.to_excel(writer, sheet_name='SYS+400MW')
print(df2)

print("total P load:",tot_pload[0])
print("total Q load:",tot_qload[0])
print("total P GEN:",tot_pgen[0])
print("total Q GEN:",tot_qgen[0])


psspy.case(CASE3)
psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
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


df3 = pd.DataFrame()
df3['FROM BUS'] = fromname
df3['FROMNAME'] = fromname
df3['TO BUS'] = tobus
df3['TONAME'] = toname
df3['Length'] = Len
df3['P'] = p_val
df3['Q'] = q_val
df3['MVA'] = mva_val
df3.to_csv('D:\Dev\PSSE Automation\data3.csv',mode='a',index=False,header=False)
# with pd.ExcelWriter('D:\Dev\PSSE Automation\RE_data.xlsx',mode='a') as writer:  
#     df3.to_excel(writer, sheet_name='SYS+700MW')
print(df3)

print("total P load:",tot_pload[0])
print("total Q load:",tot_qload[0])
print("total P GEN:",tot_pgen[0])
print("total Q GEN:",tot_qgen[0])



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
plt.show()

# dff.to_csv('D:\Dev\PSSE Automation\data2.csv',mode='a',index=False,header=False)

