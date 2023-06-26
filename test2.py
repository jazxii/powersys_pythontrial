import os,sys

PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 

import psspy
import pandas as pd
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
print(df3)

print("total P load:",tot_pload[0])
print("total Q load:",tot_qload[0])
print("total P GEN:",tot_pgen[0])
print("total Q GEN:",tot_qgen[0])


