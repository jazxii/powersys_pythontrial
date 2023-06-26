import os,sys

PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 

import psspy

#--------------------------------
# PSS/E Saved case

CASE = r"C:\Users\jassi\Downloads\KERALA_sld.sav"

if __name__ == '__main__':
  psspy.psseinit(2000)
  psspy.case(CASE)
  psspy.fnsl(
    options1=0, # disable tap stepping adjustment.
    options5=0, # disable switched shunt adjustment.
  )
  
# rarray = psspy.abrnreal(-1)
# print("hi")
# print(psspy.abrnreal(-1,1,1,1,1))  
ierr, xarray = psspy.abrncplx(-1, 1, 1, 1, 1, 'RX')
ierr, rarray = psspy.abrnreal(-1, 1, 1, 1, 1, 'CHARGING')
for i in rarray:
      print(i)