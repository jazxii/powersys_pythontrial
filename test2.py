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
  