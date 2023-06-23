import os,sys
PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 


import excelpy
x1 = excelpy.workbook()
x1.show()
x1.set_cell('a1','Bus_Numbers')
x1.set_cell('b1','voltage(pu)')
import psspy
import redirect
redirect.psse2py
psspy.psseinit(10000)

