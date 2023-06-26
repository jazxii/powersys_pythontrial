import os,sys
PSSE_LOCATION = r"C:\Program Files\PTI\PSSE35\35.4\PSSPY39"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 


import pandas as pd

df = pd.read_csv('D:\Dev\PSSE Automation\data.csv')

print(df) 


