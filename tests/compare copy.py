import pandas as pd
import matplotlib.pyplot as plt


df = pd.read_csv('D:\Dev\PSSE Automation\RE_data_T.csv')
print(df)
df.plot()
plt.show()