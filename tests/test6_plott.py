import pandas as pd
import matplotlib.pyplot as plt


df = pd.read_csv('D:\Dev\PSSE Automation\data_P_comp.csv')
print(df.head)

# fig, ax = plt.subplots(figsize=(12,5))
# df[['KL_DEMAND','KL_HYDRO_GEN']].plot(use_index=True)
# df[['P','P(+100MW)','P(+400MW)','P(+700MW)','KL_DEMAND']].plot(use_index=True,secondary_y=True)
# plt.show() 


#define subplots
fig,ax = plt.subplots()

#add first line to plot
ax.plot(df[['KL_DEMAND','KL_HYDRO_GEN']])
ax.legend(['KL_DEMAND','KL_GEN'])


#add x-axis label
ax.set_xlabel('TIME', fontsize=14)

#add y-axis label
ax.set_ylabel('DEMAND & GEN(MW)')

#define second y-axis that shares x-axis with current plot
ax2 = ax.twinx()

#add second line to plot
ax2.plot(df[['P','P(+100MW)','P(+400MW)','P(+700MW)']],linewidth=1)
ax2.legend(['P','P(+100MW)','P(+400MW)','P(+700MW)'],loc = 'lower right')

#add second y-axis label
ax2.set_ylabel('POWER (MW)')
plt.show()