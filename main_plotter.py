import pandas as pd
import matplotlib.pyplot as plt


df = pd.read_csv('D:\Dev\PSSE Automation\data_P_comp.csv')
print(df)


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
#------------------------------------------------------------------------------------------------------

#define subplots
fig2,dx = plt.subplots()

#add first line to plot
dx.plot(df[['KL_DEMAND','KL_HYDRO_GEN']])
dx.legend(['KL_DEMAND','KL_GEN'])


#add x-axis label
dx.set_xlabel('TIME', fontsize=14)

#add y-axis label
dx.set_ylabel('DEMAND & GEN(MW)')

#define second y-axis that shares x-axis with current plot
dx2 = dx.twinx()

#add second line to plot
dx2.plot(df[['PLOS','PLOS(+100MW)','PLOS(+400MW)','PLOS(+700MW)']],linewidth=1)
dx2.legend(['PLOS','PLOS(+100MW)','PLOS(+400MW)','PLOS(+700MW)'],loc = 'lower right')

#add second y-axis label
dx2.set_ylabel('POWER (MW)')
# df[['PLOS','PLOS(+100MW)','PLOS(+400MW)','PLOS(+700MW)']].plot(use_index=True)

#------------------------------------------------------------------------------------------------------

#define subplots
fig2,px = plt.subplots()

#add first line to plot
px.plot(df[['KL_DEMAND','KL_HYDRO_GEN']])
px.legend(['KL_DEMAND','KL_GEN'])


#add x-axis label
px.set_xlabel('TIME', fontsize=14)

#add y-axis label
px.set_ylabel('DEMAND & GEN(MW)')

#define second y-axis that shares x-axis with current plot
px2 = px.twinx()

#add second line to plot
px2.plot(df[['PGEN','PGEN(+100MW)','PGEN(+400MW)','PGEN(+700MW)']],linewidth=1)
px2.legend(['PGEN','PGEN(+100MW)','PGEN(+400MW)','PGEN(+700MW)'],loc = 'lower right')

#add second y-axis label
px2.set_ylabel('POWER (MW)')

plt.show()