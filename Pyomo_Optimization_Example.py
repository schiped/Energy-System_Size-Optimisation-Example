# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 23:37:58 2022

@author: schiaffipn
"""
import pandas as pd
import pyomo.environ as pyo
import matplotlib.pyplot as plt
import os

#Read Impout File
cwd = os.getcwd()     #get Current Working Directory        
 
input_file = os.path.join(cwd, 'input', "input_file.xlsx")                      #create Input File directory
# input_data = pd.ExcelFile(input_file)                                         #Read input file

#Reading Technical Data Sheet of the file 
df_technical_data = pd.read_excel(input_file, sheet_name="Technical Data", header = 0, usecols=[0,1,2], index_col=(0))

#Reading Irradiance Sheet of the file 
df_irradiance = pd.read_excel(input_file,sheet_name="Irradiance", header = 0, usecols=[0,1])

#Reading Demand Sheet of the file 
df_demand = pd.read_excel(input_file,sheet_name="Demand", header = 0, usecols=[2,3])

#Reading Economic Sheet of the file, use index_col argunment to have the first column as the value of rows..
df_economic = pd.read_excel(input_file,sheet_name="Economic", header = 0, usecols=[0,1,2], index_col=(0))



"PYOMO MODEL"
model = pyo.ConcreteModel(name='PV - Storage Model Optimisation')

#Defining Sets
# model.T = pyo.Set(ordered = True, initialize=range(8760)) ATTENTION: RangeSet starts in 1!, Range starts in 0! 
model.T = pyo.Set(initialize = pyo.RangeSet(len(df_irradiance)), ordered=True) 

model.PV_Costs = pyo.Param(initialize=(df_economic.loc['CAPEX_PV','Input']), mutable=(True))    #In order to perform sensitiviy analisis, the variable has to be saved in a PYOMO Parameter... 
model.Storage_Costs = pyo.Param(initialize=df_economic.loc['CAPEX_Storage','Input'], mutable=(True))

model.x1 = pyo.Var(within=pyo.NonNegativeReals, initialize=(1))                                 # Area PV
model.x2 = pyo.Var(within=pyo.NonNegativeReals, initialize=(1))                                 # Storage Size

model.EPVd = pyo.Var(model.T,within=pyo.NonNegativeReals, initialize=1 )                                                   #Electricity from PV to demand (EPV_used)        
model.SoC = pyo.Var(model.T, within=pyo.NonNegativeReals)
model.B_charge_flow = pyo.Var(model.T, initialize=(0),within=pyo.NonNegativeReals)                                          
model.B_discharge_flow = pyo.Var(model.T, initialize=(0),within=pyo.NonNegativeReals)
model.curtailment = pyo.Var(model.T,within=pyo.NonNegativeReals)

def balance (model,t):
    return model.EPVd[t] + model.B_discharge_flow[t] - df_demand.loc[t-1,'Total Demand'] == 0   #t-1 bc RangeSet starts in 1

def PV_generation (model,t):
    return model.x1*df_irradiance.loc[t-1,'Irradiance']*df_technical_data.loc['PV_Efficiency','Input'] == model.EPVd[t] + model.B_charge_flow [t] + model.curtailment[t]

def SoC_eq1 (model,t):
    # if t == 0:
    if t == 1:
        return model.SoC[t] == 150 # if initial storage is not enough for the sum of the demand in the first hours without irradiance, then the model does not work! in this case the sum is 136.9 kWh #+ (model.charge_flow[t]*df_technical_data.loc['charge_rate','Input']) - (model.discharge_flow[t]/df_technical_data.loc['discharge_rate','Input'])
    else:
        return model.SoC[t] == model.SoC[t-1] + (model.B_charge_flow[t]*df_technical_data.loc['charge_rate','Input']) - (model.B_discharge_flow[t]/df_technical_data.loc['discharge_rate','Input'])

def SoC_eq2 (model, t):
    return model.SoC[t] <= model.x2  

# def SoC_eq3 (model, t):
#     return model.SoC[t] >= model.x2 * 0.1

def discharge_flow (model, t):
    # if t == 0:
    if t == 1:
        return model.B_discharge_flow[t]/df_technical_data.loc['discharge_rate','Input'] == df_demand.loc[t-1,'Total Demand']/df_technical_data.loc['discharge_rate','Input']
    else:
        return model.B_discharge_flow[t]/df_technical_data.loc['discharge_rate','Input'] <= model.SoC[t-1]

def charge_flow (model, t):
    # if t == 0:
    if t == 1:
        return model.B_charge_flow[t]*df_technical_data.loc['charge_rate','Input'] == 0
    else:
        return model.B_charge_flow[t]*df_technical_data.loc['charge_rate','Input'] <= model.x2 - model.SoC[t-1]

model.c1 = pyo.Constraint(model.T, rule=balance)
model.c2 = pyo.Constraint(model.T, rule=PV_generation)
model.c3 = pyo.Constraint(model.T, rule=SoC_eq1)
model.c4 = pyo.Constraint(model.T, rule=SoC_eq2)
model.c5 = pyo.Constraint(model.T, rule=discharge_flow)
model.c6 = pyo.Constraint(model.T, rule=charge_flow)
# model.c7 = pyo.Constraint(model.T, rule=SoC_eq3)

def OF (model):
     return (model.x1*model.PV_Costs) + (model.x2*model.Storage_Costs)

model.OB = pyo.Objective(rule=OF, sense = pyo.minimize)

opt = pyo.SolverFactory('glpk')
results = opt.solve(model, tee=True)
results.write()

print('--------------------------------')
print('Decision Variables: ')
print("PV Area [ha] = ", pyo.value(model.x1))
print("Storage Size [kWh] = ", pyo.value(model.x2))
# model.c1[1].pprint()

print('Objective Function= ',pyo.value(model.OB))


#Saving Results in lists to export to DF, create graphs, etc...:
c1 = []                                                                         # Constraint 1
SoC = []
Battery_charge = []
Battery_discharge = []
PV_production = []
EPV_used = []
PVcurtailment = []

for i in model.T:
        c1.append(pyo.value(model.c1[i]))
        SoC.append(pyo.value(model.SoC[i]))
        Battery_charge.append(pyo.value(model.B_charge_flow[i]))
        Battery_discharge.append(pyo.value(model.B_discharge_flow[i]))
        PV_production.append(pyo.value(model.x1)*df_irradiance.loc[i-1,'Irradiance']*df_technical_data.loc['PV_Efficiency','Input'])
        EPV_used.append(pyo.value(model.EPVd[i]))
        PVcurtailment.append(pyo.value(model.curtailment[i]))
        
x = [t for t in model.T]                                                        # time varialbe for the plots.. 
#Plot 1 week: 
plt.plot(x[0:168],SoC[0:168],label='SoC')
plt.plot(x[0:168],Battery_charge[0:168],label='Charge') 
plt.plot(x[0:168],Battery_discharge[0:168],label='DisCharge')
plt.plot(x[0:168],PV_production[0:168],label='PV')

# plt.xticks(X,XL)
plt.legend()
plt.show()


df_c1 = pd.DataFrame(c1)   
df_SoC = pd.DataFrame(SoC)     
df_Battery_charge = pd.DataFrame(Battery_charge)      
df_Battery_discharge = pd.DataFrame(Battery_discharge)  
df_PV_production = pd.DataFrame(PV_production) 
df_EPV_used = pd.DataFrame(EPV_used) 
df_PVcurtailment = pd.DataFrame(PVcurtailment) 

#Export to excel

with pd.ExcelWriter(' Python output.xlsx') as writer:
    df_c1[0].to_excel(writer, sheet_name="Balance Constraint")
    df_SoC[0].to_excel(writer, sheet_name="SoC")
    df_PV_production[0].to_excel(writer, sheet_name="PV Production")
    df_Battery_charge[0].to_excel(writer, sheet_name="Charge")
    df_Battery_discharge[0].to_excel(writer, sheet_name="Discharge")
    df_EPV_used[0].to_excel(writer, sheet_name="PV")
    df_PVcurtailment[0].to_excel(writer, sheet_name="PV Curtailed")
    
# # Sensitivity Analysis - How the objective changes with for example different PV cost? 
Sensitivityity={}
for variable_PVcosts in range(1500000,5000000,500000): #range(start, stop, step)
    model.PV_Costs=variable_PVcosts         # Making the Parameter Mutuable=True, now I change the variable to perform a sensitivity analysis. 
    results = opt.solve(model) 
    print(variable_PVcosts, pyo.value(model.OB))
    Sensitivityity[variable_PVcosts,'OF']=pyo.value(model.OB)

#Creating Graph:
Y=[Sensitivityity[variable_PVcosts,'OF'] for variable_PVcosts in range(1500000,5000000,500000)]
X=[variable_PVcosts for variable_PVcosts in range(1500000,5000000,500000)]
plt.plot(X,Y)
plt.ylabel('Cost function')
plt.xlabel('CAPEX PV')   
    
