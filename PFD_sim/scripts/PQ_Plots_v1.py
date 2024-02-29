# -*- coding: utf-8 -*-
"""
Created on Fri Dec  1 16:06:01 2023

@author: OX2 - Michael Magpantay

"""
import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import time
#import cmath
timestr=time.strftime("%Y%m%d-%H%M%S")
start_time = datetime.now()


main_folder_path=os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
sys.path.append(main_folder_path+"\\PowerFactory_sim\\scripts\\Libs")
import readtestinfo
#==============================================================================
#USER INPUTS
#Allign these inputs with the PQ Study script
#==============================================================================
TestDefinitionSheet=r'20231117_SUM_TESTINFO_V1.xlsx'
raw_PQ_result_folder = '20240114-143635_PQ'
simulation_batch_label = 'PQ' # same as PQ analysis script


# Create result folder
def make_dir(dir_path, dir_name=""):
    dir_make = os.path.join(dir_path, dir_name)
    try:
        os.mkdir(dir_make)
    except OSError:
        pass
    
dir_path =  main_folder_path +"\\Plots\\PQ"
make_dir(dir_path)


###############################################################################
# Define Project Paths
###############################################################################
try:
    testRun = timestr + '_' + simulation_batch_label #define a test name for the batch or configuration that is being tested -> link to time stamp for auto update
except:
    testRun = timestr            
            
current_dir = os.path.dirname(__file__) #directory of the script
#current_dir=r"C:\Users\Dao Vu\ESCO Pacific\ESCO - Projects\19. LAN\3. Grid\1. Power System Studies\1. Main Test Environment\20220318_LSF\PSSE_sim\scripts"
main_folder = os.path.dirname(current_dir) # Identify main_folder: to be compatible with previous version. / main folder is one level above

# Locating the existing folders
testDefinitionDir= os.path.abspath(os.path.join(main_folder, os.pardir))+"\\test_scenario_definitions"
inputInforFile = testDefinitionDir+"\\"+TestDefinitionSheet

# Directory to store PQ result
outputResultPath=dir_path+"\\"+testRun
make_dir(outputResultPath)

###############################################################################
# Model Information
###############################################################################
# SteadyStateDict =  readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet)
#import readtestinfo as readtestinfo
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ModelDetailsPF', 'PowerQuality','PQ Limits'])
PQDict = return_dict['PowerQuality']
ModelDetailsPF = return_dict['ModelDetailsPF']
lmt_Plan = return_dict['Planning Limits']
lmt_alloc_AAS = return_dict['Allocation Limits_AAS']
lmt_alloc_MAS = return_dict['Allocation Limits_MAS']
Bkg_Harm = return_dict['Bkg_Harm']
del Bkg_Harm['THD']

result_sheet_path = main_folder +"\\result_data\\PQ\\" + raw_PQ_result_folder + "\\" + raw_PQ_result_folder + ".xlsx"
cases=[]
calc_dfs={}
for case in PQDict.keys():
    if PQDict[case][0]['Run?']==1:
        cases.append(case)
        
cases_with_filters=[]
cases_wo_filters=[]

for case in cases:
    if case[4]=='a':
        cases_wo_filters.append(case)
    if case[4]=='b':
        cases_with_filters.append(case)

# =============================================================================
# This function uses the PQ analysis results excel file and plots the results
# # ===========================================================================
from functools import reduce

def read_result_sheet(casegroup):
    sheets_dict  = pd.read_excel(result_sheet_path, sheet_name = None)
    summary_dict = pd.read_excel(result_sheet_path, sheet_name = 'POC LF_THD')
    for i in range(len(summary_dict)):
        if summary_dict['Case name'][i] not in casegroup:
            summary_dict = summary_dict.drop(i)
    summary_dict = summary_dict.reset_index(drop=True)
    cases_dfs={}
    for case in casegroup:
        for name,sheet in sheets_dict.items():
            if name == case + "_Zequiv": #name == case + "_Zequiv":
                Zequiv_df = sheets_dict[case + "_Zequiv"] #sheets_dict[case + "_Zequiv"]
                Zequiv_df.rename(columns = {'HD in %':'HD_equiv','Network Resistance in Ohm':'R_equiv','Network Reactance in Ohm':'X_equiv','Network Impedance, Magnitude in Ohm':'Z_equiv','Network Impedance, Angle in deg':'phi_equiv'}, inplace = True) 
            if name == case + "_Zext": #name == case + "_Zext":
                Zext_df = sheets_dict[case + "_Zext"] #sheets_dict[case + "_Zext"]
                Zext_df.rename(columns = {'HD in %':'HD_ext','Network Resistance in Ohm':'R_ext','Network Reactance in Ohm':'X_ext','Network Impedance, Magnitude in Ohm':'Z_ext','Network Impedance, Angle in deg':'phi_ext'}, inplace = True)  
            if name == case + "_Zgen": #name == case + "_Zgen":
                Zgen_df = sheets_dict[case + "_Zgen"] #sheets_dict[case + "_Zgen"]
                Zgen_df.rename(columns = {'HD in %':'HD_gen','Network Resistance in Ohm':'R_gen','Network Reactance in Ohm':'X_gen','Network Impedance, Magnitude in Ohm':'Z_gen','Network Impedance, Angle in deg':'phi_gen'}, inplace = True)          
            else:
                pass
        cases_dfs[case]=reduce(lambda  left,right: pd.merge(left,right,on=['h'],how='outer'), [Zequiv_df,Zext_df,Zgen_df])
    return cases_dfs, summary_dict

def calc_plant_emission(casegroup):
    plant_emissions=dict.fromkeys(globals().get(casegroup,None))
    for case in plant_emissions.keys():
        plant_emissions[case]=pd.DataFrame(index=range(0,50),columns=['h','HD in %'])
        plant_emissions[case]['h'][0]='THD'
    for case in globals().get(casegroup,None):
        for i in range(len(Bkg_Harm)):
            plant_emissions[case]['h'][i+1]=str(i+2)
            if PFresults[casegroup]['HD & Impedances'][case]['HD_equiv'][i+1] - Bkg_Harm[i+2] <0:
                plant_emissions[case]['HD in %'][i+1] = 0
            else:
                plant_emissions[case]['HD in %'][i+1] = pow(pow(PFresults[casegroup]['HD & Impedances'][case]['HD_equiv'][i+1],return_dict['Alpha Factors'][i+2]) - pow(Bkg_Harm[i+2],return_dict['Alpha Factors'][i+2]),1/return_dict['Alpha Factors'][i+2])   
            #plant_emissions[case]['HD in %'][i+1]=cases_dfs[case]['HD_equiv'][i+1] - Bkg_Harm[i+2]
            plant_emissions[case]['HD in %'][i+1]=float("{:.2f}".format(plant_emissions[case]['HD in %'][i+1]))
        Sum_sq = 0
        for i in range(len(plant_emissions[case])-1):
            HD_sq = pow(plant_emissions[case]['HD in %'][i+1],2)
            Sum_sq += HD_sq
        THD = pow(Sum_sq,0.5)
        THD=float("{:.2f}".format(THD))
        plant_emissions[case]['HD in %'][0]=THD
    return plant_emissions

def total_emission(casegroup):
    total_emissions=dict.fromkeys(globals().get(casegroup,None))
    for case in total_emissions.keys():
        total_emissions[case]=pd.DataFrame(index=range(0,50),columns=['h','HD in %'])
        total_emissions[case]['h'][0]='THD'
        THD = PFresults[casegroup]['Summary'].loc[PFresults[casegroup]['Summary']['Case name']==case,'THD %'].iloc[0]
        total_emissions[case]['HD in %'][0]=THD
        for i in range(len(PFresults[casegroup]['HD & Impedances'][case])-1):
            total_emissions[case]['h'][i+1]=str(i+2)
            total_emissions[case]['HD in %'][i+1]=PFresults[casegroup]['HD & Impedances'][case]['HD_equiv'][i+1]
    return total_emissions

def amplification_factors(casegroup):
    amp_factor=dict.fromkeys(globals().get(casegroup,None))
    for case in amp_factor.keys():
        amp_factor[case]=pd.DataFrame(index=range(0,49),columns=['h','Amplification Factor'])
        for i in range(len(PFresults[casegroup]['HD & Impedances'][case])-1):
            amp_factor[case]['h'][i]=str(i+2)
            #Zequiv=complex(cases_dfs[case]['R_equiv'][i+1],cases_dfs[case]['X_equiv'][i+1])
            Zext=complex(PFresults[casegroup]['HD & Impedances'][case]['R_ext'][i+1],PFresults[casegroup]['HD & Impedances'][case]['X_ext'][i+1])
            Zgen=complex(PFresults[casegroup]['HD & Impedances'][case]['R_gen'][i+1],PFresults[casegroup]['HD & Impedances'][case]['X_gen'][i+1])
            AF= abs(Zgen/(Zext+Zgen))
            amp_factor[case]['Amplification Factor'][i]=AF
    return amp_factor

def allocation_limit_chart(casegroup):
    for case in calcresults[casegroup]['Plant Emissions'].keys():
        x = calcresults[casegroup]['Plant Emissions'][case]['h'].tolist()
        y = calcresults[casegroup]['Plant Emissions'][case]['HD in %'].tolist()
        AAS=list(lmt_alloc_AAS.values()) 
        MAS=list(lmt_alloc_MAS.values())
        fig, ax = plt.subplots(figsize=(20,10))
        ax.bar(x, y, width=0.8, edgecolor="white", linewidth=0.7,zorder=1,label=ModelDetailsPF['Project Name']+ " Emissions")
        ax.scatter(x, AAS,s=250,color='orange',marker='_',zorder=2,label='Allocation Limits - AAS')
        ax.scatter(x, MAS,s=250,color='red',marker='_',zorder=2,label='Allocation Limits - MAS')
        ax.grid(axis='y',zorder=0)
        ax.set(xlim=(-1, 50), xticks=x,
               ylim=(0, max(MAS)+1.5), yticks=np.arange(1, max(MAS)+1.5))
        plt.xlabel('Harmonic Order')
        plt.ylabel('Harmonic Distortion % of fundamental')
        plt.legend()
        plt.title(case + " - " + ModelDetailsPF['Project Name']+ " Emissions")
        plt.show()

def planning_limit_chart(casegroup):
    for case in calcresults[casegroup]['Total Emissions'].keys():
        x = calcresults[casegroup]['Total Emissions'][case]['h'].tolist()
        y = calcresults[casegroup]['Total Emissions'][case]['HD in %'].tolist()
        Plan_limit_list=list(lmt_Plan.values()) 
        fig, ax = plt.subplots(figsize=(20,10))
        ax.bar(x, y, width=0.8, edgecolor="white", color='orange', linewidth=0.7,zorder=1,label='Total Emissions')
        ax.scatter(x, Plan_limit_list,s=250,color='green',marker='_',zorder=2,label='Planning Limits')
        ax.grid(axis='y',zorder=0)
        ax.set(xlim=(-1, 50), xticks=x,
               ylim=(0, max(Plan_limit_list)+2), yticks=np.arange(1, max(Plan_limit_list)+2))
        plt.xlabel('Harmonic Order')
        plt.ylabel('Harmonic Distortion % of fundamental')
        plt.legend()
        plt.title(case + " - " + "Total Emissions")
        plt.show()    
    
def AF_chart(casegroup):
    for case in calcresults[casegroup]['Amplification Factors'].keys():
        x = calcresults[casegroup]['Amplification Factors'][case]['h'].tolist()
        y = calcresults[casegroup]['Amplification Factors'][case]['Amplification Factor'].tolist()
        fig, ax = plt.subplots(figsize=(20,10))
        ax.plot(x, y,color='red',zorder=2)
        ax.grid()
        ax.set(xlim=(1, 48), xticks=x,
               ylim=(0, 2), yticks=np.arange(1,max(y)+1))
        plt.xlabel('Harmonic Order')
        plt.ylabel('Amplification Factor')
        plt.title(case + " - " + "Amplification Factors")
        plt.show()    

def AF_allcases(casegroup):
    AF_all=pd.DataFrame(index=range(0,49),columns=['h'])
    for i in range(0,49):
        AF_all['h'][i]=str(i+2)     
    for case in calcresults[casegroup]['Amplification Factors'].keys():
        df=calcresults[casegroup]['Amplification Factors'][case].copy(deep=True)
        df.rename(columns = {'Amplification Factor':case}, inplace = True)
        AF_all=reduce(lambda  left,right: pd.merge(left,right,on=['h'],how='outer'), [AF_all,df])
    AF_all=AF_all.set_index('h')  
    AF_all = AF_all.apply(pd.to_numeric, errors='coerce')
    minmax_AF = pd.DataFrame()
    minmax_AF['Max AF'] = AF_all.max(axis=1)
    minmax_AF['Max AF Case'] = AF_all.idxmax(axis=1)
    minmax_AF['Min AF'] = AF_all.min(axis=1)
    minmax_AF['Min AF Case'] = AF_all.idxmin(axis=1)
    return AF_all, minmax_AF

def AF_minmax_chart(casegroup):
    x = calcresults[casegroup]['Min Max AF'].index.tolist()
    y1 = calcresults[casegroup]['Min Max AF']['Max AF'].tolist()
    y2 = calcresults[casegroup]['Min Max AF']['Min AF'].tolist()
    fig, ax = plt.subplots(figsize=(20,10))
    ax.grid()
    ax.fill_between(x, y1, y2, alpha=0.8, linewidth=0)
    ax.set(xlim=(1, 48), xticks=x,
           ylim=(0, 2), yticks=np.arange(1,max(y1)+1))
    plt.xlabel('Harmonic Order')
    plt.ylabel('Min/Max Amplification Factor')
    if casegroup == 'cases_wo_filters':
        titlesuffix = "All Scenarios without Filters"
    if casegroup == 'cases_with_filters':
        titlesuffix = "All Scenarios with Filters"   
    plt.title("Amplification Factor Summary - " + titlesuffix)
    plt.show()

def maxtotalemission_allcases(casegroup):
    TE_all=pd.DataFrame(index=range(0,50),columns=['h'])
    TE_all['h'][0]='THD'
    for i in range(0,49):
        TE_all['h'][i+1]=str(i+2)     
    for case in calcresults[casegroup]['Total Emissions'].keys():
        df=calcresults[casegroup]['Total Emissions'][case].copy(deep=True)
        df.rename(columns = {'HD in %':case}, inplace = True)
        TE_all=reduce(lambda  left,right: pd.merge(left,right,on=['h'],how='outer'), [TE_all,df])
    TE_all=TE_all.set_index('h')  
    TE_all = TE_all.apply(pd.to_numeric, errors='coerce')
    max_TE = pd.DataFrame()
    max_TE['Max Total Emission'] = TE_all.max(axis=1)
    max_TE['Max Total Emission Case'] = TE_all.idxmax(axis=1)
    return TE_all, max_TE

def maxtotalemission_chart(casegroup):
    x = calcresults[casegroup]['Max Total Emission'].index.tolist()
    y = calcresults[casegroup]['Max Total Emission']['Max Total Emission'].tolist()
    Plan_limit_list=list(lmt_Plan.values()) 
    fig, ax = plt.subplots(figsize=(20,10))
    ax.bar(x, y, width=0.8, edgecolor="white", color='orange', linewidth=0.7,zorder=1,label='Total Emissions')
    ax.scatter(x, Plan_limit_list,s=250,color='green',marker='_',zorder=2,label='Planning Limits')
    ax.grid(axis='y',zorder=0)
    ax.set(xlim=(-1, 50), xticks=x,
           ylim=(0, max(Plan_limit_list)+4), yticks=np.arange(1, max(Plan_limit_list)+4))
    plt.xlabel('Harmonic Order')
    plt.ylabel('Harmonic Distortion % of fundamental')
    plt.legend()
    if casegroup == 'cases_wo_filters':
        titlesuffix = "All Scenarios without Filters"
    if casegroup == 'cases_with_filters':
        titlesuffix = "All Scenarios with Filters"   
    plt.title("Maximum Total Emissions - " + titlesuffix)
    plt.show()    

def maxplantemission_allcases(casegroup):
    PE_all=pd.DataFrame(index=range(0,50),columns=['h'])
    PE_all['h'][0]='THD'
    for i in range(0,49):
        PE_all['h'][i+1]=str(i+2)     
    for case in calcresults[casegroup]['Plant Emissions'].keys():
        df=calcresults[casegroup]['Plant Emissions'][case].copy(deep=True)
        df.rename(columns = {'HD in %':case}, inplace = True)
        PE_all=reduce(lambda  left,right: pd.merge(left,right,on=['h'],how='outer'), [PE_all,df])
    PE_all=PE_all.set_index('h')  
    PE_all = PE_all.apply(pd.to_numeric, errors='coerce')
    max_PE = pd.DataFrame()
    max_PE['Max Plant Emission'] = PE_all.max(axis=1)
    max_PE['Max Plant Emission Case'] = PE_all.idxmax(axis=1)
    return PE_all, max_PE

def maxplantemission_chart(casegroup):
    x = calcresults[casegroup]['Max Plant Emission'].index.tolist()
    y = calcresults[casegroup]['Max Plant Emission']['Max Plant Emission'].tolist()
    AAS=list(lmt_alloc_AAS.values()) 
    MAS=list(lmt_alloc_MAS.values())
    fig, ax = plt.subplots(figsize=(20,10))
    ax.bar(x, y, width=0.8, edgecolor="white", linewidth=0.7,zorder=1,label=ModelDetailsPF['Project Name']+ " Emissions")
    ax.scatter(x, AAS,s=250,color='orange',marker='_',zorder=2,label='Allocation Limits - AAS')
    ax.scatter(x, MAS,s=250,color='red',marker='_',zorder=2,label='Allocation Limits - MAS')
    ax.grid(axis='y',zorder=0)
    ax.set(xlim=(-1, 50), xticks=x,
           ylim=(0, max(MAS)+1.5), yticks=np.arange(1, max(MAS)+1.5))
    plt.xlabel('Harmonic Order')
    plt.ylabel('Harmonic Distortion % of fundamental')
    plt.legend()
    if casegroup == 'cases_wo_filters':
        titlesuffix = "All Scenarios without Filters"
    if casegroup == 'cases_with_filters':
        titlesuffix = "All Scenarios with Filters" 
    plt.title("Max " + ModelDetailsPF['Project Name']+ " Emissions - " +titlesuffix)
    plt.show()
 
def maxHD_impedances(casegroup):
    worstHD_case=calcresults[casegroup]['Max Total Emission']['Max Total Emission Case']
    maxHD_Z=pd.DataFrame(index=range(0,49),columns=['h','R_ext','X_ext','Case'])
    for i in range(0,49):
        maxHD_Z['h'][i]=str(i+2)
    for i in range(len(worstHD_case)-1):
        maxHD_Z['R_ext'][i]=PFresults[casegroup]['HD & Impedances'][worstHD_case[i+1]]['R_ext'][i+1]
        maxHD_Z['X_ext'][i]=PFresults[casegroup]['HD & Impedances'][worstHD_case[i+1]]['X_ext'][i+1]
        maxHD_Z['Case'][i]=worstHD_case[i+1]
    return maxHD_Z

#For debugging only
def Zext_points(casegroup):
    Zext_allcases=pd.DataFrame(index=range(0, 49), columns=['h'])
    for i in range(0, 49):
        Zext_allcases['h'][i]=str(i+2) 
    for case in PFresults[casegroup]['HD & Impedances'].keys():
        df1=PFresults[casegroup]['HD & Impedances'][case]['R_ext'].copy(deep=True)
        #df1=df1.drop(0)
        df1=df1.to_frame()
        #df1.rename(columns = {'R_ext':case + '_Rext'}, inplace = True)
        df2=PFresults[casegroup]['HD & Impedances'][case]['X_ext'].copy(deep=True)
        #df2=df2.drop(0)
        df2=df2.to_frame()
        #df2.rename(columns = {'X_ext':case + '_Xext'}, inplace = True)
        Zext_allcases[case + '_Rext']=''
        Zext_allcases[case + '_Xext']=''
        for i in range(0,49):
            Zext_allcases[case + '_Rext'][i]=df1['R_ext'][i+1]
            Zext_allcases[case + '_Xext'][i]=df2['X_ext'][i+1]
    Zext_allcases = Zext_allcases.set_index('h')
    Zext_allcases = Zext_allcases.apply(pd.to_numeric)
    Zext_allcases_pu = Zext_allcases.copy(deep=True)
    Zbase = 174.24 #Zbase of polygon
    Zext_allcases_pu = Zext_allcases_pu/Zbase
    column_names = []
    for i in range(2,51):
        column_names.append(str(i)+'_Rext')
        column_names.append(str(i)+'_Xext')
    Zext_allcases_polygon = pd.DataFrame(index=range(0,int(Zext_allcases_pu.shape[1]/2)), columns=column_names)
    Zext_allcases_polygon = Zext_allcases_polygon.apply(pd.to_numeric)
    for case in PFresults[casegroup]['HD & Impedances'].keys():
        #for i in range(0,int(Zext_allcases_pu.shape[1]/2)):
        x=0
        while x<49:
            Zext_allcases_polygon[str(x+2) + '_Rext'][list(PFresults[casegroup]['HD & Impedances'].keys()).index(case)] = Zext_allcases_pu[case + '_Rext'][x]
            Zext_allcases_polygon[str(x+2) + '_Xext'][list(PFresults[casegroup]['HD & Impedances'].keys()).index(case)] = Zext_allcases_pu[case + '_Xext'][x]
            x += 1      
    return Zext_allcases, Zext_allcases_pu, Zext_allcases_polygon

# =============================================================================
#Results from PowerFactory
PFresults = {'cases_wo_filters':{},'cases_with_filters':{}}
PFresults['cases_wo_filters']['HD & Impedances'],PFresults['cases_wo_filters']['Summary'] = read_result_sheet(cases_wo_filters)
PFresults['cases_with_filters']['HD & Impedances'],PFresults['cases_with_filters']['Summary'] = read_result_sheet(cases_with_filters)

#Plant Emissions
calcresults = {'cases_wo_filters':{},'cases_with_filters':{}}
calcresults['cases_wo_filters']['Plant Emissions'] = calc_plant_emission(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['Plant Emissions'] = calc_plant_emission(f'{cases_with_filters=}'.split('=')[0])

#Total Emissions
calcresults['cases_wo_filters']['Total Emissions'] = total_emission(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['Total Emissions'] = total_emission(f'{cases_with_filters=}'.split('=')[0])

#Amplification Factors
calcresults['cases_wo_filters']['Amplification Factors'] = amplification_factors(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['Amplification Factors'] = amplification_factors(f'{cases_with_filters=}'.split('=')[0])

#Individual Planning Limit Chart
planning_limit_chart(f'{cases_wo_filters=}'.split('=')[0])
planning_limit_chart(f'{cases_with_filters=}'.split('=')[0])

#Individual Allocation Limit Chart
allocation_limit_chart(f'{cases_wo_filters=}'.split('=')[0])
allocation_limit_chart(f'{cases_with_filters=}'.split('=')[0])

#Individual Amplification Factor Chart
AF_chart(f'{cases_wo_filters=}'.split('=')[0])
AF_chart(f'{cases_with_filters=}'.split('=')[0])

#Min/Max Amplification Factor
calcresults['cases_wo_filters']['AF Merged'], calcresults['cases_wo_filters']['Min Max AF'] = AF_allcases(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['AF Merged'], calcresults['cases_with_filters']['Min Max AF'] = AF_allcases(f'{cases_with_filters=}'.split('=')[0])

#Min/Max Amplification Factor Chart
AF_minmax_chart(f'{cases_wo_filters=}'.split('=')[0])
AF_minmax_chart(f'{cases_with_filters=}'.split('=')[0])

#Max total emissions
calcresults['cases_wo_filters']['Total Emissions Merged'], calcresults['cases_wo_filters']['Max Total Emission'] = maxtotalemission_allcases(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['Total Emissions Merged'], calcresults['cases_with_filters']['Max Total Emission'] = maxtotalemission_allcases(f'{cases_with_filters=}'.split('=')[0])

#Max Total Emission Chart
maxtotalemission_chart(f'{cases_wo_filters=}'.split('=')[0])
maxtotalemission_chart(f'{cases_with_filters=}'.split('=')[0])

#Max plant emissions
calcresults['cases_wo_filters']['Plant Emissions Merged'], calcresults['cases_wo_filters']['Max Plant Emission'] = maxplantemission_allcases(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['Plant Emissions Merged'], calcresults['cases_with_filters']['Max Plant Emission'] = maxplantemission_allcases(f'{cases_with_filters=}'.split('=')[0])

#Max Plant Emission Chart
maxplantemission_chart(f'{cases_wo_filters=}'.split('=')[0])
maxplantemission_chart(f'{cases_with_filters=}'.split('=')[0])

#Max HD Zext
calcresults['cases_wo_filters']['Max HD Zext'] = maxHD_impedances(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['Max HD Zext'] = maxHD_impedances(f'{cases_with_filters=}'.split('=')[0])

#Zext points for case group
calcresults['cases_wo_filters']['Zext points'], calcresults['cases_wo_filters']['Zext points_pu'], calcresults['cases_wo_filters']['Zext polygon'] = Zext_points(f'{cases_wo_filters=}'.split('=')[0])
calcresults['cases_with_filters']['Zext points'], calcresults['cases_with_filters']['Zext points_pu'], calcresults['cases_with_filters']['Zext polygon'] = Zext_points(f'{cases_with_filters=}'.split('=')[0])


# cases_dfs, summary_dict =read_result_sheet()
# calc_dfs['Plant Emissions'] = calc_plant_emission()
# calc_dfs['Total Emissions'] = total_emission()
# calc_dfs['Amplification Factor'] = amplification_factors()
# calcs_merged={}
# calcs_merged['AF'], calcs_merged['Min Max AF']=AF_allcases()
# calcs_merged['Total Emissions'], calcs_merged['Max Total Emission']=maxtotalemission_allcases()
# calcs_merged['Plant Emissions'], calcs_merged['Max Plant Emission']=maxplantemission_allcases()

# planning_limit_chart()
# allocation_limit_chart()
# AF_chart()
# maxtotalemission_chart()
# maxplantemission_chart()
# AF_minmax_chart()
# maxHD_Z = maxHD_impedances()
end_time = datetime.now()

print('Script finished running. Total Duration: {}'.format(end_time - start_time))