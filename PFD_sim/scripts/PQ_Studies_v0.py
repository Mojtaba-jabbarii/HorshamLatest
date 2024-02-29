# -*- coding: utf-8 -*-
"""
Created on Wed Nov 15 10:31:30 2023

@author: OX2 - Michael Magpantay 

Functionality:
    
    1. Setup each scenarios in PowerFactory
    2. Run all scenarios for (In sequence)
        a. Impedance Loci DPL script in PowerFactory
        b. Harmonic Load Flow
    3. Retrieve Harmonic Load Flow results (*.csv format)
    4. Perform frequency sweep pre- and post-connection
    4. Consolidate all results into one excel file (optional)
"""

###############################################################################
#USER INPUTS
###############################################################################
TestDefinitionSheet = r'20231117_SUM_TESTINFO_V0.xlsx'
sys_path_PF = r'C:\Program Files\DIgSILENT\PowerFactory 2023\Python\3.9' #Change in case path is different
simulation_batch_label='PQ'

###############################################################################
# Define location for PF path
###############################################################################
from datetime import datetime
import time
timestr=time.strftime("%Y%m%d-%H%M%S")
start_time = datetime.now()
import os, sys
sys.path.append(sys_path_PF) #Change path as needed
import pandas as pd


###############################################################################
# Supporting functions
###############################################################################
def make_dir(dir_path, dir_name=""):
    dir_make = os.path.join(dir_path, dir_name)
    try:
        os.mkdir(dir_make)
    except OSError:
        pass

def createPath(main_folder_out):
    path = os.path.normpath(main_folder_out)
    path_splits = path.split(os.sep) # Get the components of the path
    child_folder = r"C:" # Build up the output path from C: directory
    for i in range(len(path_splits)-1):
        child_folder = child_folder + "\\" + path_splits[i+1]
        make_dir(child_folder)
    return child_folder


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

# Create directory for storing the results
if "ESCO Pacific\ESCO - Projects" in main_folder: # if the current folder is online (under ESCO - Projects), create a new directory to store the result
    main_path_out = main_folder.replace("ESCO Pacific\ESCO - Projects","Documents\Projects") # Change the path from Onedrive to Local in Documents
    path = os.path.normpath(main_path_out)
    path_splits = path.split(os.sep) # Get the components of the path
    child_folder = r"C:" # Build up the output path from C: directory
    for i in range(len(path_splits)-1):
        child_folder = child_folder + "\\" + path_splits[i+1]
        make_dir(child_folder)
    main_folder_out = child_folder
else: # if the main folder is not in Onedrive, then store the results in the same location with the model
    main_folder_out = main_folder
OutputDir0= main_folder_out+"\\result_data" #location of the simulation results
make_dir(OutputDir0)

# Locating the existing folders
testDefinitionDir= os.path.abspath(os.path.join(main_folder, os.pardir))+"\\test_scenario_definitions"
libpath = main_folder_out = os.path.abspath(main_folder) + "\\scripts\\Libs"
sys.path.append(libpath)

# Directory to store PQ result
OutputDir = OutputDir0+"\\PQ"
make_dir(OutputDir)
outputResultPath=OutputDir+"\\"+testRun
make_dir(outputResultPath)


###############################################################################
# Define Inputs - Project Specific
###############################################################################
inputInforFile = testDefinitionDir+"\\"+TestDefinitionSheet

###############################################################################
# Model Information
###############################################################################
# SteadyStateDict =  readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet)
import readtestinfo as readtestinfo
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ModelDetailsPF', 'PowerQuality'])
PQDict = return_dict['PowerQuality']
ModelDetailsPF = return_dict['ModelDetailsPF']

scenarios=PQDict.keys()
#scenarios.sort(key = lambda x: x[3:] )
active_scenarios=[]
for scenario in scenarios:
    if(PQDict[scenario][0]['Run?']==1):
        active_scenarios.append(scenario)

###############################################################################
# Open PF and Run Studies
###############################################################################
if __name__ == "__main__":
    import powerfactory as pf
app = pf.GetApplication()
print("Opening PowerFactory...")
if app is None:
    raise Exception('getting Powerfactory application failed')

#Define project name and study case    
app.Show() #Show PowerFactory on screen (non-interactive mode)
proj_name =   ModelDetailsPF['Project Name']
study_case = ModelDetailsPF['Study Case'] 

#Activate project
project = app.ActivateProject(proj_name)
#proj = app.GetActiveProject()

#Initialise results dictionary
results_dict={'LF':{'PoC_MW':{},'PoC_MVar':{}},'InvDispatch':{},'THD':{},'InvDispatch_summary':{'No. of active BESS Inv':{},'No. of active PV Inv':{},'BESS kW':{},'BESS kVar':{},'PV kW':{},'PV kVar':{}}}

#Initialise DPLs
dpl_BESS_discharge_char =app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_BESS Discharge Characteristic'] + ".ComDpl")[0]
dpl_BESS_charge_char =app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_BESS Charge Characteristic'] + ".ComDpl")[0]
dpl_plus10_BoP = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_+10%BoP'] + ".ComDpl")[0]
dpl_minus10_BoP = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_-10%BoP'] + ".ComDpl")[0]
dpl_resetvariations = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_Reset Variations'] + ".ComDpl")[0]
dpl_Z_Loci = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_Z Loci']+".ComDpl")[0]
dpl_SwOffGrid = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_Switch off Grid'] + ".ComDpl")[0]
dpl_SwOnGrid = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_Switch on Grid'] + ".ComDpl")[0]
dpl_SwOffPoCFdr = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_Switch off PoC Feeder'] + ".ComDpl")[0]
dpl_SwOnPoCFdr = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_Switch on PoC Feeder'] + ".ComDpl")[0]
dpl_EmissionOff = app.GetProjectFolder("script").GetContents(ModelDetailsPF['DPL_Emissions off'] + ".ComDpl")[0]

log=''
for case_num in active_scenarios:
    scenario=PQDict[case_num]
    log+=case_num #log cases that are processed
    
    #Activate scenario,  and variation in PowerFactory
    PFactivescen = app.GetProjectFolder("scen").GetContents(scenario[0]['Scenario Name']+".IntScenario")[0]
    PFactivescen.Activate()
   
    #Activate inverter characteristic
    if PQDict[case_num][0]['Inverter Characteristic']=='BESS Discharging':
        dpl_BESS_discharge_char.Execute()
    if PQDict[case_num][0]['Inverter Characteristic']=='BESS Charging':
        dpl_BESS_charge_char.Execute()
   
    #Activate variation
    if PQDict[case_num][0]['Variation Name']=='+10%BoP':
        dpl_plus10_BoP.Execute()
    if PQDict[case_num][0]['Variation Name']=='-10%BoP':
        dpl_minus10_BoP.Execute()
    if PQDict[case_num][0]['Variation Name']=='':
        dpl_resetvariations.Execute()
                
    #Run Load Flow and retrieve results (POC MW & MVar)
    oLF = app.GetFromStudyCase('ComLdf')
    oLF.Execute()
    POC_LF=app.GetCalcRelevantObjects(ModelDetailsPF['PoC ID'] + '.ElmTerm')
    for LF in POC_LF:
        PoC_MW = -1*getattr(LF,'m:Pnet')
        PoC_MW = float("{:.2f}".format(PoC_MW))
        PoC_MVar = -1*getattr(LF,'m:Qnet')
        PoC_MVar = float("{:.2f}".format(PoC_MVar))
        results_dict['LF']['PoC_MW'][case_num]=PoC_MW
        results_dict['LF']['PoC_MVar'][case_num]=PoC_MVar
        
    #Retrieve dispatch of individual inverters
    results_dict['InvDispatch'][case_num] = {'kW':{},'kVar':{}}
    BESS_Invs=app.GetCalcRelevantObjects('*.ElmGenstat')
    for BESS_Inv in BESS_Invs:
        name = getattr(BESS_Inv, 'loc_name')
        BESS_P = 1000*getattr(BESS_Inv,'pgini')
        BESS_P = float("{:.0f}".format(BESS_P))
        BESS_Q = 1000*getattr(BESS_Inv,'qgini')
        BESS_Q = float("{:.0f}".format(BESS_Q))
        results_dict['InvDispatch'][case_num]['kW'][name]=BESS_P
        results_dict['InvDispatch'][case_num]['kVar'][name]=BESS_Q
        
    #Count number of active BESS Inverters
    count_BESS = 0
    for i in range(len(BESS_Invs)):
        status = BESS_Invs[i].IsConnected()
        if status==1:
            count_BESS += 1
    results_dict['InvDispatch_summary']['No. of active BESS Inv'][case_num] = count_BESS
    
    #Get Dispatch summary  
    for BESS_Inv in BESS_Invs:
        status = BESS_Inv.IsConnected()
        BESS_P_summary = 0
        if status == 1:
            BESS_P_summary = 1000*getattr(BESS_Inv,'pgini')
            BESS_P_summary = float("{:.0f}".format(BESS_P_summary))
            if BESS_P_summary !=0:
                BESS_P_summary = BESS_P_summary 
        results_dict['InvDispatch_summary']['BESS kW'][case_num] = BESS_P_summary
    
    for BESS_Inv in BESS_Invs:
        status = BESS_Inv.IsConnected()
        BESS_Q_summary = 0
        if status == 1:
            BESS_Q_summary = 1000*getattr(BESS_Inv,'qgini')
            BESS_Q_summary = float("{:.0f}".format(BESS_Q_summary))
            if BESS_Q_summary !=0:
                BESS_Q_summary = BESS_Q_summary 
        results_dict['InvDispatch_summary']['BESS kVar'][case_num] = BESS_Q_summary
            
    PV_Invs=app.GetCalcRelevantObjects('*.ElmPVsys')
    for PV_Inv in PV_Invs:
        name = getattr(PV_Inv, 'loc_name')
        PV_P = getattr(PV_Inv,'pgini')
        PV_P = float("{:.0f}".format(PV_P))
        PV_Q = getattr(PV_Inv,'qgini')
        PV_Q = float("{:.0f}".format(PV_Q))
        results_dict['InvDispatch'][case_num]['kW'][name]=PV_P
        results_dict['InvDispatch'][case_num]['kVar'][name]=PV_Q
        
    #Count number of active PV Inverters
    count_PV = 0
    for i in range(len(PV_Invs)):
        status = PV_Invs[i].IsConnected()
        if status==1:
            count_PV += 1
    results_dict['InvDispatch_summary']['No. of active PV Inv'][case_num] = count_PV
    
    #Get Dispatch summary  
    for PV_Inv in PV_Invs:
        status = PV_Inv.IsConnected()
        PV_P_summary = 0
        if status == 1:
            PV_P_summary = getattr(PV_Inv,'pgini')
            PV_P_summary = float("{:.0f}".format(PV_P_summary))
            if PV_P_summary !=0:
                PV_P_summary = PV_P_summary
        results_dict['InvDispatch_summary']['PV kW'][case_num] = PV_P_summary
    
    for PV_Inv in PV_Invs:
        status = PV_Inv.IsConnected()
        PV_Q_summary = 0
        if status == 1:
            PV_Q_summary = getattr(PV_Inv,'qgini')
            PV_Q_summary = float("{:.0f}".format(PV_Q_summary))
            if PV_Q_summary !=0:
                PV_Q_summary = PV_Q_summary
        results_dict['InvDispatch_summary']['PV kVar'][case_num] = PV_Q_summary
                
    #Run Impedance Loci Script for active case
    dpl_Z_Loci.Execute()
      
    #Run Harmonic Load Flow and retrieve THD results
    oHarm = app.GetFromStudyCase('ComHldf')
    oHarm.Execute()
    THDresults=app.GetCalcRelevantObjects(ModelDetailsPF['PoC ID'] + '.ElmTerm')
    for THDresult in THDresults:
        THD = getattr(THDresult,'m:THD')
        THD = float("{:.2f}".format(THD))
        results_dict['THD'][case_num]=THD
        
    #Retrieve Zequiv results and save as csv
    comRes = app.GetFromStudyCase("ComRes")
    comRes.iopt_exp = 6 # to export as csv
    comRes.f_name = outputResultPath + "\\" + case_num + "_Zequiv" + ".csv" # File Name
    comRes.iopt_sep = 1 # to use the system separator
    comRes.iopt_vars = 0 # to export values
    comRes.Execute()


    #Switch off PoC Feeder to get Zext
    dpl_SwOffPoCFdr.Execute()
    #Run Harmonic Load Flow
    oHarm = app.GetFromStudyCase('ComHldf')
    oHarm.Execute()
    #Retrieve Zbackground results and save as csv
    comRes = app.GetFromStudyCase("ComRes")
    comRes.iopt_exp = 6 # to export as csv
    comRes.f_name = outputResultPath + "\\" + case_num + "_Zext" + ".csv" # File Name
    comRes.iopt_sep = 1 # to use the system separator
    comRes.iopt_vars = 0 # to export values
    comRes.Execute()
    
    #Switch back PoC Feeder on   
    dpl_SwOnPoCFdr.Execute()

    #Switch off Grid Equivalent to get Zgen (Z of the entire plant with emissions set to zero)
    dpl_SwOffGrid.Execute()
    dpl_EmissionOff.Execute()
    #Run Harmonic Load Flow
    oHarm = app.GetFromStudyCase('ComHldf')
    oHarm.Execute()
    #Retrieve Znorton results and save as csv
    comRes = app.GetFromStudyCase("ComRes")
    comRes.iopt_exp = 6 # to export as csv
    comRes.f_name = outputResultPath + "\\" + case_num + "_Zgen" + ".csv" # File Name
    comRes.iopt_sep = 1 # to use the system separator
    comRes.iopt_vars = 0 # to export values
    comRes.Execute()
    
    #Switch back Grid Equivalent on  
    dpl_SwOnGrid.Execute()
    
    #Run Harmonic Load Flow
    oHarm = app.GetFromStudyCase('ComHldf')
    oHarm.Execute()
    
    #Reset variations
    dpl_resetvariations.Execute()
    
    print("Finished running these cases: " + log)
    end_time = datetime.now()
    print('Elapsed time after last case: {}'.format(end_time - start_time))

##############################################################################
#Merging all Results
##############################################################################
import glob    
all_files = glob.glob(os.path.join(outputResultPath, "*.csv"))
writer = pd.ExcelWriter(outputResultPath + "\\" + testRun + ".xlsx", engine='xlsxwriter')

for f in all_files:
    df = pd.read_csv(f)
    df.insert(0,"h",list(range(1,51)),True)
    df.to_excel(writer, sheet_name=os.path.splitext(os.path.basename(f))[0], index =False)



#Deleting original files
for f in all_files:
    path = os.path.join(outputResultPath, f)
    os.remove(path)

PoC_LF_THD = pd.DataFrame({"POC MW":results_dict['LF']['PoC_MW'],"POC MVar":results_dict['LF']['PoC_MVar'],"THD %":results_dict['THD'],
                           "No. of active BESS Inv":results_dict['InvDispatch_summary']['No. of active BESS Inv'],
                           "No. of active PV Inv":results_dict['InvDispatch_summary']['No. of active PV Inv'],
                           "BESS kW":results_dict['InvDispatch_summary']['BESS kW'],"BESS kVar":results_dict['InvDispatch_summary']['BESS kVar'],
                           "PV kW":results_dict['InvDispatch_summary']['PV kW'],"PV kVar":results_dict['InvDispatch_summary']['PV kVar']})
PoC_LF_THD.reset_index(inplace=True)
PoC_LF_THD.rename(columns={'index':'Case name'},inplace=True)
writer2 = pd.ExcelWriter(outputResultPath + "\\" + testRun + "_Inv Dispatch" + ".xlsx", engine='xlsxwriter')
PoC_LF_THD.to_excel(writer,sheet_name='POC LF_THD',index =False)
writer.save()

#Do work above
for name, df in results_dict['InvDispatch'].items():
    globals()[name]=df
    df = pd.DataFrame.from_dict(df)
    df.to_excel(writer2,sheet_name=name)
    

writer2.save()




# import glob    
# all_files = glob.glob(os.path.join(outputResultPath, "*.csv"))
# writer = pd.ExcelWriter(outputResultPath + "\\" + testRun + ".xlsx", engine='xlsxwriter')

# for f in all_files:
#     df = pd.read_csv(f)
#     df.insert(0,"h",list(range(1,51)),True)
#     df.to_excel(writer, sheet_name=os.path.splitext(os.path.basename(f))[0], index =False)

# writer.save()

# #Deleting original files
# for f in all_files:
#     path = os.path.join(outputResultPath, f)
#     os.remove(path)

# PoC_LF_THD = pd.DataFrame({"POC MW":results_dict['LF']['PoC_MW'],"POC MVar":results_dict['LF']['PoC_MVar'],"THD %":results_dict['THD'],
#                            "No. of active BESS Inv":results_dict['InvDispatch_summary']['No. of active BESS Inv'],
#                            "No. of active PV Inv":results_dict['InvDispatch_summary']['No. of active PV Inv'],
#                            "BESS kW":results_dict['InvDispatch_summary']['BESS kW'],"BESS kVar":results_dict['InvDispatch_summary']['BESS kVar'],
#                            "PV kW":results_dict['InvDispatch_summary']['PV kW'],"PV kVar":results_dict['InvDispatch_summary']['PV kVar']})
# PoC_LF_THD.reset_index(inplace=True)
# PoC_LF_THD.rename(columns={'index':'Case name'},inplace=True)
# writer2 = pd.ExcelWriter(outputResultPath + "\\" + testRun + "_LF_THD" + ".xlsx", engine='xlsxwriter')
# PoC_LF_THD.to_excel(writer2,sheet_name='POC LF_THD',index =False)


# #Do work above
# for name, df in results_dict['InvDispatch'].items():
#     globals()[name]=df
#     df = pd.DataFrame.from_dict(df)
#     df.to_excel(writer2,sheet_name=name)
    

# writer2.save()







end_time = datetime.now()
print('Finished running all cases. Total Duration: {}'.format(end_time - start_time))

