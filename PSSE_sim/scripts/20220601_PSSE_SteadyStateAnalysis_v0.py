# -*- coding: utf-8 -*-
"""
Created on Tue Feb 15 09:08:16 2022

@author: ESCO

FUNCTIONALYTY:
The script will run a steady state analysis, using the study inputs outlined in the 

COMMENTS:
        The script reads inputs from common excel spreadsheet located in folder: test_scenario_definitions. 
            + Contingency scenarios are defined in tab: SteadyStateStudies
            + Buses and Branches to be monitored are definded in tab: MonitorBuses and MonitorBranches
        The script then runs through differnt contingency scenarios, does the load flow analysis. 
        After that it create the output folder in local drive under Document\Projects: C:\Users\Dao Vu\Documents\Projects\ and export the results in the format of excel spreadsheet
            + Voltage Profiles and Line Loading Profile: contain the voltage profiles at different buses and loading on different lines in each contingency scenario
            + Voltage Profiles - GenChg and Line Loading Profile - GenChg: Compare the results with differnt dispatch level of generators.

@NOTE: 
        The path to the script need to be entered manually:
            current_dir=r"C:\Users\Dao Vu\ESCO Pacific\ESCO - Projects\24. SUM\3. Grid\1. Power System Studies\1. Main Test Environment\20220215_SUM\PSSE_sim\scripts"
        This can be improved (if needed) by defining the default path: 
            current_dir = os.getcwd()


V0b: Added Lib folder in Scripts folder -> add path before calling the function: sys.path.append(libpath)
V0c: Separate the input model and output results for backing up only the model and scripts on onedrive folder. Result will be in local drive: C:\Users\Dao Vu\Documents\Projects\
V0d: Add contingency scenarios into information spreadsheet -> read input from information spreadsheet when analysing contingency scenarios
V0e: tested with Summerville nearby network

"""
###############################################################################
#USER INPUTS
###############################################################################
"""
In this section the network cases to be used in the Steady State Study are defined. For the purpose
"""
TestDefinitionSheet = r'20230828_SUM_TESTINFO_V1.xlsx'

input_models={'HighLoad':{'on':'HighLoad\\SUMSF_high_genon.sav',
                           'off':'HighLoad\\SUMSF_high_genoff.sav'},
               'LowLoad': {'on':'LowLoad\\SUMSF_low_genon.sav',
                            'off':'LowLoad\\SUMSF_low_genoff.sav'},
                           }


#input_models={'HighLoad':{'on':'HighLoad\\HORSF_fault_high_genon.sav',
#                           'off':'HighLoad\\HORSF_high_genoff.sav'},
               #'LowLoad': {'on':'LowLoad\\HORSF_low_genon.sav',
                #            'off':'LowLoad\\HORSF_low_genoff.sav'},
#                           }
               
simulation_batch_label='SS'

while True:
    try:
        ext_gen_ctl = input("Is external generator control active? Respond in Y/N: ")
    except NameError:
        print("String input required")
        continue
    if ext_gen_ctl == 'Y' or ext_gen_ctl == 'N':
        break
    else:
        print("Enter 'Y' or 'N' only!")
        continue
    

###############################################################################
# Define location for PSSE path
###############################################################################
# import sys
import os, sys
import getpass
import pandas as pd
import numpy as np
#from itertools import izip, zip_longest
#import gen_ctrl_set as gcs # This is required as per the project. Depends upon the control system of the generator
from datetime import datetime
import time
from contextlib import contextmanager
timestr=time.strftime("%Y%m%d-%H%M%S")
start_time = datetime.now()
@contextmanager
def silence(file_object=None):
    """
    Discard stdout (i.e. write to null device) or
    optionally write to given file-like object.
    """
    if file_object is None:
        file_object = open(os.devnull, 'w')

    old_stdout = sys.stdout
    try:
        sys.stdout = file_object
        yield
    finally:
        sys.stdout = old_stdout


sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
sys.path.append(sys_path_PSSE)
os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE

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

def createShortcut(target, path):
    # target = ModelCopyDir # directory to which the shortcut is created
    # path = main_folder + "\\model_copies.lnk"  #This is where the shortcut will be created
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.save()

class Logger():
    stdout = sys.stdout
    messages = []
    def start(self): 
        sys.stdout = self
    def stop(self): 
        sys.stdout = self.stdout
    def write(self, text): 
        self.messages.append(text)
    def reset(self):
        self.messages=[]

def getFaultLvl(psseLog):
    line_counter=len(psseLog)-1
    while line_counter>0 and not ('X------------ BUS ------------X'  in psseLog[line_counter]):
        line_counter-=1
    if(line_counter>0):
        return float(psseLog[line_counter+1][38:47])
    else:
        return -1

def fault_lvls(scenario):
    fault_lvl = {'bus_no':[],'bus_name':[],'fault_1_genon':[],'fault_0_genon':[],'fault_1_genoff':[],'fault_0_genoff':[]}
    fault_type = [1,0] # This will depend upon the project
    gens_stat = ['genon','genoff']
    for gen_stat in gens_stat:
        if 'genon' in gen_stat:
            for fault in fault_type:
                log = Logger()
                for bus_no, bus_name in zip(bus_numbers,bus_names):
                    psspy.short_circuit_units(1)
                    psspy.short_circuit_z_units(1)
                    #log.start()
                    #psspy.bsys(1,0,[0.0,0.0],0,[],1,bus_no,0,[],0,[])
                    #psspy.iecs_4(1,0,[1,0,0,0,1,0,0,0,0,0,1,1,0,0,3,1,0],[ 0.1, 1.0],"","","")
                    #log.stop()
                    #key = bus_no
                    if fault == 1:
                        log.start()
                        psspy.bsys(1,0,[0.0,0.0],0,[],1,bus_no,0,[],0,[])
                        psspy.iecs_4(1,0,[1,0,0,0,1,0,0,0,0,0,1,1,0,0,3,1,0],[ 0.1, 1.0],"","","")
                        log.stop()
                    elif fault == 0:
                        log.start()
                        psspy.bsys(1,0,[0.0,0.0],0,[],1,bus_no,0,[],0,[])
                        psspy.iecs_4(1,0,[0,1,0,0,1,0,0,0,0,0,1,1,0,0,3,1,0],[ 0.1, 1.0],"","","")
                        log.stop()
                    if(bus_no not in fault_lvl['bus_no']) and (bus_name not in fault_lvl['bus_name']):
                        fault_lvl['bus_no'].append(bus_no)
                        fault_lvl['bus_name'].append(bus_name)
                    fault_lvl['fault_'+str(fault)+'_'+str(gen_stat)].append(getFaultLvl(log.messages))
        elif 'genoff' in gen_stat:
            for fault in fault_type:
                for i in range(0,len(scenario)):
                    if(scenario[i]['Element']=='Bus'):
                        psspy.dscn(scenario[i]['i_bus'])
                log = Logger()
                for bus_no, bus_name in zip(bus_numbers,bus_names):
                    psspy.short_circuit_units(1)
                    psspy.short_circuit_z_units(1)
                    #key = bus_no
                    log.start()
                    #psspy.bsys(1,0,[0.0,0.0],0,[],1,bus_no,0,[],0,[])
                    #psspy.iecs_4(1,0,[1,0,0,0,1,0,0,0,0,0,1,1,0,0,3,1,0],[ 0.1, 1.0],"","","")
                    if fault == 1:
                        psspy.bsys(1,0,[0.0,0.0],0,[],1,bus_no,0,[],0,[])
                        psspy.iecs_4(1,0,[1,0,0,0,1,0,0,0,0,0,1,1,0,0,3,1,0],[ 0.1, 1.0],"","","")
                    elif fault == 0:
                        psspy.bsys(1,0,[0.0,0.0],0,[],1,bus_no,0,[],0,[])
                        psspy.iecs_4(1,0,[0,1,0,0,1,0,0,0,0,0,1,1,0,0,3,1,0],[ 0.1, 1.0],"","","")
                    #psspy.ascc_3(1,0,[1,0,0,0,0,1,0,0,0,0,fault,0,0,0,0,0,0], 1.0,"","","")
                    log.stop()
                    if(bus_no not in fault_lvl['bus_no']) and (bus_name not in fault_lvl['bus_name']):
                        fault_lvl['bus_no'].append(bus_no)
                        fault_lvl['bus_name'].append(bus_name)
                    fault_lvl['fault_'+str(fault)+'_'+str(gen_stat)].append(getFaultLvl(log.messages))
    
    return fault_lvl


def monitoring():

    bus_df = pd.DataFrame()
    brch_df = pd.DataFrame()
    #branch results
    brch_info = []
    for i,j,k in zip(from_buses,to_buses,brch_ids):
        info = af.get_branch_info(i,j,k)
        brch_info.append(info)
    brch_df1 = brch_df.append(pd.DataFrame(data=brch_info))
    brch_df1 = brch_df1[0].apply(pd.Series)
    brch_df = brch_df.append(brch_df1)
    brch_df['brch_name']= brch_names
    #bus results
    bus_df['bus_name'] = []
    for bus_no,bus_name in zip(bus_numbers,bus_names):
        temp_bus_info = af.get_bus_info(bus_no,['TYPE','PU','ANGLED'])
        temp_bus_info[bus_no].update({'bus_name':bus_name})
        bus_df1 = pd.DataFrame.from_dict(temp_bus_info,orient = 'index')
        bus_df = bus_df.append(bus_df1)
         

    return brch_df, bus_df

def applyContingencies(scenario):
    for i in range(0, len(scenario)):
        if(scenario[i]['Element']=='Line'): # If fault elememt is a line
            psspy.branch_chng_3(scenario[i]['i_bus'],scenario[i]['j_bus'],str(scenario[i]['id']),[scenario[i]['status'],_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
        elif(scenario[i]['Element']=='Shunt'): # If fault elememt is a Shunt
            psspy.shunt_chng(scenario[i]['i_bus'],str(scenario[i]['id']),scenario[i]['status'],[_f,_f])
        elif(scenario[i]['Element']=='Machine'): # If fault elememt is a Machine
            psspy.machine_chng_2(scenario[i]['i_bus'],str(scenario[i]['id']),[scenario[i]['status'],1,0,0,0,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
        elif(scenario[i]['Element']=='Tx_2w'): # If fault elememt is a two windig transformer
            psspy.two_winding_chng_6(scenario[i]['i_bus'],scenario[i]['j_bus'],str(scenario[i]['id']),[scenario[i]['status'],_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
        elif(scenario[i]['Element']=='Bus'):
            psspy.dscn(scenario[i]['i_bus'])

def apply_gen_change(scenario, generation_value = 'Final'):
    if generation_value == 'Initial':
        for i in range(0, len(scenario)):
            if(scenario[i]['Element']=='Shunt'): # If fault elememt is a Shunt
                psspy.shunt_chng(scenario[i]['i_bus'],str(scenario[i]['id']),_i,[_f,_f]) #shunt can be switched off as part of gen change scenario.
            if(scenario[i]['Element']=='Machine'): # If fault elememt is a Machine
                psspy.machine_chng_2(scenario[i]['i_bus'],str(scenario[i]['id']),[_i,1,0,0,0,0],[scenario[i][generation_value] if scenario[i][generation_value] != '' else _f ,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #if status is 0, trhen machine will be automatically switched off.
    elif generation_value == 'Final':
        if scenario[0]['TestType']=='GenChange':
            for i in range(0, len(scenario)):
                if(scenario[i]['Element']=='Shunt'):# If fault elememt is a Shunt
                    psspy.shunt_chng(scenario[i]['i_bus'],str(scenario[i]['id']),scenario[i]['status'],[_f,_f]) #shunt can be switched off as part of gen change scenario.
                if(scenario[i]['Element']=='Machine'): # If fault elememt is a Machine
                    psspy.machine_chng_2(scenario[i]['i_bus'],str(scenario[i]['id']),[scenario[i]['status'],1,0,0,0,0],[scenario[i][generation_value] if scenario[i][generation_value] != '' else _f ,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
        elif scenario[0]['TestType']=='GenTrip':
            for i in range(0, len(scenario)):
                if(scenario[i]['Element']=='Shunt'): # If fault elememt is a Shunt
                    psspy.shunt_chng(scenario[i]['i_bus'],str(scenario[i]['id']),int(scenario[i]['Final']),[_f,_f]) #shunt can be switched off as part of gen change scenario.
                if(scenario[i]['Element']=='Machine'): # If fault elememt is a Machine
#                    psspy.machine_chng_2(scenario[i]['i_bus'],str(scenario[i]['id']),[scenario[i]['Final'],1,0,0,0,0],[_f ,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.machine_chng_2(scenario[i]['i_bus'],str(scenario[i]['id']),[scenario[i]['status'],1,0,0,0,0],[scenario[i][generation_value] if scenario[i][generation_value] != '' else _f ,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    
                
                
###############################################################################
# Define Project Paths
###############################################################################

overwrite = False # 
max_processes = 8 #set to the number of cores on my machine. Needs to be >= scenarioPerGroup --> increase for PSCAD machine
# testRun = '20211223_testing' #define a test name for the batch or configuration that is being tested
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
# main_folder_out = r"C:\1. Power System Studies\20220318_LSF\PSSE_sim\scripts" # Option to define the absolute path of the result location
ModelCopyDir = main_folder_out+"\\model_copies" #location of the model copies used to run the simulations
OutputDir0= main_folder_out+"\\result_data" #location of the simulation results
make_dir(OutputDir0)
make_dir(ModelCopyDir)

# Create shortcut linking to result folder if it is not stored in the main folder path
if main_folder_out != main_folder:
    createShortcut(ModelCopyDir, main_folder + "\\model_copies.lnk")
    createShortcut(OutputDir0, main_folder + "\\result_data.lnk")

# Locating the existing folders
testDefinitionDir= os.path.abspath(os.path.join(main_folder, os.pardir))+"\\test_scenario_definitions"
base_model = main_folder+"\\base_model" #parent directory of the workspace folder
base_model_workspace = main_folder+"\\base_model\\SMIB" #path of the workspace folder, formerly "workspace_folder" --> in case the workspace is located in a subdirectory of the model folder (as is the case with MUL model for example)
zingen=main_folder+"\\zingen\\dsusr_zingen.dll"
libpath = main_folder_out = os.path.abspath(main_folder) + "\\scripts\\Libs"
sys.path.append(libpath)
# print ("libpath = " + libpath)

# Directory to store Steady State/Dynamic result
OutputDir = OutputDir0+"\\steady_state"
make_dir(OutputDir)
outputResultPath=OutputDir+"\\"+testRun
make_dir(outputResultPath)


###############################################################################
#Text file to log errors
###############################################################################
error_file = open(outputResultPath + '\\' + 'error_log.txt','a')

###############################################################################
# Import additional functions
###############################################################################
import auxiliary_functions as af
import readtestinfo as readtestinfo
import psspy
import redirect
import Gen_Ctrl_v0_01 as genctrl
from Gen_Ctrl_v0_01 import gens_with_pf,gens_with_vdc,gens_with_pf_vc,gens

###############################################################################
# Define Inputs - Project Specific
###############################################################################

inputInforFile = testDefinitionDir+"\\"+TestDefinitionSheet
inputModel_workspace = main_folder+"\\base_model"

#inputModels = ['SUMSF_high_genoff.sav','SUMSF_high_genon.sav']
#inputModel = "20210720-180037-SubTrans-SystemNormal.sav"


###############################################################################
# Read Monitoring Infor
###############################################################################
# Read in buses to be monitored
monitored_buses = pd.read_excel(inputInforFile, sheet_name = 'MonitorBuses')
#print(monitored_buses)
bus_numbers = monitored_buses['bus_number'].to_list()
bus_numbers = [int(x) for x in bus_numbers]
bus_names = monitored_buses['bus_name'].to_list()
print(bus_numbers)

# Read in branches to be monitored
monitored_lines = pd.read_excel(inputInforFile,sheet_name = 'MonitorBranches')
print(monitored_lines)
from_buses = monitored_lines['brch_from'].to_list()
from_buses = [int(x) for x in from_buses]
print(from_buses)
to_buses = monitored_lines['brch_to'].to_list()
to_buses = [int(x) for x in to_buses]
brch_ids = monitored_lines['brch_id'].to_list()
brch_ids = [int(x) for x in brch_ids]
brch_names = monitored_lines['brch_name'].to_list()
print('brch_ids=',brch_ids[0])

###############################################################################
# Contingencies Infor
###############################################################################
# SteadyStateDict =  readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet)
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['SteadyStateStudies', 'ModelDetailsPSSE'])
SteadyStateDict = return_dict['SteadyStateStudies']
ModelDetailsPSSE = return_dict['ModelDetailsPSSE']

scenarios=SteadyStateDict.keys()
scenarios.sort(key = lambda x: x[3:] )
active_scenarios=[]
active_scenarios_Des = []
for scenario in scenarios:
    if(SteadyStateDict[scenario][0]['Run?']==1):
        active_scenarios.append(scenario)


###############################################################################
# Initialise PSSE and Start Studies
###############################################################################
# Start PSSE
redirect.psse2py()
psspy.psseinit(10000)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()
redirect.psse2py()

def run_scenarios(sav_file):
    log=''
    ####Need to write this function#####Very Important
    #generatorStatus=check_gen_status(gen_bus) #function that checks generator status at given bus and returns 1 or 0.
    
    #Create dataframe to store voltage levels
    vltg_lvls_df = pd.DataFrame()
    
    # Create DataFrame to store Contingency results
    bus_df_final = pd.DataFrame()
    brch_df_final = pd.DataFrame()
    
    # Create DataFrame to store Generation Changes results
    bus_df_Pchange = pd.DataFrame()
    brch_df_Pchange = pd.DataFrame()
    
    #create empty dataframe for fault levels
    fault_df = pd.DataFrame()
    
    # Original Results - Without any contingencies
    # Initialise case study
    with silence():
        psspy.case(inputModel_workspace +"\\"+ sav_file)
        
        if ext_gen_ctl == 'Y':
            genctrl.init_gens_vdc(gens_with_vdc)
        else:
            if(af.test_convergence(method="fnsl", taps="step")>1):
                error_file.write('\n' + str(sav_file)+ 'System did not converge')
                print("      Model loaded ok.") #HERE THROW ERROR MESSAGE IN CASE IT DOES NOT CONVERGE
                log+="      Model loaded ok."
    
    
        #save intialise model
        base = os.path.basename(inputModel_workspace +"\\"+ sav_file)
        make_dir(main_folder + "\\" + 'model_copies', dir_name='Steady State')
        make_dir(main_folder + "\\" + 'model_copies\Steady State', dir_name='Normal State')
        psspy.save(main_folder + "\\" + 'model_copies\\Steady State\\Normal State' + "\\" + base)
        
    
    with silence():    
        # Monitor variables and add to final DataFrame
        brch_df0, bus_df0 = monitoring()
        vltg_lvls_df = vltg_lvls_df.append(bus_df0)
        
        bus_df0['CaseNr']= "case0 (base)"
        bus_df0['VolDev (%)']= 0
        
        brch_df0['CaseNr']= "case0 (base)"
        brch_df0['Loading (%)']= np.round((brch_df0['MVA'] / brch_df0['RATING1'])*100,2)
        bus_df_final = bus_df_final.append(bus_df0)
        brch_df_final = brch_df_final.append(brch_df0)
        
        
        
        
    
    # Contingency cases 
    for case_num in active_scenarios:
        scenario=SteadyStateDict[case_num]
        #log+=active_scenarios[case_num]
        log+=case_num
        
        #Fault Level Results. This require only one model which is either given by the NSP or use the high load scenario with genon settings
        if(scenario[0]['TestType']=='Fault' and scenario[0]['CaseRef'] in sav_file):
            fault_results = fault_lvls(scenario)
            fault_df = fault_df.append(pd.DataFrame.from_dict(fault_results))
        else:
            pass
        
        if(scenario[0]['TestType']=='GenChange' or scenario[0]['TestType']=='GenTrip'): # If Generation change included in the study list: Run the comparison
    
            # Create DataFrame to stor final results
            bus_df_P1 = pd.DataFrame()
            brch_df_P1 = pd.DataFrame()
            # Initialise case study
            psspy.case(inputModel_workspace +"\\"+ sav_file)
            
            # Run loadflow taps = enabled
            if ext_gen_ctl == 'Y':
                genctrl.init_gens_vdc(gens_with_vdc)
            else:
                if(af.test_convergence(method="fnsl", taps="step")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) + ' ' +  str(scenario[0]['CaseRef'])+' System did not converge')# THROW ERROR IF MISMATCH >1
            
            
            apply_gen_change(scenario, generation_value = 'Initial')
            # Run loadflow taps = enabled with the updated intial conditions
            if ext_gen_ctl == 'Y':
                genctrl.init_gens_vdc(gens_with_vdc)
            else:
                if(af.test_convergence(method="fnsl", taps="step")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) + ' ' +  str(scenario)+' System did not converge')# THROW ERROR IF MISMATCH >1
            
            #save the case with intial conditions with gen change
            base = os.path.basename(inputModel_workspace +"\\"+ sav_file)
            make_dir(main_folder + "\\" + 'model_copies\Steady State', dir_name=case_num)
            psspy.save(main_folder + "\\" + 'model_copies\\Steady State\\' + case_num + "\\" +  'Before_' +base)
            
            # Monitor variables and add to final DataFrame
            brch_df_P1, bus_df_P1 = monitoring()
            bus_df_P1['CaseNr']= case_num
            bus_df_P1['CaseRef']= scenario[0]['CaseRef']
            # bus_df_P1['CaseDes']= case_num
            # bus_df_P1['CaseNr']= "Initial Condition"
            # bus_df_P1['VolDev']= 0
            # brch_df_P1['CaseNr']= "Initial Condition"
#            RATING = max(brch_df_P1['RATING1'],brch_df_P1['RATING2']) #contingency rating
            RATING = brch_df_P1['RATING2']
            brch_df_P1['Loading_ini (%)']= np.round((brch_df_P1['MVA'] / RATING)*100,2)
            brch_df_P1['CaseNr']= case_num
            brch_df_P1['CaseRef']= scenario[0]['CaseRef']
            bus_df_Pupdate = bus_df_P1
            brch_df_Pupdate = brch_df_P1
    
    
            # Reduce Generation as required and Apply all the events listed in the given contingency scenario
            apply_gen_change(scenario, generation_value = 'Final')
            
            
            # Solve case with locked tap
            if ext_gen_ctl == 'Y':
                genctrl.lckd_gens_vdc(gens_with_vdc)
            else:
                if(af.test_convergence(method="fnsl", taps="locked")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) + ' ' +  str(scenario[0]['CaseRef'])+' System did not converge')# THROW ERROR IF MISMATCH >1
            
            # save the case with the after conditions and taps locked
            base = os.path.basename(inputModel_workspace +"\\"+ sav_file)
            make_dir(main_folder + "\\" + 'model_copies\Steady State', dir_name=case_num)
            psspy.save(main_folder + "\\" + 'model_copies\\Steady State\\' + case_num + "\\" +  'After_' +base)
            
            
            
            # Monitoring
            brch_df_P2, bus_df_P2 = monitoring()
            bus_df_Pupdate['VolDev (%)'] = np.round((bus_df_P2['PU'] - bus_df_P1['PU'])*100,4)
            bus_df_Pupdate['PU_final'] = bus_df_P2['PU'] 
            bus_df_Pupdate['ANGLED_final'] = bus_df_P2['ANGLED'] 
#            RATING = max(brch_df_P2['RATING1'],brch_df_P2['RATING2']) #contingency rating
            RATING = brch_df_P2['RATING2']
            brch_df_Pupdate['Loading_final (%)'] = np.round((brch_df_P2['MVA'] / RATING)*100,2)
            brch_df_Pupdate['MVA_final'] = brch_df_P2['MVA']
            brch_df_Pupdate['P_final'] = brch_df_P2['P']
            brch_df_Pupdate['Q_final'] = brch_df_P2['Q']
            brch_df_Pupdate['PCTMVARATE_final'] = brch_df_P2['PCTMVARATE']
    
            bus_df_Pchange = bus_df_Pchange.append(bus_df_Pupdate)
            brch_df_Pchange = brch_df_Pchange.append(brch_df_Pupdate)
            
    
        else:
    
            # Initialise case study
            psspy.case(inputModel_workspace +"\\"+ sav_file)
            
            if ext_gen_ctl == 'Y':
                genctrl.init_gens_vdc(gens_with_vdc)
            else:
                if(af.test_convergence(method="fnsl", taps="step")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) + ' ' +  str(scenario[0]['CaseRef'])+' System did not converge')# THROW ERROR IF MISMATCH >1
            
            #save the case before the contingency
            base = os.path.basename(inputModel_workspace +"\\"+ sav_file)
            make_dir(main_folder + "\\" + 'model_copies\Steady State', dir_name=case_num)
            psspy.save(main_folder + "\\" + 'model_copies\\Steady State\\' + case_num + "\\" +  'Before_' +base)
            
            
            # Apply all the events listed in the given contingency scenario
            applyContingencies(scenario)
            
            # Solve case with locked tap
            if ext_gen_ctl == 'Y':
                genctrl.lckd_gens_vdc(gens_with_vdc)
            else:
                if(af.test_convergence(method="fnsl", taps="locked")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) + ' ' +  str(scenario[0]['CaseRef'])+' System did not converge')# THROW ERROR IF MISMATCH >1
            
            #save the case after the contingency with locked taps
            base = os.path.basename(inputModel_workspace +"\\"+ sav_file)
            make_dir(main_folder + "\\" + 'model_copies\Steady State', dir_name=case_num)
            psspy.save(main_folder + "\\" + 'model_copies\\Steady State\\' + case_num + "\\" +  'After_' +base)
            
            # Monitoring
            brch_df, bus_df = monitoring()
            bus_df['CaseNr']= case_num
            bus_df['CaseRef']= scenario[0]['CaseRef']
            bus_df['VolDev (%)']= np.round((bus_df['PU'] - bus_df0['PU'])*100,4) #calculates deviaton compared to voltage level before contingency
            brch_df['CaseNr']= case_num
            brch_df['CaseRef']= scenario[0]['CaseRef']
#            RATING = np.max(brch_df['RATING1'],brch_df['RATING2']) #contingency rating
            RATING = brch_df['RATING2']
#            if brch_df['RATING2'] !='' and brch_df['RATING2'] > brch_df['RATING1']: RATING = brch_df['RATING2'] #contingency rating
            brch_df['Loading (%)']= np.round((brch_df['MVA'] / RATING)*100,2)
    
            bus_df_final = bus_df_final.append(bus_df)
            brch_df_final = brch_df_final.append(brch_df)
    bus_df_final = bus_df_final[bus_df_final.CaseNr != 'case0 (base)']       
    results={'voltage_levels':vltg_lvls_df, 'line_loadings':brch_df_final, 'volt_fluc_gen_chng':bus_df_Pchange, 'volt_fluc_lol':bus_df_final, 'fault_lvls': fault_df }
    return results, log

###############################################################################
# Exporting Results
###############################################################################
out_file_name = 'Steady State Analysis Results.xlsx'
writer = pd.ExcelWriter(outputResultPath+'\\'+out_file_name,engine = 'xlsxwriter')

log=''
for snapshot_name in input_models.keys(): 
    log+='\nrunning '+snapshot_name
    for config in ['on', 'off']:
        log+="\n   Generator "+config
        print("   running scenarios..")
        results, sim_log=run_scenarios(input_models[snapshot_name][config])
        print(sim_log)

        results['voltage_levels'].to_excel(writer,sheet_name = 'Voltage Level_'+snapshot_name+'_'+config, index=True )
        results['line_loadings'].to_excel(writer,sheet_name = 'Line Loadings_'+snapshot_name+'_'+config, index=True )
        results['volt_fluc_gen_chng'].to_excel(writer,sheet_name = 'Volt Fluc GenChg_'+ snapshot_name+'_'+config, index=True )
        results['volt_fluc_lol'].to_excel(writer,sheet_name = 'Volt Fluc Lol_'+ snapshot_name+'_'+config, index=True )
        results['fault_lvls'].to_excel(writer,sheet_name = 'Fault Level_'+ snapshot_name+'_'+config, index=True )

writer.save()
error_file_size = os.path.getsize(outputResultPath + '\\' + 'error_log.txt')
if(error_file_size ==0):error_file.write('The system converged in all models and all scenarios')
error_file.close()


#Do work above
end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))
