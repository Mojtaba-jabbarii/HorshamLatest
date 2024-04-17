# -*- coding: utf-8 -*-
"""
Created on Tue Oct 20 09:23:53 2020

CONCEPT: 
        Same as the PSCAD script, this script reads a list of (SMIB-type) tests from an excel spreadsheet
        It is provided with a model and folder path to a source folder structured in a particular way and the aforementioned spreadsheet.
        
        the script will then automatically execute the tests defined in the spreadsheet 
        
        For every scenario, a model copy is generated, in which the required changes for the specific test type are implemented. This includes required changes to the dyr file(s) and automated generation of test profiles for the ZinGen model if required.
        The test is then performed on that model copy in a separtate Python subprocess using the multiprocessing toolbox, to avoid access conflicts.
        
        Results are saved in a standarised way along with metadata for further processing
        Plots will be generated using a separate script based on the ESCO plot tools that can be fed with both PSS/E and PSCAD results

ATTENTION:        
        Plot channels (other than Voltage, Angle, Frequency, Active power and reactive power at the locations specified in the excel test spreadsheet)
        need to be added manually in the script 'run_simulation.py'
        
        The way setpoint changes are spplied in models may change between differnet models. Make sure the code in run_simulation for the setpoint change cases matches the model at hand. 

NOTES:
        + 30/3/2022: Update the path to store results and model copies
        + 31/3/2022: Update base_model to base_model_workspace for copying only model in the work space  
        + 06/12/2023: copy the plot folder to local drive so the plot activities can be done locally
        
@author: Mervin Kall
 """
import os, sys
from multiprocessing import*
import time
import datetime
from subprocess import*
import shutil
import math
import cmath
import getpass
# import readtestinfo
from win32com.client import Dispatch
import time
#timestr = time.strftime("%Y%m%d")
timestr = str(datetime.datetime.now().strftime("%Y%m%d-%H%M"))

#-----------------------------------------------------------------------------
# USER CONFIGURABLE PARAMETERS
#-----------------------------------------------------------------------------
TestDefinitionSheet = r'20230828_SUM_TESTINFO_V2.xlsx'
#simulation_batches=['DMAT', 'Prof_chng', 'AEMO_fdb', 'missing', 'legend', 'SCR_chng', 'timing'] #specify batch from spreadsheet that shall be run. If empty, run all batches
#simulation_batches=['DMATsl1','DMATsl2','DMATsl3','DMATsl4','DMATsl5','DMATsl6','DMATsl']
simulation_batches=['S5253','S5254','S52511','S52513','S52514','S5255Iq1','S5255Iq2','S5255Iq3']
#simulation_batches=['gps_dbg']
#simulation_batches=['S5254','S5255','S5257','S52511','S52513','S52514','S52515','S52516']
#The below can alternatively be defined in the Excel sheet

overwrite = False # 
max_processes = 8 #set to the number of cores on my machine. Needs to be >= scenarioPerGroup --> increase for PSCAD machine
# testRun = '20211223_testing' #define a test name for the batch or configuration that is being tested
# testRun = str(datetime.datetime.now().strftime("%Y%m%d-%H%M"))

try:
    testRun = timestr + '_' + simulation_batches[0] #define a test name for the batch or configuration that is being tested -> link to time stamp for auto update
except:
    testRun = timestr
    
#-----------------------------------------------------------------------------
# Auxiliary Functions
#-----------------------------------------------------------------------------
def initializer(l, sem):
    global semaphore
    semaphore =l
    global taskSem
    taskSem=sem
    return

def worker(scenario):# i is the scenario key
    taskSem.acquire()
    #workspace_folder_path_local = createModelCopy(i)
    #workspace_folder_name_local = TestConfigDict['test_names'][test_list[i]-1]
    workspace_folder = createModelCopy(scenario, testRun) # SCRIPT CHANGE: create one new folder for the simulation set. In that folder, 
    scenario_params = ScenariosDict[scenario]
    run_simulation(main_folder, scenario, scenario_params, workspace_folder, ProjectDetailsDict, PSSEmodelDict, SetpointsDict, ProfilesDict)
    taskSem.release()
    return

def createModelCopy(scenario, testRun):
    global base_model
    global ModelCopyDir
    pass
    if (testRun==''):
        existing_batches=next(os.walk('.'))[1]    
        #determine next index
        testRunNr=0
        for testRun in existing_batches: # if an existing batch has the same or higher ID:
            if('testSeries'in testRun):
                nr=int(testRun[testRun.index('_'):-1])
                if nr>=testRunNr:
                    testRunNr+=1
        testRun='testSeries_'+str(testRunNr)        
     
    # model_copy_dir = main_folder+"\\model_copies\\"+testRun+"\\"+scenario
    model_copy_dir = ModelCopyDir+"\\"+testRun+"\\"+scenario
    try:
        # shutil.copytree(base_model, model_copy_dir)
        shutil.copytree(base_model_workspace, model_copy_dir)
    except OSError:
       print("Creation of the directory %s failed" % model_copy_dir)
    else:
       print("Successfully created the directory %s" % model_copy_dir)

    #add Zingen dll to model directory
    subdir=''
    if(base_model != base_model_workspace): #workspace can be subfolder of model folder
        subdir="\\"+os.path.basename(os.path.normpath(base_model_workspace))
    try:
        # shutil.copyfile(zingen, model_copy_dir+subdir+'\\dsusr_zingen.dll')
        shutil.copyfile(zingen, model_copy_dir+'\\dsusr_zingen.dll')
    except:
        print("Copying Zingen lib file failed")
    else:
        print("copied Zingen lib file")
    
    return model_copy_dir, testRun

def make_dir(dir_path, dir_name=""):
    dir_make = os.path.join(dir_path, dir_name)
    try:
        os.mkdir(dir_make)
    except OSError:
        pass

def createShortcut(target, path):
    # target = ModelCopyDir # directory to which the shortcut is created
    # path = main_folder + "\\model_copies.lnk"  #This is where the shortcut will be created
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.save()

def createPath(main_folder_out):
    path = os.path.normpath(main_folder_out)
    path_splits = path.split(os.sep) # Get the components of the path
    child_folder = r""+ path_splits[0] # Build up the output path from base drive
    for i in range(len(path_splits)-1):
        child_folder = child_folder + "\\" + path_splits[i+1]
        make_dir(child_folder)
    return child_folder



#-----------------------------------------------------------------------------
# Define Project Paths
#-----------------------------------------------------------------------------
script_dir=os.getcwd()
main_folder=os.path.abspath(os.path.join(script_dir, os.pardir))

# Create directory for storing the results
if "OneDrive - OX2" in main_folder: # if the current folder is online (under OneDrive - OX2), create a new directory to store the result
    user = os.path.expanduser('~')
    main_path_out = main_folder.replace(user + "\OneDrive - OX2","C:\work") # Change the path from Onedrive to Local in c drive
    main_folder_out = createPath(main_path_out)
else: # if the main folder is not in Onedrive, then store the results in the same location with the model
    main_folder_out = main_folder
ModelCopyDir = main_folder_out+"\\model_copies" #location of the model copies used to run the simulations
OutputDir= main_folder_out+"\\result_data" #location of the simulation results
make_dir(OutputDir)
make_dir(ModelCopyDir)

# Create shortcut linking to result folder if it is not stored in the main folder path
if main_folder_out != main_folder:
    createShortcut(ModelCopyDir, main_folder + "\\model_copies.lnk")
    createShortcut(OutputDir, main_folder + "\\result_data.lnk")

# Locating the existing folders
testDefinitionDir= os.path.abspath(os.path.join(main_folder, os.pardir))+"\\test_scenario_definitions"
base_model = main_folder+"\\base_model" #parent directory of the workspace folder
base_model_workspace = main_folder+"\\base_model\\SMIB" #path of the workspace folder, formerly "workspace_folder" --> in case the workspace is located in a subdirectory of the model folder (as is the case with MUL model for example)
zingen=main_folder+"\\zingen\\dsusr_zingen.dll"
libpath = os.path.abspath(main_folder) + "\\scripts\\Libs"
sys.path.append(libpath)
# print ("libpath = " + libpath)

# Directory to store Steady State/Dynamic result
ResultsDir = OutputDir+"\\dynamic_smib"
make_dir(ResultsDir)

#-----------------------------------------------------------------------------
# GLOBAL VARIABLES
#-----------------------------------------------------------------------------
import auxiliary_functions as af
import readtestinfo as readtestinfo
import run_simulation

# return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'SimulationSettings', 'ModelDetailsPSSE', 'SetpointsDict', 'ScenariosSMIB', 'Profiles'])
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'ModelDetailsPSSE', 'Setpoints', 'ScenariosSMIB', 'Profiles', 'OutputChannels'])
ProjectDetailsDict = return_dict['ProjectDetails']
# SimulationSettingsDict = return_dict['SimulationSettings']
PSSEmodelDict = return_dict['ModelDetailsPSSE']
SetpointsDict = return_dict['Setpoints']
ScenariosDict = return_dict['ScenariosSMIB']
ProfilesDict = return_dict['Profiles']
OutChansDict = return_dict['OutputChannels']

#-----------------------------------------------------------------------------
# Main
#-----------------------------------------------------------------------------
def main():
    pass

    scenarios = ScenariosDict.keys()
    global activeScenarios
    global Semaphore 
    
    activeScenarios=[]
    for scenario in scenarios:
        if(ScenariosDict[scenario]['run in PSS/E?']=='yes'):
            if('simulation batch' in ScenariosDict[scenario].keys()):
                if( (simulation_batches==[]) or (ScenariosDict[scenario]['simulation batch'] in simulation_batches) ):
                    activeScenarios.append(scenario)
            else:
                activeScenarios.append(scenario)
 
    #uncomment this section for debugging, to execute simulation without multiprocessing. 
    for scenario in activeScenarios:
        workspace_folder, testRun_ = createModelCopy(scenario, testRun) # SCRIPT CHANGE: create one new folder for the simulation set. In that folder, 
        # if(base_model != base_model_workspace): #workspace can be subfolder of model folder
        #     workspace_folder=workspace_folder+"\\"+os.path.basename(os.path.normpath(base_model_workspace))
        scenario_params = ScenariosDict[scenario]
        # run_simulation.run(main_folder, scenario, scenario_params, workspace_folder, testRun_, ProjectDetailsDict, PSSEmodelDict, SetpointsDict, ProfilesDict)
        run_simulation.run(ResultsDir, scenario, scenario_params, workspace_folder, testRun_, ProjectDetailsDict, PSSEmodelDict, SetpointsDict, ProfilesDict, OutChansDict)
    
    pass
    
#    l = Semaphore(1)
#    sem = Semaphore(max_processes)
#    p = Pool(processes = max_processes, initializer = initializer, initargs = (l, sem))
#    p.map(worker, activeScenarios)
    

if __name__ == '__main__':
    main()