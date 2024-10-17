# -*- coding: utf-8 -*-
"""
Created on Fri Jun 26 11:02:06 2020

CONCEPT:
    Script reads Test list from Excel spreadsheet
        test type ( frequencyProfile, voltageProfile, ORT, AngleDisturbance, VstpProfile, PstpProfile, QstpProfile, PF_stp_profile, Fault)
            Test type list can be extended for the future
        Depending on the test type the script will expect a different specific set of parameters
        depending on the test type, differnet settings will be chosen for the PSCAD "testing blocks". The script translates between the "test Type" and appropriate settings of the model blocks.
        

    test list is divided up into groups of X (usually 8) (or less if model consists of multipel project files. In order to have only one model per set, set nunmber of project files per set equal to the number of project files in the model ):   
    for every group:
        For every test scenario in group:
            script creates a model copy
            script configures model copy per the test settings and saves the copy with modified settings. This can include:
                producing "profile Files" which are placed in the workspace directory of the model copy
                settings of PSCAD "testing blocks"
                changing variables in the in verter or PPC model (for example to change operating mode from V=-control to Q-control)
                activating/deactivating layers
        for every test scenario in group: 
            project file copuy is added to master workspace and configured per scenario settings
        master workspace is saved
        script runs the test
            script saves simulation results labelled per test as csv file (possibly truncated), along with information about the test scenario

                
        
        
    Advantages compared to setting everything up in PSCAD using model copies and parallel run:
        no need to set up project copies in PSCAD and configure all of them for the different scenarios
        if model is changed, only the base copy of the model needs to be modified --> easy to re-run full study batch after the model or a setting has changed
        easy to add scenarios per copy-past in excel --> lower likely hood of mistakes in test setup, can quickly set up large test series.
        Pre-configured test series can be set up and re-used across projects (MAT, Benchmarking)
            ideally shared interface with PSS/E study tool --> scenarios only need to be set up once, to cover both softwares.
            
        if any tests are faulty, popup messages can be set to be ignored and the other tests will still run (without PSCAD getting stuck)
        

Initialisation process:
    INPUTS:
        P at pwr_init_loc
        Q at pwr_init_loc
        Voltage at v_init_loc
        Grid SCR at test_src_loc
        Grid X/R at test_src_loc    
        --> five locations: 
                pwr_init_loc --> Power flow location (for initialisation)          
                v_init_loc --> Voltage location (for intialisation) 
                test_src_loc --> location where the test source connects to the model
                pwr_ctrl_loc --> location where the PPC monitors the Power flows
                v_ctrl_loc --> location where the PPC monitors the voltage                
 
                   (  Possible cases:
                        1) all 5 locations are the same. This is the standard case. This would most likley represent the POC
                        2) Mulwala: 
                            pwr_init_loc: 22kV  side of HV/MV transformer
                            v_init_loc: 132 kV side of HV/MV transformer
                            test_src_loc: 132 kV side of HV/MV transformer
                            pwr_ctrl_loc: 22kV  side of HV/MV transformer
                            v_ctrl_loc: 132 kV side of HV/MV transformer
                        3) Coppabella:
                            pwr_init_loc: 330 kV
                            v_init_loc: 330 kV
                            test_scr_loc: 330 kV
                            prw_ctrl_loc: 132 kV
                            v_ctrl_loc: 132 kV
                        4) Clarke Creek:
                            pwr_init_loc: 275 kV 1
                            v_init_loc: 275 kV 1
                            test_scr_loc: 275 kV 1
                            pwr_ctrl_loc: 275 kV 2,3,4
                            v_ctrl_loc: 275 kV 2,3,4    )
    
        Information about TAP changer logic
        Control mode of PPC
            pf control
            direct  Q control
            direct voltage control
            voltage droop control (most common)
        PPC setpoint (in case fix)

                        
    OUTPUTS:
        Transformer taps
        PPC target (if PPC in voltage control mode, it will be the V-stp, otherwise Q or PF setpoint)
        PPC P target
        Vthevenin for Grid representation
        Grid reactance
        Grid Impedance
    
    Strategy for implementation:
        
      

@author: Mervin Kall

V2.3: Updated script from input_modify; run okay in PSCAD machine
V2.4: Separate results from Onedrive folder
V2.5: Update Vth calculation from Mainscript
        Put the path creation part into a function - similar to PSSE script
        Update Resultdir: ResultsDir = OutputDir+"\\dynamic_smib"

"""

from multiprocessing import*
import time
import datetime
import os, sys
from subprocess import*
import shutil
import math
import shelve
import random

from openpyxl import load_workbook, workbook
import json
import string

#from itertools import combination_with_replacement as cwr

from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.text import RichText
#from openpyxl.drawing.text import (Paragraph, ParagraphProperties, ChracterProperties, Font)

from openpyxl.chart import ScatterChart, Reference, Series

# libpath_0 = os.path.dirname(__file__) #directory of the script
# libpath = os.path.abspath(main_folder) + "\\scripts\\Libs"
libpath =  os.path.dirname(__file__) + "\\Libs"
# libpath = r"C:\work\Summerville (SUM)\1. Power System Studies\1. Main Test Environment\20230310_SUM_DMAT_BackBreak\PSCAD_sim\scripts\Libs"
sys.path.append(libpath)


import readtestinfo
import TOV_calc
import vth_initialisation

import pandas as pd
import copy
from win32com.client import Dispatch
import time
# timestr = time.strftime("%Y%m%d")
timestr = str(datetime.datetime.now().strftime("%Y%m%d-%H%M"))

import numpy

#-----------------------------------------------------------------------------
# Auxiliary Functions
#-----------------------------------------------------------------------------

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

def createPath(main_path_out):
    path = os.path.normpath(main_path_out)
    path_splits = path.split(os.sep) # Get the components of the path
    child_folder = r"C:" # Build up the output path from C: directory
    for i in range(len(path_splits)-1):
        child_folder = child_folder + "\\" + path_splits[i+1]
        make_dir(child_folder)
    return child_folder

##############################################################################
# USER CONFIGURABLE PARAMETERS
##############################################################################
TestDefinitionSheet = r'20240403_HSFBESS_TESTINFO_V1.xlsx'
#simulation_batches=['DMAT', 'Prof_chng', 'AEMO_fdb', 'missing', 'legend', 'SCR_chng', 'timing'] #specify batch from spreadsheet that shall be run. If empty, run all batches
# simulation_batches=['DMATsl1','DMATsl2','DMATsl3','DMATsl4','DMATsl5','DMATsl6','DMATsl','Benchmarking']
#simulation_batches=['DMATsl1','DMATsl2','DMATsl3','DMATsl4','DMATsl5','DMATsl6','DMATsl','DMATsl1_db','DMATsl2_db','DMATsl3_db','DMATsl4_db','DMATsl5_db','DMATsl6_db','DMATsl_db']
#simulation_batches=['DMATsl1_db','DMATsl2_db','DMATsl3_db','DMATsl4_db','DMATsl5_db','DMATsl6_db','DMATsl_db', 'Benchmarking_db']
#simulation_batches=['Benchmarking', 'Benchmarking_db']
# simulation_batches=['S52514']
# simulation_batches=['DMATsl1','DMATsl2','DMATsl3','DMATsl4','DMATsl5','DMATsl6','DMATsl','Benchmarking','Benchmarking2','Benchmarking3']


# simulation_batches=['S52511','S52513','S52514','S5255Iq1','S5255Iq2','S5255Iq3']
#simulation_batches=['S5255Iq1','S5255Iq3']# last batch that Dao ran
simulation_batches= ['Benchmarking3_dbg']#['Benchmarking', 'Benchmarking2', 'Benchmarking3', 'Benchmarking3_dbg','Benchmarking_frt', 'Benchmarking2_frt','Benchmarking3_frt']##'Benchmarking','shallow_fault_dbg']#, []'Benchmarking2_frt_dbg', 'Benchmarking3_frt_dbg','Benchmarking_frt_dbg']#'Benchmarking_dbg', 
#The below can alternatively be defined in the Excel sheet

overwrite = False # 
max_processes = 28 #set to the number of cores on my machine. Needs to be >= scenarioPerGroup --> increase for PSCAD machine
# testRun = '20220406_v0' #define a test name for the batch or configuration that is being tested
try:
    testRun = timestr + '_' + simulation_batches[0] #define a test name for the batch or configuration that is being tested -> link to time stamp for auto update
except:
    testRun = timestr

compiler_acronym = '.if18_x86' #This is the ending applied to the result folder, depenging on the compiler that is used
#compiler_acronym = '.gf81_x86' #This is the ending applied to the result folder, depenging on the compiler that is used
# compiler_acronym = '.if18_x86' #This is the ending applied to the result folder, depenging on the compiler that is used
scenariosPerGroup=28
# scenariosPerGroup=8
 #set number of scenarios to be included per simulation set.

##############################################################################
# Define Project Paths
##############################################################################
# current_dir = os.path.dirname(__file__) #directory of the script
# # current_dir=r"C:\work\Summerville (SUM)\1. Power System Studies\1. Main Test Environment\20230325_SUM_GridForming\PSCAD_sim\scripts" #path to PSCAD script.
# main_folder = os.path.dirname(current_dir) # Identify main_folder: to be compatible with previous version. / main folder is one level above

script_dir=os.getcwd()
main_folder=os.path.abspath(os.path.join(script_dir, os.pardir))

# Create directory for storing the results
if "OneDrive - OX2"  in main_folder: # if the current folder is online (under OneDrive - OX2), create a new directory to store the result
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
else: # if the output location is same as input location, then delete the links
    try:os.remove(main_folder + "\\model_copies.lnk")
    except: pass
    try: os.remove(main_folder + "\\result_data.lnk")
    except: pass

# Locating the existing folders
testDefinitionDir= os.path.abspath(os.path.join(main_folder, os.pardir))+"\\test_scenario_definitions"
base_model = main_folder+"\\base_model" #parent directory of the workspace folder
base_model_workspace = main_folder+"\\base_model\\20240718_HSFBESS_V1_FW10" #path of the workspace folder, formerly "workspace_folder" --> in case the workspace is located in a subdirectory of the model folder (as is the case with MUL model for example)
libpath = os.path.abspath(main_folder) + "\\scripts\\Libs"
sys.path.append(libpath)
# print ("libpath = " + libpath)

# Directory to store Steady State/Dynamic result
ResultsDir = OutputDir+"\\dynamic_smib"
make_dir(ResultsDir)

##############################################################################
# READ TEST INFO
##############################################################################
# return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'SimulationSettings', 'ModelDetailsPSCAD', 'SetpointsDict', 'ScenariosSMIB', 'Profiles'])
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'ModelDetailsPSCAD', 'Setpoints', 'ScenariosSMIB', 'Profiles'])
ProjectDetailsDict = return_dict['ProjectDetails']
# SimulationSettingsDict = return_dict['SimulationSettings']
PSCADmodelDict = return_dict['ModelDetailsPSCAD']
SetpointsDict = return_dict['Setpoints']
ScenariosDict = return_dict['ScenariosSMIB']
ProfilesDict = return_dict['Profiles']# --> All dictionaries will be available as global variables, as they are not being reinitiated in the subroutines
##############################################################################
# GLOBAL VARIABLES
##############################################################################
stpScal={} #create scaling dictionary. This will later be int
activeScenarios=[]
scenario_groups={}


#READING AND GROUPING SCENARIOS
scenarios = ScenariosDict.keys()
for scenario in scenarios:
    if(ScenariosDict[scenario]['run in PSCAD?']=='yes'):
        if('simulation batch' in ScenariosDict[scenario].keys()):
            if( (simulation_batches==[]) or (ScenariosDict[scenario]['simulation batch'] in simulation_batches) ):
                activeScenarios.append(scenario)# add keys of scenarios listed as "yes" under "run in PSCAD?" to the activeScenarios list
        else:
                activeScenarios.append(scenario)
#scenario_groups = group_scenarios(activeScenarios, scenariosPerGroup) #

activeScenario_id=0
group_cnt=1
while activeScenario_id<len(activeScenarios):
    scenario_group=[]
    scenario_cnt=0        
    while ( (activeScenario_id<len(activeScenarios)) and (scenario_cnt<scenariosPerGroup)):
        scenario_group.append(activeScenarios[activeScenario_id])
        scenario_cnt+=1
        activeScenario_id+=1        
    scenario_groups['group_'+str(group_cnt)]=scenario_group
    group_cnt+=1
##############################################################################
# FUNCTIONS
##############################################################################
def initializer(l, sem):
    global semaphore
    semaphore =l
    global taskSem
    taskSem=sem
    return

def worker(scenario_group):# i is the scenario key
    global scenario_groups
    print('now executing worker process')
    print('scenario_group: '+scenario_group)
    print(scenario_groups)
    taskSem.acquire()
    
    # avai_cer = PSCAD.get_available_certificates()
    # first_cer = avai_cer[2891396743]
    # second_cer = avai_cer[1965517995]
    # pscad.get_certificate(first_cer) # Possibly use this one for PSCAD V5
    #workspace_folder_path_local = createModelCopy(i)
    #workspace_folder_name_local = TestConfigDict['test_names'][test_list[i]-1]
    #workspace_folder_path_local = createModelCopies(scenario_group) # SCRIPT CHANGE: create one new folder for the simulation set. In that folder, 
    
    current_set_folder, testRun_ = createModelCopies(scenario_group)        
    runTest(scenario_group, current_set_folder, testRun_)
    
    #runTest(scenario_group, current_workspace_folder=workspace_folder_path_local)
    taskSem.release()
    # pscad.release_certificate()
    return

#check if string contains number
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def runTest(scenario_group, current_workspace_folder, testRun): #scenario_group is the name of the set in the dictionary "scenario_groups"; current_workspace_folder points to the folder containing the model copy that will be transformed into a simulation set
    import datetime
    import csv
    import tkinter as tk
    from tkinter import filedialog
    #import PSCAD automation config
    import logging
    import sys, os
    sys.path.append(r"C:\Program Files\Python37\Lib\site-packages") #required to run the script in Spyder  !!!LIKELY NEEDS TO CHANGE
    
    import mhi.pscad
    from mhi.pscad.utilities.file import File
    
    versions = mhi.pscad.versions()
    #LOG.info("PSCAD Versions: %s", versions)

    # Skip any 'Alpha' versions, if other choices exist
    vers = [(ver, x64) for ver, x64 in versions if ver != 'Alpha']
    if len(vers) > 0:
        versions = vers
    
    # Skip any 'Beta' versions, if other choices exist
    vers = [(ver, x64) for ver, x64 in versions if ver != 'Beta']
    if len(vers) > 0:
        versions = vers
    
    # Skip any 32-bit versions, if other choices exist
    vers = [(ver, x64) for ver, x64 in versions if x64]
    if len(vers) > 0:
        versions = vers
        
    version, x64 = sorted(versions)[-1]
    
    print("Automation Library:", mhi.pscad.VERSION)
    fortrans=mhi.pscad.fortran_versions()
    print(fortrans)
    settings = {'fortran_version': 'Intel 19.2.3787'}
    #settings = {'fortran_version': 'GFortran 8.1'}
    # settings = {'fortran_version': 'Intel® Fortran Compiler Classic 2021.7.0'}
    #fortran_ext = 
    
    # import mhrc.automation
    # import mhrc.automation.handler
    from win32com.client.gencache import EnsureDispatch as Dispatch
    # from mhrc.automation.utilities.word import Word
    # from mhrc.automation.utilities.file import File
    import win32com.client
    import shutil
    ###########################################################################
    global base_model
    global base_model_workspace
    # log info messagesk, include hh:mm:ss, level & name
    # logging.basicConfig(level=logging.INFO, datefmt='%I:%M:%S',
    #                     format="%(asctime)s %levelname)-8s %(name)-26s %(message)s")
    # logging.getLogger('mhrc.automation').setLevel(logging.WARN)
    logging.getLogger('mhi.pscad').setLevel(logging.WARNING)
    LOG=logging.getLogger('main')
    # controller=mhrc.automation.controller()
    #versions=controller.get_paramlist_names('pscad')    LOG.info("PSCAD versions: %s, versions")
    #pscad=mhrc.automation.launch_pscad(PSCADmodelDict['pscad version'])#, minimize=True)
    
    pscad=mhi.pscad.launch(version=version, x64=x64, settings=settings,) # minimize=True)
    pscad.settings(MaxConcurrentSim=8, LCP_MaxConcurrentExec=8) #the second option should indicate the number of CPU cores. However, this is per instance of PSCAD, Given every scenario is launched in a separate process in our setup, this should not matter
    #Try to acquire the lic 04/04/2023:
    avai_cer = pscad.get_available_certificates()
    # first_cer = avai_cer[1394977386] # Certificate[PSCAD 5.0.2 PRO (PSCAD V5 Pro-1 w/32 PS), 0/1] # may need to improve to recognise this certificate from the avaialbe list.
    # second_cer = avai_cer[1965517995] # Certificate[PSCAD 5.0.1 PRO (PSCAD V5 Pro-1), 0/1]
    # pscad.get_certificate(first_cer) # use this one for PSCAD V5 -> get licence
    # pscad.get_certificate(second_cer)
    # try:
    #     pscad.get_certificate(first_cer) # use this one for PSCAD V5 -> get licence 
    # except: 
    #     pscad.get_certificate(second_cer)
    
    # check installed fortran compiler versions
    # fortran_versions = controller.get_paramlist_names('fortran')
    # fortran_version = fortran_versions[-2]
    # fortran_version = controller.get_param('fortran', fortran_version)  
    pscad.settings(cl_use_advanced='true',) #fortran_version='Intel(R) Visual Fortran Compiler 17.0.2.187')#fortran_version) #!!!!!!!! WILL NEED TO BE CHANGED TO version 15
    
     
    # pscad.load(r"C:\Users\Mervin Kall\OneDrive - ESCO Pacific\Mulwala\20200622_PSCAD_SMIB_STUDIES\v3_I1_PSCAD_model\IJX\20200528_MUL_SMIB_PSCAD.pswx")
    # pscad.run_all_simulation_sets()
    # pscad.quit()
    project_files={}
    for key in PSCADmodelDict:
        if('Project' in key and key!='Projects'):
            project_files[key]=PSCADmodelDict[key] # for the case where multiple instances of the model are in one workspace (simulation batch) the project file names need to be mapped to the changed changed names
            
    # prepate example cases. These are copies of the model set up to produce the test results of a given scenario, but will not be used for the actual simulation and only serve as examples for debugging or to be shared with 
    #check if example case is to be created or not
    if("create example cases" in PSCADmodelDict.keys()):
        if(PSCADmodelDict['create example cases']=='yes'):
            for scenario in scenario_groups[scenario_group]:
                """SCRIPT CHANGE:    
                    -open the individual model copy, parametrise per parameters of the scenario and then close
                    
                    -open the model copy for the current simulation set
                    -creates a project file copy of the base model within the batch workspace--> make sure attributes keep the same names
                    -parametrises the project file(s) copy(s) per the parameters of the scenario
                    -adds the new project file to simulation set
                """
                
                workspaceFileDir=ModelCopyDir+"\\"+testRun+"\\"+scenario
                if(base_model != base_model_workspace): #workspace can be subfolder of model folder
                    workspaceFileDir=workspaceFileDir+"\\"+os.path.basename(os.path.normpath(base_model_workspace))
                
                workspaceFilePath=workspaceFileDir+"\\"+PSCADmodelDict['Workspace File Name']
                # models={workspaces:[current_workspace_folder+"\\"+PSCADmodelDict['Workspace File Name'], #This is the joint workspace file including the batch
                #                     individual_project_folder],
                #         project_files: [project1+"-"+str(scenario), PSCADmodelDict['Project1']],
                #         }
                #setScenarioParams(ScenariosDict[scenario], workspaceFilePath, project_files) #set scenario parameters for individual model copies and save those
                    #-------------------------------------------------------------------------
                #global PSCADmodelDict
                global ProjectDetailsDict
                global SetpointsDict
                global ProfilesDict
                #OPEN MODEL COPY
                pscad.load(workspaceFilePath)
                # create new simulations set 
                scenario_params=ScenariosDict[scenario]
                
                #-------------------------------------------------------------------------
                #DEFAULT INITIALISATION
                # Set profile generator to use default values, set GridSource to SourceNormal
                project=pscad.project(project_files['Project1']) #This assumes that the test equipment is always place in the project listed under "Project1" in the FileSettings sheet
                #Setting Grid Source to SourceNormal, define Vbase and Fbase, based on "ProjectDetails" Sheet
                if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(PSCADmodelDict['GridSourceID'])
                if(ProjectDetailsDict['VbaseTestSrc']!=''):
                    VbaseTestSrc = ProjectDetailsDict['VbaseTestSrc']
                else: 
                    VbaseTestSrc = ProjectDetailsDict['VPOCkv']
                parameters = {'mode':0, 'Fbase': ProjectDetailsDict['Fbase'], 'Vbase':VbaseTestSrc}
                comp_handle.set_parameters(**parameters)
                
                #Setting profile generator to use default values
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                parameters = {'VstpMethod':3, 'QstpMethod':3,'Q1stpMethod':3, 'PFstpMethod':3,'PF1stpMethod':3, 'PstpMethod':3, 'P1stpMethod':3}
                comp_handle.set_parameters(**parameters)
                
                #Initialise fault component to 'off'
                if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(PSCADmodelDict['FaultBlockID'])
                parameters = {'status':0}
                comp_handle.set_parameters(**parameters)
                
                SCRMVAbase=ProjectDetailsDict['SCRMVA']
            
                
                #-------------------------------------------------------------------------
                #SET COMPONENT PARAMETERS PER INITIALISATION
                    #Component settings per initialisation sheet
                        # THis includes Test source and test profile generator
                        # THe settings of these MUST be defined
                            #The profileGenerator must be set to "default" and 
                            #THe internal default value must be selected so that the system initialises to the intended setpoint
                    #Grid source per SCR, X/R and voltage, --> also in initialisation sheet
                    # base MVA and base voltage to be set from "ProjectDetails" sheet
                    #--> long term: create a routine that reads the PSS/E model file OR integrate sheet with "Working model details spreadsheet" to have a proper definition of the buses and allow more complex automated initialisation
                #look into scenarioDict to determine Setpoint ID, then iterate over all components that need to be set in SetpointDict Entry
                stpID=scenario_params['setpoint ID']
                setpoint_params = SetpointsDict[stpID]
                comp_settings = setpoint_params['comp_settings']
                #iterate over all components from initialisation tab --> in the long term, the calculation of initialisation parameters could at least partly be automated
            
                # for comp_set_id in range(0, len(comp_settings['values'])):
                #     project=pscad.project(project_files[comp_settings['ProjectNo'][comp_set_id]])    #get project handler
                #     canvas_handle = project.user_canvas(comp_settings['Module'][comp_set_id])
                #     comp_handle = canvas_handle.user_cmp(comp_settings['PSCAD ID'][comp_set_id])
                #     parameters = { comp_settings['Symbol'][comp_set_id]:comp_settings['values'][comp_set_id]}# prepare dict.                  
                #     comp_handle.set_parameters(**parameters)
                    
                for comp_set_id in range(0, len(comp_settings['values'])):
                    project=pscad.project(project_files[comp_settings['ProjectNo'][comp_set_id]])    #get project handler
                    canvas_handle = project.canvas(comp_settings['Module'][comp_set_id])
                    canvas=canvas_handle
                    comp_handle = canvas.component(int(comp_settings['PSCAD ID'][comp_set_id]))
                    parameters = { comp_settings['Symbol'][comp_set_id]:comp_settings['values'][comp_set_id]}# prepare dict.
                    #print(parameters)
                    print(comp_handle)
                    comp_handle.set_parameters(**parameters)
                    
                #Add the settings for second SCR if applicable. Otherwise set to "no"
                if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(PSCADmodelDict['GridSourceID'])


                
                #calculate Vth and set automatically
                Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                parameters={'Vpu':Vth_prev}
                comp_handle.set_parameters(**parameters)   
                
                # 12/4/2022: Set Q control mode:
                # PpcVArMod=SetpointsDict[scenario_params['setpoint ID']]['PpcVArMod']
                # parameters={'QControlMode':PpcVArMod}
                # comp_handle.set_parameters(**parameters)
                
                if('small' in scenario):
                    if( (scenario_params['Secondary SCL time']!='') and (scenario_params['Secondary SCL']!='') and (scenario_params['Secondary X_R']!='') ):
                        #determine Vth_new and set parameters in Grid source accordingly
                        Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=scenario_params['Secondary X_R'], SCC=scenario_params['Secondary SCL'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        parameters={'enableSecSCR':1, 'switchTime': scenario_params['Secondary SCL time'], 'Vpu_new': Vth_new, 'X_R_new':scenario_params['Secondary X_R'], 'SCL_new':scenario_params['Secondary SCL'], 'angle_offset':ang_new-ang}

                    elif('init_help' in scenario_params.keys()):
                        if(scenario_params['init_help']==1):
                            Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            parameters={'enableSecSCR':1, 'switchTime': 2.0, 'Vpu': Vth_prev, 'SCL':99999.0, 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

                        if(scenario_params['init_help']==2): # in case low SCR, 'Vpu'is fixed to 1.0 as Vth_prev may go wrong with low SCR condition. Also, 'SCL' is not fixed to 99999.0, but updated from the setpoint
                            # Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            parameters={'enableSecSCR':1, 'switchTime': 1.5, 'Vpu': 1.0, 'SCL':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}


                elif( ('large' in scenario) or ('tov' in scenario) ):
                    if((scenario_params['SCL_post']!='') and (scenario_params['X_R_post']!='')):
                        switchTime=scenario_params['Ftime']+scenario_params['Fduration']#detemrin time at which fault ends (and SCR switches)
                        Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        Vth_new, ang_new=vth_initialisation.calc_Vth_pu(X_R=scenario_params['X_R_post'], SCC=scenario_params['SCL_post'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        parameters={'enableSecSCR':1, 'switchTime': switchTime, 'Vpu_new': Vth_new, 'X_R_new':scenario_params['X_R_post'], 'SCL_new':scenario_params['SCL_post'], 'angle_offset':ang_new-ang} #knowing previous Vth angle was 0 in the model, the new angle is chosen in relation to what the previous anbgle shoudl have been.
                    elif('init_help' in scenario_params.keys()):
                        if(scenario_params['init_help']==1):
                            Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            parameters={'enableSecSCR':1, 'switchTime': 2.0, 'Vpu': Vth_prev,'SCL':99999.0,'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}
                        if(scenario_params['init_help']==2):
                            # Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                            parameters={'enableSecSCR':1, 'switchTime': 1.5, 'Vpu': 1.0, 'SCL':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

                elif('init_help' in scenario_params.keys()):
                    if(scenario_params['init_help']==1):
                        Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        parameters={'enableSecSCR':1, 'switchTime': 2.0, 'SCL':99999.0,'Vpu': Vth_prev, 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

                    if(scenario_params['init_help']==2):
                        # Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                        parameters={'enableSecSCR':1, 'switchTime': 1.5, 'Vpu': 1.0, 'SCL':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

                else:
                    #set ensableSecSCR to 'no'
                    parameters={'enableSecSCR':0}
                    pass
                comp_handle.set_parameters(**parameters)    
                    
                #-------------------------------------------------------------------------
                #SET TEST COMPONENTS PER SCENARIO DEFINITION
                simulation_duration = PSCADmodelDict['default_sim_duration']        
                project=pscad.project(project_files['Project1'])    # The test components always need to be placed in file "Project1"

                #NEW addition to allow for multiple test types and associated profiles in a single scenario.
                for test_type_id in range (0, len(scenario_params['Test Type'])):
                    test_type=scenario_params['Test Type'][test_type_id]
                    test_profile_name=scenario_params['Test profile'][test_type_id] 
                    #test type
                        # for fault scenarios use context menu of the test block
                    #if test type is one of the known types, select appropriate setting in context menu of every test block
                    if (test_type=='F_profile'): 
                        if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                            canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['GridSourceID'])
                        #Set grid source to F_profile mode --> will disable all other modes
                        parameters = {'mode': 1, 'FprofileMethod':0, 'FprofileOffset':0} #set operating mode to F_profile and profile entry method to 'file' and offset to 'none'
                        # set scaling option in context menu to "Hz"
                        #set offset to none (in the long term, might add that function to the excel sheet)        
                        parameters['FprofileScal']=0 #set mode to Hz
                        #write settings into component
                        comp_handle.set_parameters(**parameters)    
                        
                        
                
                    elif(test_type=='V_profile'):
                        if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                            canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['GridSourceID'])
                        #Set grid source to V_profile mode
                        #set scaling to Vbase (absolute) or Vpu (relative)
                        parameters = {'mode': 2, 'VprofileMethod':0, 'VprofileOffset':0} #set operating mode to F_profile and profile entry method to 'file' and offset to 'none'
                        # set scaling option in context menu to "Hz"
                        #set offset to none (in the long term, might add that function to the excel sheet)        
                        if(ProfilesDict[scenario_params['Test profile']]['scaling']=='relative'):
                            parameters['VprofileScal']=2 #set mode to interpret profile as expressed in pu on Vbase x Vpu
                        else:
                            parameters['VprofileScal']=1 #set mode to interpret profile as expressed in pu on Vbase
                        comp_handle.set_parameters(**parameters)    
                            
                    elif(test_type=='ANG_profile'):
                        if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                            canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['GridSourceID'])
                        #Set grid source to F_profile mode --> will disable all other modes
                        parameters = {'mode': 4, 'PHprofileMethod':0}
                        comp_handle.set_parameters(**parameters)
                        
                    elif(test_type== 'ORT'):
                        if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                            canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['GridSourceID'])
                        #Set grid source to F_profile mode --> will disable all other modes
                        parameters = {'mode': 3, 'MagBase':float(scenario_params['Disturbance Magnitude']), 'Fdist': float(scenario_params['Disturbance Frequency']), 'PhDistMag': float(scenario_params['PhaseOsc Magnitude']), 'OscStartTime': float(scenario_params['time'])} ##ADD HERE THE PARAMETERS FOR THE DISTURBANCE
                        comp_handle.set_parameters(**parameters)
                        simulation_duration = 20 # 30/11/2022: run full 20sec as requested from AEMO
                    
                    #Setpoint profiles
                    #
                    elif(test_type== 'V_stp_profile'):
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'VstpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        VstDefault=comp_handle.get_parameters()['VstpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['V_stp_profile']=VstDefault
                        else:
                            stpScal['V_stp_profile']=1.0        
                        
                    elif(test_type== 'Q_stp_profile'):
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                            canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'QstpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        QstDefault=comp_handle.get_parameters()['QstpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['Q_stp_profile']=QstDefault
                        else:
                            stpScal['Q_stp_profile']=1.0 
                                        
                    elif(test_type== 'Q1_stp_profile'):
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                            canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'Q1stpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        QstDefault=comp_handle.get_parameters()['Q1stpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['Q1_stp_profile']=QstDefault
                        else:
                            stpScal['Q1_stp_profile']=1.0 
                            
                    elif(test_type==  'PF_stp_profile'): 
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'PFstpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        PFstDefault=comp_handle.get_parameters()['PFstpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['PF_stp_profile']=PFstDefault
                        else:
                            stpScal['PF_stp_profile']=1.0  
                            
                    elif(test_type==  'PF1_stp_profile'): 
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'PF1stpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        PFstDefault=comp_handle.get_parameters()['PF1stpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['PF1_stp_profile']=PFstDefault
                        else:
                            stpScal['PF1_stp_profile']=1.0  
                            
                    elif(test_type== 'P_stp_profile'):
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'PstpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        PstDefault=comp_handle.get_parameters()['PstpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['P_stp_profile']=PstDefault
                        else:
                            stpScal['P_stp_profile']=1.0    
                    
                    elif(test_type== 'P1_stp_profile'):
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'P1stpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        PstDefault=comp_handle.get_parameters()['P1stpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['P1_stp_profile']=PstDefault
                        else:
                            stpScal['P1_stp_profile']=1.0    
                            
                    elif(test_type== 'Auxiliary_profile'):
                        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
                        # For stp profile, set the corresponging profile in the 
                        #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                        #elif scaling = absolute: use parameterBase as scaling factor
                        parameters = {'AUXstpMethod':0} #set profile entyr method to 'file'
                        comp_handle.set_parameters(**parameters)  
                        
                        AUXstpDefault=comp_handle.get_parameters()['AUXstpDefault']
                        if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                            stpScal['Auxiliary_profile']=AUXstpDefault
                        else:
                            stpScal['Auxiliary_profile']=1.0 
                            
                    elif(test_type=='Fault'):
                        if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['FaultBlockID'])
                        
                        if(scenario_params['Ftype']=='3PHG'):fault_type=0
                        elif(scenario_params['Ftype']=='2PHG'):fault_type=1
                        elif(scenario_params['Ftype']=='1PHG'):fault_type=2
                        elif(scenario_params['Ftype']=='L-L'):fault_type=3
                        
                        parameters = {'status':1,'N_faults':1, 'Time1':scenario_params['Ftime'], 'Duration1':scenario_params['Fduration'], 'Type1':fault_type} #set to block to 'active' and single fault
                        #set impedance depending on if it is defined directly or via residual voltage
                        if(scenario_params['Fault X_R']==''):
                            scenario_params['Fault X_R']=3 #default to 3 if not defined
                        if(scenario_params['F_Impedance']!=''): #impedance is defined
                            parameters['Impedance1']=scenario_params['F_Impedance']
                            parameters['X_R1']=scenario_params['Fault X_R']
                        elif(scenario_params['Vresidual']!=''):
                            Zfault, X_R_fault = calc_fault_impedance(scenario_params['Vresidual'],scenario_params['Fault X_R'],  SetpointsDict[scenario_params['setpoint ID']]['V_POC'],VbaseTestSrc, SCRMVAbase, SetpointsDict[scenario_params['setpoint ID']]['SCR'], SetpointsDict[scenario_params['setpoint ID']]['X_R'])#
                            parameters['Impedance1']= Zfault  
                            parameters['X_R1']=X_R_fault
                        else:
                            parameters['Impedance1']=0.01 #default to 0 Ohm fault
                            parameters['X_R1']=scenario_params['Fault X_R']
                
                        comp_handle.set_parameters(**parameters)
                        
                        simulation_duration = round(scenario_params['Ftime']+scenario_params['Fduration']+5.0)
                        # set up fault block
                                #with data from table if type is fault
                                #with data from table if type is Multifault
                                # with data from table AND data from auto-generated Multifault profile if type is Mutlifault_random            
                        # set scenario duration to Ftime+Fduration+5s if type=Fault
                        # set scenario duration to Ftime+Fduration+5s of last fault in series if Type=Multifault
                        # set scenario duration to Ftime+ProfileLength+5s if type=Multifault_random
                        
                    elif(test_type=='Multifault'):
                        if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['FaultBlockID'])
                        
                        N_faults = len(scenario_params['Ftype'])
                        parameters = {'status':1, 'N_faults': N_faults}
                        
                        for faultID in range(0, len(scenario_params['Ftype'])):
                            if(scenario_params['Ftype'][faultID]=='3PHG'):fault_type=0
                            elif(scenario_params['Ftype'][faultID]=='2PHG'):fault_type=1
                            elif(scenario_params['Ftype'][faultID]=='1PHG'):fault_type=2
                            elif(scenario_params['Ftype'][faultID]=='L-L'):fault_type=3
                            
                            if(scenario_params['Fault X_R'][faultID]==''):
                                scenario_params['Fault X_R'][faultID]=3 #default fault X_R to 3
                            if(scenario_params['F_Impedance'][faultID] !=''): #If impedance is defined, set it directly                
                                parameters['Impedance'+str(faultID+1)]=scenario_params['F_Impedance'][faultID]
                                parameters['X_R'+str(faultID+1)]=scenario_params['Fault X_R'][faultID]
                            elif(scenario_params['Vresidual'][faultID] !=''):
                                Zfault, X_R_fault=calc_fault_impedance(scenario_params['Vresidual'][faultID],scenario_params['Fault X_R'][faultID], SetpointsDict[scenario_params['setpoint ID']]['V_POC'],VbaseTestSrc, SCRMVAbase, SetpointsDict[scenario_params['setpoint ID']]['SCR'], SetpointsDict[scenario_params['setpoint ID']]['X_R'])#
                                parameters['Impedance'+str(faultID+1)]=Zfault
                                parameters['X_R'+str(faultID+1)]=X_R_fault
                            else:
                                parameters['Impedance'+str(faultID+1)]=0.01 #default to 0 Ohm fault
                                scenario_params['Fault X_R'][faultID]
                            
                            parameters['Time'+str(faultID+1)]=scenario_params['Ftime'][faultID]
                            parameters['Duration'+str(faultID+1)]=scenario_params['Fduration'][faultID]
                            parameters['Type'+str(faultID+1)]=scenario_params['Ftype'][faultID]
                            #parameters['Impedance'+str(faultID)]=scenario_params['F_Impedance'][faultID]
                            #parameters['X_R'+str(faultID)]=scenario_params['Fault X_R'][faultID]
                    
                        comp_handle.set_parameters(**parameters)   
                        simulation_duration = round(scenario_params['Ftime'][faultID]+scenario_params['Fduration'][faultID]+5.0) #faultID points to the last fault in the sequence. Take the values form this fault to determine the length of the simulation
                    
                    elif(test_type=='Multifault_random'):
                        random_times=[0.01, 0.01, 0.2, 0.2, 0.5, 0.5, 0.75, 1, 1.5, 2, 2, 3, 5, 7, 10]
                        random.shuffle(random_times)
                        random_fault_duration=[0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.22, 0.22, 0.22, 0.22, 0.22, 0.22, 0.43]
                        random.shuffle(random_fault_duration)
                        Zgrid=math.pow((ProjectDetailsDict['VbaseTestSrc']*1000),2)/setpoint_params['GridMVA']/1000000 #Zgrid
                        random_impedances=[0,0,0,0,0,0,0, 3,3,3,3,3, 2,2,2]
                        random.shuffle(random_impedances)
                        random_fault_types=[2,2,2,2,2,2,1,1,1,1,1,1,1,0,0]
                        random.shuffle(random_fault_types)
                        
                        if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                            canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['FaultBlockID'])
                        parameters = {'status':1, 'N_faults': 15}
                        faultTimeOffset=scenario_params['Ftime']
                        for faultID in range(0, 15):
                            parameters['Impedance'+str(faultID+1)]=random_impedances[faultID]*Zgrid
                            parameters['X_R'+str(faultID+1)]=setpoint_params['X_R']
                            parameters['Time'+str(faultID+1)]=faultTimeOffset
                            parameters['Duration'+str(faultID+1)]=random_fault_duration[faultID]
                            parameters['Type'+str(faultID+1)]=random_fault_types[faultID]                
                            faultTimeOffset=faultTimeOffset+random_fault_duration[faultID]+random_times[faultID]            
                        comp_handle.set_parameters(**parameters)   
                        simulation_duration = round(faultTimeOffset+5.0) 
                    
                    elif(test_type=='TOV'):
                        if(is_number(scenario_params['Capacity(F)']) ):
                            capacity=float(scenario_params['Capacity(F)'])
                        else:
                            Qinj, capacity = TOV_calc.calc_capacity(setpoint_params['GridMVA'], setpoint_params['X_R'], setpoint_params['P'], setpoint_params['Q'], setpoint_params['V_POC'], ProjectDetailsDict['VbaseTestSrc']*1000, scenario_params['Vresidual'])
                        capacity_uF=math.pow(10,6)*capacity #scale to uF
                        if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                                canvas_handle = project.user_canvas("Main")
                        else:
                            canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['FaultBlockID'])            
                        parameters = {'status':1, 'N_faults': 1, 'Duration1':scenario_params['Fduration'], 'Time1':scenario_params['time'], 'Type1':4, 'Capacity1': capacity_uF}    
                        
                        comp_handle.set_parameters(**parameters)   
                        simulation_duration = round(scenario_params['time']+scenario_params['Fduration']+5.0) #faultID points to the last fault in the sequence. Take the values form this fault to determine the length of the simulation
            
                        
                    #-------------------------------------------------------------------------
                    #GENERATE TEST PROFILE FILES (either containing the test profile or default (to make the test library components work even if profile not required) )
                        # for POC profile or stp profile tests, simply copy the profiles from the excel sheet
                    if('Test profile' in scenario_params.keys()): #check if the test type allows for a test profile
                        if(test_profile_name !=''):#check if a test profile is actually listed
                            if(test_profile_name in ProfilesDict.keys()): #check if profile has been properly defined
                                simulation_duration = max(profileToFile(test_profile_name, test_type, workspaceFileDir), simulation_duration) #provide 
                            else: print("ERROR: listed test profile has not been defined")
                        else: print("LOG: No test profile defined")
            
                #-------------------------------------------------------------------------
                #SET SIMULATION PARAMETERS (partly depending on scenario definition)
                #simulation_duration =1# FOR DEBUGGING ONLY!!
                    #ScenarioDuration if defined, otherwise default based on test category and profile
                    #make sure "save to file" option is activated
                for project_id in range(1,PSCADmodelDict['Projects']+1):
                    project=pscad.project(project_files['Project'+str(project_id)])
                    # set simulation duration, set plot time step and activate "save channels to disk"
                    parameters={'time_duration':simulation_duration, 'sample_step':float(PSCADmodelDict['plotStep_us']), 'PlotType':1}
                    project.set_parameters(**parameters)
                    
                #save model as it is
                for project_id in range(1,PSCADmodelDict['Projects']+1):
                    project=pscad.project(project_files['Project'+str(project_id)])
                    project.save()
    
    ###########################################################################
    ###########################################################################
    #set up workspace including all scenarios of scenario_group

    #OPEN MODEL COPY
    workspaceFileDir=current_workspace_folder
    if(base_model != base_model_workspace): #workspace can be subfolder of model folder
        workspaceFileDir=workspaceFileDir+"\\"+os.path.basename(os.path.normpath(base_model_workspace))    
    workspaceFilePath=workspaceFileDir+"\\"+PSCADmodelDict['Workspace File Name']
    print("current_workspace_folder: "+current_workspace_folder)
    print ("workspaceFilePath: "+workspaceFilePath)
    print("workspaceFileDir: "+workspaceFileDir)
    print('sleeping 10s now')

    time.sleep(10)
    pscad.load(workspaceFilePath)
    print('loading project copy and sleeping another 20s')
    #time.sleep(10)
    for scenario_cnt in range(0, len(scenario_groups[scenario_group])):
        scenario = scenario_groups[scenario_group][scenario_cnt]
        project_files={}
        project_files_orig={}
        for key in PSCADmodelDict:
            if('Project' in key and key!='Projects'):
                project_files[key]=PSCADmodelDict[key]+'_'+scenario # for the case where multiple instances of the model are in one workspace (simulation batch) the project file names need to be mapped to the changed changed names
                project_files_orig[key]=PSCADmodelDict[key]
                #load original project file and save under different name for the 
                if(scenario_cnt>0):
                    pscad.load(workspaceFileDir+"\\"+PSCADmodelDict[key]+'.pscx')
                project=pscad.project(PSCADmodelDict[key])# Select base project to retrieve handle
                project.save_as(project_files[key])
                #.sleep(5)
                #In case an additional path to /dll files has been specified, copy those to the (yet to be created) run time folder.
                if('DllPath' in PSCADmodelDict.keys()):
                    if(PSCADmodelDict['DllPath']!=''): #dll folder path entry exists and has been specified
                        sourcepath= main_folder+"\\"+PSCADmodelDict['DllPath']
                        targetpath=workspaceFileDir+"\\"+project_files[key]+'.'+PSCADmodelDict['compiler_short']
                        #try and create target directory
                        try:
                            shutil.copytree(sourcepath, targetpath)
                        except:
                            print("Creation of the directory %s failed" % targetpath)
                        else:
                            print("Successfully created the directory %s" % targetpath)
                        #copy all files from source path ending with .dll or .lib to target directory
                pscad.load(workspaceFileDir+"\\"+project_files[key]+'.pscx')
        
        createDefaultProfiles(workspaceFileDir, scenario)
        
                
        
        #scenario =scenario_groups[scenario_group][scenario_cnt]
        #add project file (copy) to workspace
        #set_scenario_params()# set scenario parameters in project file in.
            #if the number of project files per scenario is equal to the number of project files in the model, do not rename the project files and simply use the base model that's already available.
                    #--> otherwise TL interfaces might cause trouble.
        #add simulation set and add project file to simulations set.
        #if simulations et does not yet exist, create it.
                         
    
        #-------------------------------------------------------------------------

        # create new simulations set       
        scenario_params = ScenariosDict[scenario]
        #if(scenario_cnt==0):# first model can use the existing project file(s)
            
        
        #-------------------------------------------------------------------------
        #DEFAULT INITIALISATION
        # Set profile generator to use default values, set GridSource to SourceNormal
        project=pscad.project(project_files['Project1']) #This assumes that the test equipment is always place in the project listed under "Project1" in the FileSettings sheet
        #Setting Grid Source to SourceNormal, define Vbase and Fbase, based on "ProjectDetails" Sheet
        if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
            canvas_handle = project.user_canvas("Main")
        else:
            canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['GridSourceID'])
        if(ProjectDetailsDict['VbaseTestSrc']!=''):
            VbaseTestSrc = ProjectDetailsDict['VbaseTestSrc']
        else: 
            VbaseTestSrc = ProjectDetailsDict['VPOCkv']
        parameters = {'mode':0, 'Fbase': ProjectDetailsDict['Fbase'], 'Vbase':VbaseTestSrc}
        comp_handle.set_parameters(**parameters)
        
        #Setting profile generator to use default values
        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
            canvas_handle = project.user_canvas("Main")
        else:
            canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['StpBlockID'])
        parameters = {'VstpMethod':3, 'QstpMethod':3,'Q1stpMethod':3, 'PFstpMethod':3,'PF1stpMethod':3, 'PstpMethod':3, 'P1stpMethod':3}
        comp_handle.set_parameters(**parameters)
        
        #Initialise fault component to 'off'
        if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
            canvas_handle = project.user_canvas("Main")
        else:
            canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
        comp_handle = canvas_handle.user_cmp(PSCADmodelDict['FaultBlockID'])
        parameters = {'status':0}
        comp_handle.set_parameters(**parameters)
        
        SCRMVAbase=ProjectDetailsDict['SCRMVA']
    
        
        #-------------------------------------------------------------------------
        #SET COMPONENT PARAMETERS PER INITIALISATION
            #Component settings per initialisation sheet
                # THis includes Test source and test profile generator
                # THe settings of these MUST be defined
                    #The profileGenerator must be set to "default" and 
                    #THe internal default value must be selected so that the system initialises to the intended setpoint
            #Grid source per SCR, X/R and voltage, --> also in initialisation sheet
            # base MVA and base voltage to be set from "ProjectDetails" sheet
            #--> long term: create a routine that reads the PSS/E model file OR integrate sheet with "Working model details spreadsheet" to have a proper definition of the buses and allow more complex automated initialisation
        #look into scenarioDict to determine Setpoint ID, then iterate over all components that need to be set in SetpointDict Entry
        stpID=scenario_params['setpoint ID']
        setpoint_params = SetpointsDict[stpID]
        comp_settings = setpoint_params['comp_settings']
        #iterate over all components from initialisation tab --> in the long term, the calculation of initialisation parameters could at least partly be automated
    
        for comp_set_id in range(0, len(comp_settings['values'])):
            project=pscad.project(project_files[comp_settings['ProjectNo'][comp_set_id]])    #get project handler
            canvas_handle = project.canvas(comp_settings['Module'][comp_set_id])
            canvas=canvas_handle
            comp_handle = canvas.component(int(comp_settings['PSCAD ID'][comp_set_id]))
            parameters = { comp_settings['Symbol'][comp_set_id]:comp_settings['values'][comp_set_id]}# prepare dict.
            #print(parameters)
            print(comp_handle)
            comp_handle.set_parameters(**parameters)
            
        #Add the settings for second SCR if applicable. Otherwise set to "no"
        if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                canvas_handle = project.user_canvas("Main")
        else:
            canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
        comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['GridSourceID']))

        #calculate Vth and set automatically
        Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
        parameters={'Vpu':Vth_prev}
        comp_handle.set_parameters(**parameters)           
        if('small' in scenario):
            if( (scenario_params['Secondary SCL time']!='') and (scenario_params['Secondary SCL']!='') and (scenario_params['Secondary X_R']!='') ):
                #determine Vth_new and set parameters in Grid source accordingly
                Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=scenario_params['Secondary X_R'], SCC=scenario_params['Secondary SCL'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                parameters={'enableSecSCR':1, 'switchTime': scenario_params['Secondary SCL time'], 'Vpu_new': Vth_new, 'X_R_new':scenario_params['Secondary X_R'], 'SCL_new':scenario_params['Secondary SCL'], 'angle_offset':ang_new-ang}
            elif('init_help' in scenario_params.keys()):
                if(scenario_params['init_help']==1):
                    Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    parameters={'enableSecSCR':1, 'switchTime': 2.0, 'SCL':99999.0,'Vpu': Vth_prev,'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}
                if(scenario_params['init_help']==2):
                    # Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    parameters={'enableSecSCR':1, 'switchTime': 1.5, 'Vpu': 1.0, 'SCL':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

                pass
        elif( ('large' in scenario) or ('tov' in scenario) ):
            if((scenario_params['SCL_post']!='') and (scenario_params['X_R_post']!='')):
                switchTime=scenario_params['Ftime']+scenario_params['Fduration']#detemrin time at which fault ends (and SCR switches)
                Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                Vth_new, ang_new=vth_initialisation.calc_Vth_pu(X_R=scenario_params['X_R_post'], SCC=scenario_params['SCL_post'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                parameters={'enableSecSCR':1, 'switchTime': switchTime, 'Vpu_new': Vth_new, 'X_R_new':scenario_params['X_R_post'], 'SCL_new':scenario_params['SCL_post'], 'angle_offset':ang_new-ang} #knowing previous Vth angle was 0 in the model, the new angle is chosen in relation to what the previous anbgle shoudl have been.
            elif('init_help' in scenario_params.keys()):
                if(scenario_params['init_help']==1):
                    Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    parameters={'enableSecSCR':1, 'switchTime': 2.0,'SCL':99999.0, 'Vpu': Vth_prev,'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

                if(scenario_params['init_help']==2):
                    # Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                    parameters={'enableSecSCR':1, 'switchTime': 1.5, 'Vpu': 1.0, 'SCL':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

        
        elif('init_help' in scenario_params.keys()):
            if(scenario_params['init_help']==1):
                Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                parameters={'enableSecSCR':1, 'switchTime': 2.0,'SCL':99999.0, 'Vpu': Vth_prev,'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

            if(scenario_params['init_help']==2):
                # Vth_prev, ang = vth_initialisation.calc_Vth_pu(X_R=3.0, SCC=99999.0, Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                Vth_new, ang_new = vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
                parameters={'enableSecSCR':1, 'switchTime': 1.5, 'Vpu': 1.0, 'SCL':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'Vpu_new': Vth_new, 'X_R_new':SetpointsDict[scenario_params['setpoint ID']]['X_R'], 'SCL_new':SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], 'angle_offset':ang_new-ang}

        
        else:
            #set ensableSecSCR to 'no'
            parameters={'enableSecSCR':0}
            pass
        comp_handle.set_parameters(**parameters)    

            
        #-------------------------------------------------------------------------
        #SET TEST COMPONENTS PER SCENARIO DEFINITION
        simulation_duration = PSCADmodelDict['default_sim_duration']        
        project=pscad.project(project_files['Project1'])    

        #NEW addition to allow for multiple test types and associated profiles in a single scenario.
        for test_type_id in range (0, len(scenario_params['Test Type'])):
            test_type=scenario_params['Test Type'][test_type_id]
            test_profile_name=scenario_params['Test profile'][test_type_id]  

                #test type
                    # for fault scenarios use context menu of the test block
            #if test type is one of the known types, select appropriate setting in context menu of every test block
            if (test_type=='F_profile'): 
                if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['GridSourceID']))
                #Set grid source to F_profile mode --> will disable all other modes
                parameters = {'mode': 1, 'FprofileMethod':0, 'FprofileOffset':0} #set operating mode to F_profile and profile entry method to 'file' and offset to 'none'
                # set scaling option in context menu to "Hz"
                #set offset to none (in the long term, might add that function to the excel sheet)        
                parameters['FprofileScal']=0 #set mode to Hz
                #write settings into component
                comp_handle.set_parameters(**parameters)        
        
            elif(test_type=='V_profile'):
                if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['GridSourceID']))
                #Set grid source to V_profile mode
                #set scaling to Vbase (absolute) or Vpu (relative)
                parameters = {'mode': 2, 'VprofileMethod':0, 'VprofileOffset':0} #set operating mode to F_profile and profile entry method to 'file' and offset to 'none'
                # set scaling option in context menu to "Hz"
                #set offset to none (in the long term, might add that function to the excel sheet)        
                if(ProfilesDict[scenario_params['Test profile']]['scaling']=='relative'):
                    parameters['VprofileScal']=2 #set mode to interpret profile as expressed in pu on Vbase x Vpu
                else:
                    parameters['VprofileScal']=1 #set mode to interpret profile as expressed in pu on Vbase
                comp_handle.set_parameters(**parameters)    
                    
            elif(test_type=='ANG_profile'):
                if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['GridSourceID']))
                #Set grid source to F_profile mode --> will disable all other modes
                parameters = {'mode': 4, 'PHprofileMethod':0}
                comp_handle.set_parameters(**parameters)
                
            elif(test_type== 'ORT'):
                if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['GridSourceID']))
                #Set grid source to F_profile mode --> will disable all other modes
                parameters = {'mode': 3, 'MagDist':float(scenario_params['Disturbance Magnitude']),'MagBase':1, 'Fdist': float(scenario_params['Disturbance Frequency']), 'PhDistMag': float(scenario_params['PhaseOsc Magnitude']), 'OscStartTime': float(scenario_params['time'])} ##ADD HERE THE PARAMETERS FOR THE DISTURBANCE
                comp_handle.set_parameters(**parameters)
                simulation_duration = 20 # 30/11/2022: run full 20sec as requested from AEMO
            
            #Setpoint profiles
            #
            elif(test_type== 'V_stp_profile'):
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'VstpMethod':0} #set profile entyr method to 'file'
                comp_handle.set_parameters(**parameters)  
                
                VstDefault=comp_handle.get_parameters()['VstpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['V_stp_profile']=VstDefault
                else:
                    stpScal['V_stp_profile']=1.0        
                
            elif(test_type== 'Q_stp_profile'):
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'QstpMethod':0} #set profile entyr method to 'file'
                comp_handle.set_parameters(**parameters)  
                
                QstDefault=comp_handle.get_parameters()['QstpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['Q_stp_profile']=QstDefault
                else:
                    stpScal['Q_stp_profile']=1.0  
                        
            elif(test_type== 'Q1_stp_profile'):
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'Q1stpMethod':0} #set profile entyr method to 'file'
                comp_handle.set_parameters(**parameters)  
                
                QstDefault=comp_handle.get_parameters()['Q1stpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['Q1_stp_profile']=QstDefault
                else:
                    stpScal['Q1_stp_profile']=1.0
                    
            elif(test_type==  'PF_stp_profile'): 
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'PFstpMethod':0} #set profile entyr method to 'file'
                comp_handle.set_parameters(**parameters)  
                
                PFstDefault=comp_handle.get_parameters()['PFstpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['PF_stp_profile']=PFstDefault
                else:
                    stpScal['PF_stp_profile']=1.0  
                    
            elif(test_type==  'PF1_stp_profile'): 
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'PF1stpMethod':0} #set profile entyr method to 'file'
                comp_handle.set_parameters(**parameters)  
                
                PFstDefault=comp_handle.get_parameters()['PF1stpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['PF1_stp_profile']=PFstDefault
                else:
                    stpScal['PF1_stp_profile']=1.0                 
                    
            elif(test_type== 'P_stp_profile'):
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'PstpMethod':0} #set profile entyr method to 'file'
                comp_handle.set_parameters(**parameters)  
                
                PstDefault=comp_handle.get_parameters()['PstpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['P_stp_profile']=PstDefault
                else:
                    stpScal['P_stp_profile']=1.0              
                    
            elif(test_type== 'P1_stp_profile'):
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'P1stpMethod':0} #set profile entyr method to 'file' --
                comp_handle.set_parameters(**parameters)  
                
                PstDefault=comp_handle.get_parameters()['P1stpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['P1_stp_profile']=PstDefault
                else:
                    stpScal['P1_stp_profile']=1.0    
                    
            elif(test_type== 'Auxiliary_profile'):
                if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['StpBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['StpBlockID']))
                # For stp profile, set the corresponging profile in the 
                #if scaling = relative: use "default" value times paramterBase as scaling factor --> read the value and then use it at basis for setting the profile
                #elif scaling = absolute: use parameterBase as scaling factor
                parameters = {'AUXstpMethod':0} #set profile entyr method to 'file'
                comp_handle.set_parameters(**parameters)  
                
                PstDefault=comp_handle.get_parameters()['AUXstpDefault']
                if(ProfilesDict[test_profile_name]['scaling']=='relative'):
                    stpScal['Auxiliary_profile']=PstDefault
                else:
                    stpScal['Auxiliary_profile']=1.0 
            
            # Faults                
            elif(test_type=='Fault'):
                if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['FaultBlockID']))
                
                if(scenario_params['Ftype']=='3PHG'):fault_type=0
                elif(scenario_params['Ftype']=='2PHG'):fault_type=1
                elif(scenario_params['Ftype']=='1PHG'):fault_type=2
                elif(scenario_params['Ftype']=='L-L'):fault_type=3
                
                parameters = {'status':1,'N_faults':1, 'Time1':scenario_params['Ftime'], 'Duration1':scenario_params['Fduration'], 'Type1':fault_type} #set to block to 'active' and single fault
                #set impedance depending on if it is defined directly or via residual voltage
                if(scenario_params['Fault X_R']==''):
                    scenario_params['Fault X_R']=3 #default to 3 if not defined
                    ScenariosDict[scenario]['Fault X_R']
                if(scenario_params['F_Impedance']!=''): #impedance is defined
                    parameters['Impedance1']=scenario_params['F_Impedance']
                    parameters['X_R1']=scenario_params['Fault X_R']
                elif(scenario_params['Vresidual']!=''):
                    Zfault, X_R_fault = calc_fault_impedance(scenario_params['Vresidual'],scenario_params['Fault X_R'],  SetpointsDict[scenario_params['setpoint ID']]['V_POC'],VbaseTestSrc, SCRMVAbase, SetpointsDict[scenario_params['setpoint ID']]['SCR'], SetpointsDict[scenario_params['setpoint ID']]['X_R'])#
                    parameters['Impedance1']= Zfault  
                    parameters['X_R1']=X_R_fault
                    ScenariosDict[scenario]['F_Impedance']=round(Zfault,2)
                else:
                    parameters['Impedance1']=0.01 #default to 0 Ohm fault
                    parameters['X_R1']=scenario_params['Fault X_R']
                    ScenariosDict[scenario]['F_Impedance']=0.01
        
                comp_handle.set_parameters(**parameters)
                
                simulation_duration = round(scenario_params['Ftime']+scenario_params['Fduration']+5.0)
                # set up fault block
                        #with data from table if type is fault
                        #with data from table if type is Multifault
                        # with data from table AND data from auto-generated Multifault profile if type is Mutlifault_random            
                # set scenario duration to Ftime+Fduration+5s if type=Fault
                # set scenario duration to Ftime+Fduration+5s of last fault in series if Type=Multifault
                # set scenario duration to Ftime+ProfileLength+5s if type=Multifault_random
                
            elif(test_type=='Multifault'):
                if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['FaultBlockID']))
                
                N_faults = len(scenario_params['Ftype'])
                parameters = {'status':1, 'N_faults': N_faults}
                
                for faultID in range(0, len(scenario_params['Ftype'])):
                    if(scenario_params['Ftype'][faultID]=='3PHG'):fault_type=0
                    elif(scenario_params['Ftype'][faultID]=='2PHG'):fault_type=1
                    elif(scenario_params['Ftype'][faultID]=='1PHG'):fault_type=2
                    elif(scenario_params['Ftype'][faultID]=='L-L'):fault_type=3
                    
                    if(scenario_params['Fault X_R'][faultID]==''):
                        scenario_params['Fault X_R'][faultID]=3 #default fault X_R to 3
                    if(scenario_params['F_Impedance'][faultID] !=''): #If impedance is defined, set it directly                
                        parameters['Impedance'+str(faultID+1)]=scenario_params['F_Impedance'][faultID]
                        parameters['X_R'+str(faultID+1)]=scenario_params['Fault X_R'][faultID]
                    elif(scenario_params['Vresidual'][faultID] !=''):
                        Zfault, X_R_fault=calc_fault_impedance(scenario_params['Vresidual'][faultID],scenario_params['Fault X_R'][faultID], SetpointsDict[scenario_params['setpoint ID']]['V_POC'],VbaseTestSrc, SCRMVAbase, SetpointsDict[scenario_params['setpoint ID']]['SCR'], SetpointsDict[scenario_params['setpoint ID']]['X_R'])#
                        parameters['Impedance'+str(faultID+1)]=Zfault
                        parameters['X_R'+str(faultID+1)]=X_R_fault
                        ScenariosDict[scenario]['F_Impedance'][faultID]=round(Zfault, 2)
                    else:
                        parameters['Impedance'+str(faultID+1)]=0.01 #default to 0 Ohm fault
                        ScenariosDict[scenario]['F_Impedance'][fault_ID]=0.01
                    
                    parameters['Time'+str(faultID+1)]=scenario_params['Ftime'][faultID]
                    parameters['Duration'+str(faultID+1)]=scenario_params['Fduration'][faultID]
                    parameters['Type'+str(faultID+1)]=scenario_params['Ftype'][faultID]
                    #parameters['Impedance'+str(faultID)]=scenario_params['F_Impedance'][faultID]
                    #parameters['X_R'+str(faultID)]=scenario_params['Fault X_R'][faultID]
            
                comp_handle.set_parameters(**parameters)   
                simulation_duration = round(scenario_params['Ftime'][faultID]+scenario_params['Fduration'][faultID]+5.0) #faultID points to the last fault in the sequence. Take the values form this fault to determine the length of the simulation
            
            elif(test_type=='Multifault_random'):
                random_times=[0.01, 0.01, 0.2, 0.2, 0.5, 0.5, 0.75, 1, 1.5, 2, 2, 3, 5, 7, 10]
                random.shuffle(random_times)
                random_fault_duration=[0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.22, 0.22, 0.22, 0.22, 0.22, 0.22, 0.43]
                random.shuffle(random_fault_duration)
                Zgrid=math.pow((ProjectDetailsDict['VbaseTestSrc']*1000),2)/setpoint_params['GridMVA']/1000000 #Zgrid
                random_impedances=[0,0,0,0,0,0,0, 3,3,3,3,3, 2,2,2]
                random.shuffle(random_impedances)
                random_fault_types=[2,2,2,2,2,2,1,1,1,1,1,1,1,0,0]
                random.shuffle(random_fault_types)
                
                if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                    canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp(int(PSCADmodelDict['FaultBlockID']))
                parameters = {'status':1, 'N_faults': 15}
                faultTimeOffset=scenario_params['Ftime']
                ScenariosDict[scenario]['F_Impedance']=[]
                ScenariosDict[scenario]['Fault X_R']=[]
                ScenariosDict[scenario]['Ftime']=[]
                ScenariosDict[scenario]['Fduration']=[]
                ScenariosDict[scenario]['Ftype']=[]
                for faultID in range(0, 15):
                    parameters['Impedance'+str(faultID+1)]=random_impedances[faultID]*Zgrid
                    parameters['X_R'+str(faultID+1)]=setpoint_params['X_R']
                    parameters['Time'+str(faultID+1)]=faultTimeOffset
                    parameters['Duration'+str(faultID+1)]=random_fault_duration[faultID]
                    parameters['Type'+str(faultID+1)]=random_fault_types[faultID]                
                    faultTimeOffset=faultTimeOffset+random_fault_duration[faultID]+random_times[faultID]   
                    ScenariosDict[scenario]['F_Impedance'].append(round(random_impedances[faultID]*Zgrid, 2))
                    ScenariosDict[scenario]['Fault X_R'].append(setpoint_params['X_R'])
                    ScenariosDict[scenario]['Ftime'].append(faultTimeOffset)
                    ScenariosDict[scenario]['Fduration'].append(random_fault_duration[faultID])
                    ScenariosDict[scenario]['Ftype'].append(random_fault_types[faultID])
                    
                comp_handle.set_parameters(**parameters)   
                simulation_duration = round(faultTimeOffset+5.0)             
            
            elif(test_type=='TOV'):
                if(is_number(scenario_params['Capacity(F)']) ):
                    capacity=float(scenario_params['Capacity(F)'])
                else:
                    Qinj, capacity = TOV_calc.calc_capacity(setpoint_params['GridMVA'], setpoint_params['X_R'], setpoint_params['P'], setpoint_params['Q'], setpoint_params['V_POC'], ProjectDetailsDict['VbaseTestSrc']*1000, scenario_params['Vresidual']) #Vbase must be provided in volts
                ScenariosDict[scenario]['Capacity(F)']=round(capacity,4)
                capacity_uF=math.pow(10,6)*capacity #scale to uF
                if(PSCADmodelDict['FaultBlockCanvas (optional)']==''):
                        canvas_handle = project.user_canvas("Main")
                else:
                    canvas_handle = project.user_canvas(PSCADmodelDict['FaultBlockCanvas (optional)'])
                comp_handle = canvas_handle.user_cmp((PSCADmodelDict['FaultBlockID']))            
                parameters = {'status':1, 'N_faults': 1, 'Duration1':scenario_params['Fduration'], 'Time1':scenario_params['time'], 'Type1':4, 'Capacity1': capacity_uF}    
                
                comp_handle.set_parameters(**parameters)   
                simulation_duration = round(scenario_params['time']+scenario_params['Fduration']+5.0) #faultID points to the last fault in the sequence. Take the values form this fault to determine the length of the simulation

            
            #-------------------------------------------------------------------------
            #GENERATE TEST PROFILE FILES (either containing the test profile or default (to make the test library components work even if profile not required) )
                # for POC profile or stp profile tests, simply copy the profiles from the excel sheet
            if('Test profile' in scenario_params.keys()): #check if the test type allows for a test profile
                if(test_profile_name !=''):#check if a test profile is actually listed
                    if(test_profile_name in ProfilesDict.keys()): #check if profile has been properly defined
                        simulation_duration = max(profileToFile(test_profile_name, test_type, workspaceFileDir, scenario), simulation_duration) #provide 
                    else: print("ERROR: listed test profile has not been defined")
                else: print("LOG: No test profile defined")
                #adjust file names in the voltageSource block and SetpointProfile block for the different project file copies

        if(PSCADmodelDict['StpBlockCanvas (optional)']==''):
            canvas_handle=project.canvas("Main").component(int(PSCADmodelDict['StpBlockID'])).canvas()
        else:
            canvas_handle=project.canvas(PSCADmodelDict['StpBlockCanvas (optional)']).component(int(PSCADmodelDict['StpBlockID'])).canvas()
        #canvas_handle=canvas_handle.user_canvas('setpointProfiles')
        for profile_file in ['Vstp_profile', 'Qstp_profile','Q1stp_profile', 'PFstp_profile', 'PF1stp_profile', 'Pstp_profile','P1stp_profile', 'AUX_profile']:
            comp_handle = canvas_handle.find_all(Name=str(profile_file))[0]
            parameters={'File':'testProfiles/'+profile_file+scenario+'.txt'}
            comp_handle.set_parameters(**parameters)
            
        if(PSCADmodelDict['GridSourceCanvas (optional)']==''):
            #canvas_handle = project.user_canvas("Main:GridSource")
            canvas_handle=project.canvas("Main").component(int(PSCADmodelDict['GridSourceID'])).canvas()
        else:
            #canvas_handle = project.user_canvas(PSCADmodelDict['GridSourceCanvas (optional)']+':GridSource')
            canvas_handle=project.canvas(PSCADmodelDict['GridSourceCanvas (optional)']).component(int(PSCADmodelDict['GridSourceID'])).canvas()
        #canvas_handle=canvas_handle.user_canvas('GridSource')
        for profile_file in ['Fprofile', 'Vprofile', 'PHprofile']:
            #project_handle=
            comp_handle = canvas_handle.find_all(Name=str(profile_file))[0]
            parameters={'File':'testProfiles/'+profile_file+scenario+'.txt'}
            comp_handle.set_parameters(**parameters)
            
    
        #-------------------------------------------------------------------------
        #SET SIMULATION PARAMETERS (partly depending on scenario definition)
        #simulation_duration =0.5# FOR DEBUGGING ONLY!!
            #ScenarioDuration if defined, otherwise default based on test category and profile
            #make sure "save to file" option is activated
        for project_id in range(1,PSCADmodelDict['Projects']+1):
            project=pscad.project(project_files['Project1'])
            # set simulation duration, set plot time step and activate "save channels to disk"
            
            parameters={'time_duration':simulation_duration, 'sample_step':float(PSCADmodelDict['plotStep_us']), 'PlotType':1}
            project.set_parameters(**parameters)       
        
            #save model as it is
        for project_id in range(1,PSCADmodelDict['Projects']+1):
            project=pscad.project(project_files['Project1'])
            project.save()        
        """
            -run the simulation
            -for every project file in the set:
                    copies the results to the folders containing the project copies and does the conversion to .csv             
            
        """ 
        ws = pscad.workspace()
        sets = ws.list_simulation_sets()
        if not (scenario_group in sets):
            ws.create_simulation_set(scenario_group)#create simulation set
        ss = pscad.simulation_set(scenario_group)#retrieve set handle and write into variable SS
        for key in project_files:
            ss.add_tasks(project_files[key])
            #ss = ws.simulation_set("default") #!!! project specific
    #ws.save()        
    #-------------------------------------------------------------------------           
    #RUN TEST

    #run simulation
    # ws = pscad.workspace()
    # #check if simulation set exists. If not, create one            
    # sets = ws.list_simulation_sets()
    # if len(sets)==0:
    #     ws.create_simulation_set('default')
    #     ss = mhrc.automation.simulation.SimulationSet(pscad, 'default')
    #     ss.add_tasks(FileInfoDict['Project1'])
    #     ss = ws.simulation_set("default") #!!! project specific
    ss.run()
    # else:
    #     pscad.run_all_simulation_sets()
        
    #-------------------------------------------------------------------------
    #SAVE RESULTS AS CSV IN "RESULTS" FOLDER --> for every project out of the scenario group
    #Assumption: only the main project has relevant output files.
    #take .inf file to read channels
    #take the other files with same name as project file and read+append
    #if last batch is run
    #pscad.release_license()#release the license after each batch, to make sure it is free in the end
    pscad.quit()
    
    for scenario in scenario_groups[scenario_group]:      
        for key in PSCADmodelDict:
            if('Project' in key and key!='Projects'):
                project_files[key]=PSCADmodelDict[key]+'_'+scenario # for the case where multiple instances of the model are in one workspace (simulation batch) the project file names need to be mapped to the changed changed names

        project_name =project_files['Project1']
        outfile_name=project_files_orig['Project1']
        file_names=os.listdir(workspaceFileDir+"\\"+project_name+compiler_acronym)
        result_files=[]
        for file_name in file_names:
            if( (outfile_name in file_name) and (file_name.endswith('.inf') or file_name.endswith('.out')) ):
                result_files.append(file_name)
        
        #create result directory
        results_folder=ResultsDir+"\\"+testRun+"\\"+scenario
        try:
            os.mkdir(ResultsDir+'\\'+testRun)
        except:
            print('creation of dir unsuccessful')

        try:
            os.mkdir(results_folder)
        except OSError:
            print("Creation of the directory %s failed" % results_folder)
        else: 
            
            print("Successfully created the directory %s " % results_folder)
        #read signal labels from .inf file    
        headers=open(workspaceFileDir+"\\"+project_name+compiler_acronym+"\\"+result_files[0], 'r')
        signal_names = ['time(s)']
        for line in headers:
            startpos = line.find("Desc=")+6
            endpos = line.find("Group")-3
            signal_names.append(line[startpos:endpos])
        headers.close()
        
        results=pd.read_csv(workspaceFileDir+"\\"+project_name+compiler_acronym+"\\"+result_files[1], delim_whitespace = True, header=None)
        col_offset=len(results.columns)
        #read all files ending with .out into pandas data frame adn appred in in chroniological order to combined pandas dataframe
        for result_id in range (2, len(result_files)):
            temp_df=pd.read_csv(workspaceFileDir+"\\"+project_name+compiler_acronym+"\\"+result_files[result_id], delim_whitespace = True, header=None)
            temp_df=temp_df.drop(columns=0) #deletes column containing "time"
            results=pd.concat([results, temp_df], axis=1)
        #add the signal labels as column names to the combined frame.
        results.columns=signal_names
        results.to_csv(results_folder+"\\"+str(scenario)+"_results.csv",index=False)
    
    
        #write configuration dictionaries to .dir file for analysis script or just future reference
        scenario_params=ScenariosDict[scenario]
        stpID=scenario_params['setpoint ID']
        setpoint_params = SetpointsDict[stpID]
        
        testInfoDir=results_folder+"\\testInfo"
        try:
            os.mkdir(testInfoDir)
        except OSError:
            print("Creation of the directory %s failed" % testInfoDir)
        else:
            print("Successfully created the directory %s " % testInfoDir)
        testInfo=shelve.open(testInfoDir+"\\"+str(scenario), protocol=2)
        testInfo['scenario']=scenario
        testInfo['scenario_params']=scenario_params
        testInfo['setpoint']=setpoint_params
        if('Test profile' in scenario_params.keys()):
            if( (scenario_params['Test profile']!= None) and (scenario_params['Test profile']!='') ):
                testInfo['profile']=ProfilesDict[scenario_params['Test profile']]
        testInfo.close()  
    
    #delete scenario_group folder content
    scenario_sets_path=main_folder_out+"\\model_copies_sets\\"+testRun
    shutil.rmtree(scenario_sets_path, ignore_errors=True)    
    
    return
    
    # for result_id in range (1, len(result_files)):
    #     out_file=open(workspace_folder+"\\"+project_name+compiler_acronym+"\\"+result_files[result_id], 'r')
    #     row_id=-1 #start at -1 because first row is always empty
    #     for line in out_file:
    #         id=0
    #         while id+1<len(line):
    #             while line[id]==' ' :
    #                 id+=1
    #             startpos = id
    #             while line[id]!=' ' and line[id]!='\n':
    #                 id+=1
    #             endpos = id  
    #             if(endpos>startpos):
    #                 if( len(data_rows)>row_id):
    #                     data_rows[row_id].append(float(line[startpos:endpos]))
    #                 else:
    #                     data_rows.append([float(line[startpos:endpos])]) #add new row. This is only for the first column
    #         row_id+=1
            
    # for row in data_rows:
    #     file_entry=''
    #     for column_id in range(0, len(row)-1):
    #         file_entry+=str(row[column_id]+",")
    #     file_entry+=str(row[column_id])
    #     results.write(file_entry+"\n")
  
def calc_fault_impedance(Vresidual, Fault_X_R, Vpoc, Vbase, MVAbase, grid_SCR, grid_X_R):
    if(Vresidual<0.0001):
        Vresidual=0.0001
    # calc grid resistance and reactance in Ohms
    # derive equation for voltage at POC as a function of the given parameters including fautl X_R
        #--> solve for variable Fault Impedance in Ohms. 
    Zbase=(Vbase*Vbase*1000**2)/(MVAbase*1000000) # Zbase assuming MVA base is provided in MVA and Vbase in kV
    Zgrid_abs=1.0/grid_SCR*Zbase
    
    # Zfault=Zgrid_abs/((Vpoc/Vresidual)-1) #calcualte Zfault based on voltage divider
    ANG_Zgrid=math.atan(grid_X_R)
    ANG_Zfault=math.atan(Fault_X_R)
    phi=ANG_Zgrid-ANG_Zfault
    Zfault=Zgrid_abs/(-1*math.cos(phi)+math.sqrt(math.pow(math.cos(phi),2)+(1/(math.pow((Vresidual/Vpoc),2)))-1))    
    X_R_fault=Fault_X_R    
    
    return Zfault, X_R_fault
    
    
    
    #return fault impedance in Ohms              
    
def createModelCopies(scenario_group): #scenario_group is only the key for the scenario_groups dictionary
    #print("creating model copy")
    global scenario_groups
    #print(scenario_groups)
    global testRun
    global ModelCopyDir
    #print('ModelCopyDir: '+ModelCopyDir)
    #global activeScenarios
    #print(activeScenarios)
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
    
    #targetDir = ResultsDir+"\\"+testRun+"\\"+scenario_group
    set_directory= main_folder_out+"\\model_copies_sets\\"+testRun+"\\"+scenario_group
    
    
    #Create model copy used for the set
    try:
        shutil.copytree(base_model, set_directory)
    except OSError:
        print("Creation of the directory %s failed" % set_directory)
    else:
        print("Successfully created the directory %s" % set_directory)
        
    # # 07/04/2022: Adapt to Ingeteam model: copy the DLL and LIB files to the simulation folder --> NOT REQUIRED. ROUTINE DOES THIS as long as file path is configured.
    # key = 'Project1'
    # for scenario in scenario_groups[scenario_group]:
    #     project_name = PSCADmodelDict[key]+'_'+scenario
    #     fromdir = set_directory+"\\IJH"
    #     todir = set_directory+"\\IJX\\"+project_name+compiler_acronym        
    #     try:
    #         os.mkdir(todir)
    #     except OSError:
    #         print("Creation of the directory %s failed" % todir)
    #     else:
    #         print("Successfully created the directory %s" % todir)            
    #     src_files = os.listdir(fromdir)
    #     for file_name in src_files:
    #         full_file_name = os.path.join(fromdir, file_name)
    #         if os.path.isfile(full_file_name):
    #             shutil.copy(full_file_name, todir)        
        
               
    #check if example case is to be created or not
    if("create example cases" in PSCADmodelDict.keys()):
        if(PSCADmodelDict['create example cases']=='yes'):
            for scenario in scenario_groups[scenario_group]:
                model_copy_dir = main_folder_out+"\\model_copies\\"+testRun+"\\"+scenario
                try:
                    shutil.copytree(base_model, model_copy_dir)
                except OSError:
                   print("Creation of the directory %s failed" % model_copy_dir)
                else:
                   print("Successfully created the directory %s" % model_copy_dir)

    return set_directory, testRun

def profileToFile(profile_name, TestType, folder_path, name_ext=''):
    profile = copy.deepcopy(ProfilesDict[profile_name])
    #set file name to match naming expectations of the pre-configured test tool in PSCAD
    if (TestType == 'F_profile'): filename = 'Fprofile'+name_ext+'.txt' # profile scaling should always be set to Hz when running automated tests (Hz)
    elif (TestType == 'V_profile'): filename ='Vprofile'+name_ext+'.txt' # profile scaling should either be based on Vbase defined in context menu of block (absolute) or based on Vpu (relative)
    elif (TestType == 'ANG_profile'): 
        filename ='PHprofile'+name_ext+'.txt' # No profile scaling 
        #offset y-column by 1
        for index in range(0, len(profile['y_data'])-1):
            profile['y_data'][len(profile['y_data'])-1-index]=profile['y_data'][len(profile['y_data'])-2-index]
        profile['y_data'][0]=0
    elif (TestType == 'V_stp_profile'): filename ='Vstp_profile'+name_ext+'.txt' #
    elif (TestType == 'Q_stp_profile'): filename ='Qstp_profile'+name_ext+'.txt'
    elif (TestType == 'Q1_stp_profile'): filename ='Q1stp_profile'+name_ext+'.txt'
    elif (TestType == 'P_stp_profile'): filename ='Pstp_profile'+name_ext+'.txt'
    elif (TestType == 'P1_stp_profile'): filename ='P1stp_profile'+name_ext+'.txt'
    elif (TestType == 'PF_stp_profile'): filename ='PFstp_profile'+name_ext+'.txt'
    elif (TestType == 'PF1_stp_profile'): filename ='PF1stp_profile'+name_ext+'.txt'
    elif (TestType == 'Auxiliary_profile'): filename = 'AUX_profile'+name_ext+'.txt'
    f=open(folder_path+"\\testProfiles\\"+filename, 'w+')
    for line_id in range (0, len(profile['x_data'])):
        #if PSCAD offset defined, apply that. If scaling defined, apply that as well.
        if(is_number(profile['offset_PSCAD']) ):
            offset=float(profile['offset_PSCAD'])
        else:
            offset=0.0
        if(is_number(profile['scaling_factor_PSCAD'])):
            scaling=float(profile['scaling_factor_PSCAD'])
        else:
            scaling=1.0
        
        if(TestType in stpScal.keys()):
            y=float(profile['y_data'][line_id]+offset)*float(stpScal[TestType])*scaling #Apply any additional scaling is profile model is "relative" and it is a setpoint profile)
        else:
            y=float(profile['y_data'][line_id]+offset)*scaling #
            

            
        x=float(profile['x_data'][line_id])
        
        f.write(str(x)+"\t"+str(y)+"\n") #write values to text file, then newline
    simulation_duration = float(x) #use last defined x-value of the profile as the duration of the simulation
    f.write("ENDFILE:")
    f.close()   
    return simulation_duration

def createDefaultProfiles(workspaceFileDir, scenario):
    for profileName in ['Fprofile', 'Vprofile', 'PHprofile', 'Vstp_profile', 'Qstp_profile','Q1stp_profile', 'Pstp_profile', 'P1stp_profile', 'PFstp_profile', 'PF1stp_profile','Aux_profile']:
        f=open(workspaceFileDir+"\\testProfiles\\"+profileName+scenario+'.txt', 'w+')
        default_profile=[[1,1],[2,1],[3,1],[4,1],[5,1]]
        for data_point_id in range(0, len(default_profile)):
            data_point=default_profile[data_point_id]
            f.write(str(data_point[0])+"\t"+str(data_point[1])+"\n")
        f.write("ENDFILE:")
        f.close()   
    

def group_scenarios(activeScenarios, scenariosPerGroup): #group active scenarios in batches of length=scenariosPerGroup
    global scenario_groups
    activeScenario_id=0
    group_cnt=1
    while activeScenario_id<len(activeScenarios):
        scenario_group=[]
        scenario_cnt=0        
        while ( (activeScenario_id<len(activeScenarios)) and (scenario_cnt<scenariosPerGroup)):
            scenario_group.append(activeScenarios[activeScenario_id])
            scenario_cnt+=1
            activeScenario_id+=1        
        scenario_groups['group_'+str(group_cnt)]=scenario_group
        group_cnt+=1
    return scenario_groups #returns a pointer towards the global variable 'scenario_groups'
    
def main():
    i=1
    global workspace_folder_location
    global workspace_folder_path
    global max_processes
    #scenarios = ScenariosDict.keys()
    global activeScenarios
    #for scenario in scenarios:
        #if(ScenariosDict[scenario]['run in PSCAD?']=='yes'):

            #activeScenarios.append(scenario)# add keys of scenarios listed as "yes" under "run in PSCAD?" to the activeScenarios list
    
    #CHANGE OF STRATEGY: group active scenarios in batches of n scenarios (list of lists or dict of dicts, or list of dicts or whatever)   
    global scenariosPerGroup
    global scenario_groups
    #scenario_groups = group_scenarios(activeScenarios, scenariosPerGroup) #
            
    startTime = datetime.datetime.now()
    print("\nStart Time: " + startTime.strftime("%d/%m/%Y - %H:/%M:/S %p") + "\n")
    
    #for debugging uncomment the next three lines, to run sequential processes in stead of parallel
    for scenario_group in scenario_groups:
        current_set_folder, testRun_ = createModelCopies(scenario_group)        
        print("current_set_folder: "+current_set_folder)
        runTest(scenario_group, current_set_folder, testRun_)
    
    l = Semaphore(1)
    sem = Semaphore(max_processes)
    
    ##CHANGE OF STRATEGY: map scenario groups to workers, insdtead of individual scenarios
    #p = Pool(processes = max_processes, initializer = initializer, initargs = (l, sem))
    ##i=activeScenarios
    ##i=scenario_groups
    ##print(i)
    #p.map(worker, scenario_groups) #this will parse only the dict keys to the worker processes
    
    finishTime = datetime.datetime.now()
    print("\nFinish Time: " +finishTime.strftime("%d/%m/%Y - %H:%M:%S %p") + "\n")
    
    return 0

if __name__ == '__main__':
    main()