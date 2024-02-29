# -*- coding: utf-8 -*-
"""
Created on Tue Feb 15 09:08:16 2022

@author: ESCO

COMMENTS:
    Script is aiming to test the PSSE network focusing on Fault studies and control analysis
    1. Fault simulation:
        + can apply to various elements including lines, Tx 2w, Tx 3w or bus.
        + Only line fault can be considered at differnt types: 3phG, 2phG, 1phG. Other element fault is only with 3phG
        + Muliti fault studies is aimed at 3phG faults on a number of lines
        + Current methodology for line fault: 
            insert a fault bus between ibus and jbus with a distance from ibus defined in "location" collumn
            apply fault to this bus, 
            deactivate/switch off/on the line to represent the CB operation.
    2. Switching analysis:
        + can switch on/off (defined in event_type) all type of equipment listed on Event_Element
        + apply to line, Tx 2w, Tx 3w, bus, machine, shunt
    3. voltage control:
        + work similar to smib test
        + test profile to be defined in Test Profile collumn
        + type of variable change (absolute or relative) to be defined in Event_type collumn
    

    
"""


from __future__ import with_statement
from contextlib import contextmanager

import os, sys
import csv
import ast
import pdb
import pandas as pd
import shutil
import shelve

from win32com.client import Dispatch
import datetime
import time
#timestr = time.strftime("%Y%m%d")
timestr = str(datetime.datetime.now().strftime("%Y%m%d-%H%M"))

#import readtestinfo

from openpyxl import load_workbook
from openpyxl import Workbook
#from openpyxl.chart import (ScatterChart,Reference,Series)
#from openpyxl.chart.text import RichText
#import openpyxl
#import openpyxl.drawing as drawing

#more openpyxl stuff but might not be required.
###############################################################################
#IMPORT PYTHON
###############################################################################
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

import psse34
import pssarrays
import redirect
import bsntools
import psspy

import usrout
import rav
import pssras
import pssppe
import psseloc
import excelpy
import lntpy
import gicdata
import pssexcel
import pssplot

redirect.psse2py()
with silence():
    psspy.psseinit(80000)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()
standalone_script=""

import dyntools

###############################################################################
# Auxiliary Functions
###############################################################################
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

def createPath(main_path_out):
    path = os.path.normpath(main_path_out)
    path_splits = path.split(os.sep) # Get the components of the path
    child_folder = r"C:" # Build up the output path from C: directory
    for i in range(len(path_splits)-1):
        child_folder = child_folder + "\\" + path_splits[i+1]
        make_dir(child_folder)
    return child_folder

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False
    
plot_channels={}

###############################################################################
# USER CONFIGURABLE PARAMETERS
###############################################################################
TestDefinitionSheet = r'20230828_SUM_TESTINFO_V1.xlsx'
simulation_batches=['S52513_NW']

try:
    testRun = timestr + '_' + simulation_batches[0] #define a test name for the batch or configuration that is being tested -> link to time stamp for auto update
except:
    testRun = timestr
    

###############################################################################
# Define Project Paths
###############################################################################
    
#main_folder=r"C:\Users\Mervin Kall\Documents\GitHub\PowerSystemStudyTool\20220203_APE\PSSE_sim" #put relative file path here rather than absolutepath (given script expected to always be in same location)
script_dir=os.getcwd()
main_folder=os.path.abspath(os.path.join(script_dir, os.pardir))

# Create directory for storing the results
if "ESCO Pacific\ESCO - Projects" in main_folder: # if the current folder is online (under ESCO - Projects), create a new directory to store the result
    main_path_out = main_folder.replace("ESCO Pacific\ESCO - Projects","Documents\Projects") # Change the path from Onedrive to Local in Documents
    main_folder_out = createPath(main_path_out)
else: # if the main folder is not in Onedrive, then store the results in the same location with the model
    main_folder_out = main_folder
# main_folder_out = r"C:\1. Power System Studies\20220318_LSF\PSSE_sim\scripts" # Option to define the absolute path of the result location
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
base_model_workspace = main_folder+"\\base_model" #path of the workspace folder, formerly "workspace_folder" --> in case the workspace is located in a subdirectory of the model folder (as is the case with MUL model for example)
zingen=main_folder+"\\zingen\\dsusr_zingen.dll"
libpath = os.path.abspath(main_folder) + "\\scripts\\Libs"
sys.path.append(libpath)
# print ("libpath = " + libpath)

# Directory to store Steady State/Dynamic result
ResultsDir = OutputDir+"\\dynamic_network"
make_dir(ResultsDir)

###############################################################################
#GLOBAL VARIABLES
###############################################################################
import auxiliary_functions as af
import readtestinfo as readtestinfo

cb_bus_start_index = 99998
cb_bus = 99998 #99999 was the EXT_GRID
reset_index = 1

event_queue=[]
var_init_dict={}
breakers=[]
active_faults={} #variable keeps track which faults are active in PSS/E and maps the ID's that are assignet to the faults in this script to the IDs that PSS/E assigns to the faults

#return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'SimulationSettings', 'ModelDetailsPSSE', 'Setpoints', 'ScenariosSMIB', 'Profiles', 'NetworkFaults','MonitorBuses', 'MonitorBranches'])
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'SimulationSettings', 'ModelDetailsPSSE', 'Setpoints', 'ScenariosSMIB', 'Profiles', 'NetworkScenarios'])
ProjectDetailsDict = return_dict['ProjectDetails']
SimulationSettingsDict = return_dict['SimulationSettings']
PSSEmodelDict = return_dict['ModelDetailsPSSE']
SetpointsDict = return_dict['Setpoints']
ScenariosDict = return_dict['ScenariosSMIB']
ProfilesDict = return_dict['Profiles']
NetworkScenDict = return_dict['NetworkScenarios']
#MonitorBuses = return_dict['MonitorBuses']
#MonitorBranches = return_dict['MonitorBranches']

###############################################################################
#LIBRARIES IMPORT
###############################################################################

###############################################################################
# FUNCTIONS
###############################################################################

def save_test_description(testInfoDir, scenario, scenario_params):
    try:
        os.mkdir(testInfoDir)
    except OSError:
        print("Creation of the directory %s failed" % testInfoDir)
    else:
        print("Successfully created the directory %s " % testInfoDir)
    testInfo=shelve.open(testInfoDir+"\\"+str(scenario), protocol=2)
    testInfo['scenario']=scenario
    testInfo['scenario_params']=scenario_params
    testInfo.close() 
    pass

def init_val(param_type, ID): #function returns the value of a variable at the time the function was first called for that variable. It is used for applying setpoint profiles relative to the initial value (e.g. +/-5%): The offset needs to refer to the initial value of that variable
    if(param_type=='VAR'):
        if(ID in var_init_dict.keys()):
            return var_init_dict[ID]
        else:
            var_init_dict[ID]=psspy.dsrval('VAR', ID)[1]
            return var_init_dict[ID]
    elif(param_type=='CHN'):
        if(ID in chn_init_dict.keys()):
            return chn_init_dict[ID]
        else:
            chn_init_dict[ID]=psspy.chnval(ID)[1]
            return chn_init_dict[ID]


def interpolate(profile, TimeStep, density = 20.0, scaling=1.0, offset=0.0):
    # in PSS/E profiles applied via variable changes can ony occur in steps, however in the excel spreadsheet profiles are defined vial multiple points. 
    # interpolate as follows: for every signal slope, itentify minimum and maximum y value in profile, interpolate, this resolution as 20 interpolate between minimum and maximum in 40 steps, however, limit step size to variable change every 10ms.
    x_data=profile['x_data']
    y_data=profile['y_data']
    y_max=max(y_data)
    y_min=min(y_data)
    y_delta_max=y_max-y_min

    x_interpol=[]
    y_interpol=[] 
    if(TimeStep<0.01):
        minStep=0.01
    else:
        minStep=TimeStep
    
    for point_id in range (0, len(x_data)-1):
        X0=x_data[point_id]
        X1=x_data[point_id+1]
        
        Y0=y_data[point_id]
        Y1=y_data[point_id+1]
        
        x_interpol.append(X0)# append point from original profile
        y_interpol.append(scaling*(Y0+offset))# append point from original profile
        y_delta=Y1-Y0
        x_delta=X1-X0
        if(abs(y_delta)>y_delta_max/density) and x_delta > minStep: #if y_change is more than 1/40th of the maximum change do interpolation
            interpol_step=x_delta/(abs(y_delta)/y_delta_max*density) #number of interpolation steps in proportion to y_delta compared to maximum y_delta
            if(interpol_step<minStep):
               interpol_step=minStep
            x=X0+interpol_step
            while x<X1:
                y=Y0+(y_delta/x_delta)*(x-X0)
                x_interpol.append(x)
                y_interpol.append(scaling*(y-offset))
                x+=interpol_step
    
    point_id+=1         
    x_interpol.append(x_data[point_id])# append point from original profile
    y_interpol.append(scaling*(y_data[point_id]-offset) )# append point from original profile 
    pass
    new_profile={'scaling':profile['scaling'], 'x_data':x_interpol, 'y_data':y_interpol}
    return new_profile

#create output folder for results and debugging
def createOutputDir():
    result_path = r"""Results"""
    try:
        os.mkdir(result_path)
    except OSError:
        print("Creation of directory %s failed" %result_path)
    else:
        print("Successfully created the directory %s " % result_path)
        
        
def add_breaker(frombus, tobus, line_id, caption, breakers):
    for i in range(0, len(breakers)):
        if( (breakers[i][0] == frombus) and (breakers[i][1] == tobus) and (breakers[i][2]==line_id ) ): #breaker already exists, 
            print('breaker on line from ' +str(frombus)+' to '+str(tobus)+' already exists. data of existing breaker will be returned' )
            return i #(returns id of breaker in breakers_list)    
    print('inserting breaker from '+str(frombus)+' to '+str(tobus))
    global cb_bus
    global reset_index
    if(reset_index==1):
        reset_index = 0
        cb_bus = cb_bus_start_index
    ierr, Vbase = psspy.busdat(frombus, 'BASE') #read base voltage on line top which breaker is to be connected
    psspy.splt(frombus, cb_bus, 'breaker'+str(cb_bus),Vbase) #insert additional bus in between breaker and existing line
    psspy.movebrn(tobus, frombus, str(line_id), cb_bus, str(line_id)) #connect existing line to breaker bus instead
    psspy.purgbrn(frombus, cb_bus, r"""1""") #delete segment between frombus and cb_bus
    psspy.system_swd_data(frombus, cb_bus,r"""1""",[_i,_i,frombus,2],_f,[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],"")
    cb_bus -=1
    print('breaker successfully added')
    breakers.append([frombus, tobus, line_id, cb_bus+1])
    return len(breakers)-1 #returns id os breaker in breakers list
    
def checkFaultListConsistency(fault_list):
    for i in range(0,len(fault_list)):
        current_fault = fault_list[i]
        previous_location = []
        previous_location2=[]
        for j in range(0,i):
            previous_fault = fault_list[j] 
            #if any previous fault in same location included in fault_list for which no AR is derfined 
            #if the two faults are on the same line between the same bus
            if('flocation' in current_fault.keys()):
                current_location = current_fault['flocation']
                current_location2 = current_fault['distBus']
                current_ID = current_fault['lineID']
            elif('fromBus' in current_fault.keys()):
                current_location = current_fault['fromBus']
                current_location2 = current_fault['toBus']
                current_ID = current_fault['Id']
            if('flocation' in previous_fault.keys()):
                previous_location = previous_fault['fLocation']
                previous_location2 = previous_fault['distBus']
                previous_ID = previous_fault['lineID']
            elif('fromBus' in current_fault.keys()):
                previous_location = previous_fault['fromBus']
                previous_location2= previous_fault['toBus']
                previous_ID=previous_fault['Id']
            #detect if fuatl is included twice
            if( ( (  (current_location==previous_location) and (current_location2==previous_location2)) or ((current_location == previous_location2) and (current_location2 == previous_location)) ) and (current_ID == previous_ID) ):
                #if second fault happens before first fault is cleared
                #other condition
                #-->: exclude fault from list or prompt error
                print 0
    
    return fault_list


#add event to event queue. The event queeu will be used to executing the faults) Possibel events are:
    #apply fault, (time, event_type, [location_ID, itemID], status, impedance)
    #clear fault (time, event_type, [Location_ID, itemID], status, impedance) --> status and impedance would not be required, but this keeps the formatting the same as for applying the fault.
    # open and close breaker (time, event_type, [location_ID, item_id], switch_status) --> component Id will be provided as the position in the breaker array --> actuyal bus numbers for calling PSS/E commands will need to be looked up from that table
    # switch line on (time, event_type, [location_ID, lineID], switch_status )
    #switch 2WD transformer on/off (time, [bus1, bus2, lineID], switch_status)
    # switch 3WD transformer on/off (time, [bus1, bus2, bus3, lineID], switch_status)
    

def add_event(time, event_type, IDs, switch_status, aux_parameter=[0.0 ,0.0]):
    event_queue.append({'time':time, 'event_type':event_type, 'IDs':IDs, 'switch_status':switch_status, 'aux_parameter': aux_parameter})
    
def execute_event(event_queue):
    global standalone_script
    if(event_queue!=[]):
        #run simulation until next event shoudl occur
#        psspy.run(0, event_queue[0]['time'],100,1,100)
        psspy.run(0, event_queue[0]['time'], 500,1,0)
        standalone_script+="psspy.run(0,"+str(event_queue[0]['time'])+", 500, 1, 0)\n"
        
        #execute event in position 0 of queue
        event=event_queue[0]

        #####################################################################################
        # Fault
        #3PHG branch (Tx 2windings) fault -> apply a fault at the IBUS
        if(str(event['event_type'])=='FLT_TX_2W_3PHG'): #in case event is a 3PHG fault of a two winding Tx
            impedance = event['aux_parameter']
            if(impedance==[]):
#                impedance = [0.5, 5.0]
                impedance = [0.0, 0.0]
            if(event['switch_status']==1):
                psspy.dist_branch_fault(event['IDs'][0], event['IDs'][1], event['IDs'][2], 3, 0.0, [impedance[0],impedance[1]]) #apply a fault at the IBUS
                standalone_script+="psspy.dist_branch_fault("+str(event['IDs'][0])+", "+str(event['IDs'][1])+", "+str(event['IDs'][2])+", 3, 0.0, "+str(impedance[0])+","+str(impedance[1])+"])\n"
                add_to_active_faults(event['IDs'][1])
            elif(event['switch_status']==0):
                ierr=psspy.dist_clear_fault(get_psseFaultID(event['IDs'][1])) # clear fault
                standalone_script+="psspy.dist_clear_fault("+str(get_psseFaultID(event['IDs'][1]))+")\n"
                print('breaker ID= '+str(get_psseFaultID(event['IDs'][1])))
                #retrieve ID that fault has in PSSE
                remove_from_active_faults(event['IDs'][1])

        #3PHG branch (Tx 3windings) fault -> apply a fault at the IBUS
        if(str(event['event_type'])=='FLT_TX_3W_3PHG'): #in case event is a 3PHG fault of a three winding Tx
            impedance = event['aux_parameter']
            if(impedance==[]):
#                impedance = [0.5, 5.0]
                impedance = [0.0, 0.0]
            if(event['switch_status']==1):
                psspy.dist_3wind_fault(event['IDs'][0], event['IDs'][1], event['IDs'][2], event['IDs'][3], 3, 0.0, [impedance[0],impedance[1]]) #apply a fault at the IBUS
                standalone_script+="psspy.dist_branch_fault("+str(event['IDs'][0])+", "+str(event['IDs'][1])+", "+str(event['IDs'][2])+", "+str(event['IDs'][3])+", 3, 0.0, "+str(impedance[0])+","+str(impedance[1])+"])\n"
                add_to_active_faults(event['IDs'][1])
            elif(event['switch_status']==0):
                ierr=psspy.dist_clear_fault(get_psseFaultID(event['IDs'][1])) # clear fault
                standalone_script+="psspy.dist_clear_fault("+str(get_psseFaultID(event['IDs'][1]))+")\n"
                print('breaker ID= '+str(get_psseFaultID(event['IDs'][1])))
                #retrieve ID that fault has in PSSE
                remove_from_active_faults(event['IDs'][1])

        #3PHG Bus fault
        if (str(event['event_type'])=='FLT_BUS_3PHG'):
            impedance = event['aux_parameter']
            if(impedance == []): #assign default impedance if no value is given
#                impedance = [0.5,5.0]
                impedance = [0.0, 0.0]
            if(event['switch_status']==1):
                psspy.dist_bus_fault(event['IDs'][0],3,0.0,[impedance[0],impedance[1]])
                standalone_script+="psspy.dist_bus_fault("+str(event['IDs'][0])+",3,0.0,["+str(impedance[0])+","+str(impedance[1])+"])\n"
                add_to_active_faults(event['IDs'][1])
            elif(event['switch_status']==0):
                ierr=psspy.dist_clear_fault(get_psseFaultID(event['IDs'][1])) #retrieve psse fault ID and clear the fault in question
                standalone_script+="psspy.dist_clear_fault("+str(get_psseFaultID(event['IDs'][1]))+")\n"
                print('breaker ID= '+str(get_psseFaultID(event['IDs'][1])))
                remove_from_active_faults(event['IDs'][1])

        #2PHG Bus fault
        if (str(event['event_type'])=='FLT_BUS_2PHG'):
            impedance = event['aux_parameter']
            if(impedance == []): #assign default impedance if no value is given
#                impedance = [0.5,5.0]
                impedance = [0.0, 0.0]
            if(event['switch_status']==1):
                psspy.dist_bus_fault_2(3,0.0,[_i,2, event['IDs'][0],_i],[impedance[0],impedance[1],_f,_f,_f,_f])
                standalone_script+="psspy.dist_bus_fault_2(3,0.0,[_i,2,"+str(event['IDs'][0])+",_i],["+str(impedance[0])+","+str(impedance[1])+",_f,_f,_f,_f])\n"
                add_to_active_faults(event['IDs'][1])
            elif(event['switch_status']==0):
                ierr=psspy.dist_clear_fault(get_psseFaultID(event['IDs'][1])) #retrieve psse fault ID and clear the fault in question
                standalone_script+="psspy.dist_clear_fault("+str(get_psseFaultID(event['IDs'][1]))+")\n"
                print('breaker ID= '+str(get_psseFaultID(event['IDs'][1])))
                remove_from_active_faults(event['IDs'][1])
                
        #1PHG Bus fault
        if (str(event['event_type'])=='FLT_BUS_1PHG'):
            impedance = event['aux_parameter']
            if(impedance == []): #assign default impedance if no value is given
#                impedance = [0.5,5.0]
                impedance = [0.0, 0.0]
            if(event['switch_status']==1):
                psspy.dist_bus_fault_2(3,0.0,[_i,_i, event['IDs'][0],_i],[impedance[0],impedance[1],_f,_f,_f,_f])
                standalone_script+="psspy.dist_bus_fault_2(3,0.0,[_i,_i,"+str(event['IDs'][0])+",_i],["+str(impedance[0])+","+str(impedance[1])+",_f,_f,_f,_f])\n"
                add_to_active_faults(event['IDs'][1])
            elif(event['switch_status']==0):
                ierr=psspy.dist_clear_fault(get_psseFaultID(event['IDs'][1])) #retrieve psse fault ID and clear the fault in question
                standalone_script+="psspy.dist_clear_fault("+str(get_psseFaultID(event['IDs'][1]))+")\n"
                print('breaker ID= '+str(get_psseFaultID(event['IDs'][1])))
                remove_from_active_faults(event['IDs'][1])

        #####################################################################################
        # Switching
        #Switch line on/off
        if (str(event['event_type'])=='SWT_LINE'): #if case is line being switched on or off
            psspy.branch_chng_3(event['IDs'][0],event['IDs'][1], str(event['IDs'][2]),[event['switch_status'],_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
            standalone_script+="psspy.branch_chng_3("+str(event['IDs'][0])+','+str(event['IDs'][1])+',"'+str(event['IDs'][2])+'",['+str(event['switch_status'])+',_i,_i,_i,_i,_i],[_f, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)\n'
 
        #Switch Tx_2w on/off
        if (str(event['event_type'])=='SWT_TX_2W'): #if case is Tx 2w being switched on or off
            psspy.two_winding_chng_6(event['IDs'][0],event['IDs'][1], str(event['IDs'][2]),[event['switch_status'],_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
            standalone_script+="psspy.branch_chng_3("+str(event['IDs'][0])+','+str(event['IDs'][1])+',"'+str(event['IDs'][2])+'",['+str(event['switch_status'])+',_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)\n'
  
        #Switch Tx_3w on/off
        if (str(event['event_type'])=='SWT_TX_3W'): #if case is Tx 3w being switched on or off
            psspy.three_wnd_imped_chng_4(event['IDs'][0],event['IDs'][1],event['IDs'][2], str(event['IDs'][3]),[_i,_i,_i,_i,_i,_i,_i,event['switch_status'],_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
            standalone_script+="psspy.branch_chng_3("+str(event['IDs'][0])+','+str(event['IDs'][1])+','+str(event['IDs'][2])+',"'+str(event['IDs'][3])+'",[_i,_i,_i,_i,_i,_i,_i,'+str(event['switch_status'])+',_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)\n'
  
        #Switch a bus on/off
        if (str(event['event_type'])=='SWT_BUS'): #if case is a bus being switched on or off
            if(event['switch_status']==1): 
                psspy.recn(event['IDs'][0])
                standalone_script+="psspy.recn("+str(event['IDs'][0])+"])\n"
            elif(event['switch_status']==0):
                psspy.dscn(event['IDs'][0])
                standalone_script+="psspy.dscn("+str(event['IDs'][0])+"])\n"

        #Switch a machine on/off
        if (str(event['event_type'])=='SWT_MAC'): #if case is a machine being switched on or off
            psspy.machine_chng_2(event['IDs'][0],event['IDs'][1],[event['switch_status'],_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            standalone_script+="psspy.machine_chng_2("+str(event['IDs'][0])+',"'+str(event['IDs'][1])+'",['+str(event['switch_status'])+',_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n'
 
        #Switch a shunt on/off
        if (str(event['event_type'])=='SWT_SHT'): #if case is a shunt being switched on or off
            psspy.shunt_chng(event['IDs'][0],event['IDs'][1],event['switch_status'],[_f,_f])
            standalone_script+="psspy.shunt_chng("+str(event['IDs'][0])+',"'+str(event['IDs'][1])+'",'+str(event['switch_status'])+',[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n'
 
        #####################################################################################
        # Setpoint change
        # Update the setpoint
        elif(event['event_type']=='var_change_abs'): #change variable during runtime (e.g. for changing setpoint(s))        
            if(event['model_type']=='WAUX' or event['model_type']=='WGEN'):
#                L=psspy.windmind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
                psspy.change_wnmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value']))
        
        elif(event['event_type']=='var_change_rel'): #change variable during runtime (e.g. for changing setpoint(s)) 
            if(event['model_type']=='OTHER'):
                L = psspy.cctmind_buso(event['bus'],event['model'],'VAR')[1]
                event['abs_id']=L+event['rel_id']
                var_value=init_val('VAR', event['abs_id']) #read previous value of variable and apply scaling to that  
                psspy.change_cctbusomod_var(event['bus'],event['model'],event['rel_id']+1,float(event['value'])*float(var_value))
                standalone_script+="psspy.change_cctbusomod_var("+str(event['bus'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value'])*float(var_value))+")\n"

        #####################################################################################
        # tap changes change
        elif(event['event_type']=='tap_change_rel'): #change Tx tap during runtime  
#            psspy.bsys(0,0,[0.0,0.0],0,[],2,[event['IDs'][0],event['IDs'][1]],0,[],0,[])
#            ierr, rarray = psspy.amachreal(0, 1, 'GENTAP') 
            ierr, rval = psspy.xfrdat(event['IDs'][0],event['IDs'][1], str(event['IDs'][2]), 'RATIO')
            var_value = rval #read previous value of variable and apply scaling to that 
            event['value'] = 0.5
            psspy.two_winding_chng_6(event['IDs'][0],event['IDs'][1], str(event['IDs'][2]),[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,-1,_i,_i,_i],[_f,_f,_f, float(event['value'])+float(var_value),_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
            standalone_script+="two_winding_chng_6(event['IDs'][0],event['IDs'][1], str(event['IDs'][2]),[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,-1,_i,_i,_i],[_f,_f,_f, float(event['value']+float(var_value),_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)\n"


        #breaker -> can consider to use breaker to clear fault and reclose the line
        elif(str(event['event_type'])=='breaker'): # in case event is breaker being switched
            if(event['switch_status']==1):
                psspy.system_swd_chng (breakers[event['IDs']][0],breakers[event['IDs']][3],r"""1""",[1,_i,_i,_i],_f,[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s)
            elif(event['switch_status']==0):    
                psspy.system_swd_chng (breakers[event['IDs']][0],breakers[event['IDs']][3],r"""1""",[0,_i,_i,_i],_f,[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s)
        # auxiliary breaker
        elif(str(event['event_type'])=='aux_breaker'): # in case event is breaker being switched
            if(event['switch_status']==0):    
                psspy.system_swd_chng (breakers[event['IDs']][0],breakers[event['IDs']][3],r"""1""",[0,_i,_i,_i],_f,[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s)

        #delete event from queue
        if(len(event_queue)>0):
            event_queue = event_queue[1:len(event_queue)]
        else:
            event_queue = []
    
    return event_queue

def order_event_queue(event_queue):
    ordered_event_queue = []
    while(event_queue !=[]):
        first_event=event_queue[0]
        first_index = 0
        for i in range(0,len(event_queue)):
            if event_queue[i]['time']<first_event['time']:
                first_event = event_queue[i]
                first_index = i
        del event_queue[first_index]
        ordered_event_queue.append(first_event)
    return ordered_event_queue
                
        
def add_to_active_faults(scriptFaultID):
    if(active_faults == {}):
        active_faults[scriptFaultID]=1
    else:
        active_faults[scriptFaultID]=max(active_faults[key]for key in active_faults.keys())+1 #counting PSS/E fault ID one up

def get_psseFaultID(scriptFaultID):
    if(active_faults !={}):
        return active_faults[scriptFaultID]
    else:
        return -1
    
#The table of active faults in PSS/E is compressed when a fault is cleared. THe oldest fautl has the ID1. When a new fault is added it gets a higher ID.
#The PSS/E documentation mentions that when a fault is cleared the table is compressed (new fault gets the highest ID)
def remove_from_active_faults(scriptFaultID):
    if(scriptFaultID in active_faults.keys()):
        psseFaultID=active_faults.pop(scriptFaultID)
        #delete key entry from active fault list
        if(active_faults!={}):
            for key in active_faults.keys():
                if(active_faults[key]>psseFaultID):
                    active_faults[key]=active_faults[key]-1
                    
#determine the time until a given event is fully cleared
def determine_clearing_times(fault_details): 
    if('trip_near' in fault_details.keys()):
        if(fault_details['trip_near']!=None):
            floc_clear=float(fault_details['trip_near'])
        else:
            floc_clear=120
    else: floc_clear=120
    
    if('trip_far' in fault_details.keys()):       
        if(fault_details['trip_far']!=None):
            dist_clear=float(fault_details['trip_far'])
        else:
            dist_clear=220
    else: dist_clear = 220
    
    return floc_clear, dist_clear
                   
#initialise network (adding breakers etc.) the breaker IDs will be added to the faults, and event_queue gets initialised
def initialise_network_and_build_event_queue(event_list, breakers, target_dir, sav_file):
    global cb_bus, standalone_script, event_queue
#    psspy.case(target_dir+'\\'+sav_file)
#    precondition_network() # for debugging purposes only
#    cb_bus = checkBusNum(cb_bus)
    for i in range(0, len(event_list)):
        cb_bus = checkBusNum(cb_bus)
        print("starting initialisation")
        print(event_list[i]["Event_Type"])
        if(event_list[i]['run in PSS/E?']=='yes')or(event_list[i]['run in PSS/E?']==1) :
#            if(fault_list[i]['Runback?']=='yes')or(fault_list[i]['Runback?']==1): # defining runback scheme

            #####################################################################################
            # Switching Event: add_event(time, event_type, IDs, switch_status, aux_parameter=0):
            if(event_list[i]['Test Type']=='Switching'):
                if(event_list[i]['Event_Element']=='Line'):
                    floc_bus = int(event_list[i]['i_bus'])
                    dist_bus= int(event_list[i]['j_bus'])
                    element_id = str(event_list[i]['id'])                        
                    if event_list[i]['Event_Type'] == 'ON':
                        add_event(event_list[i]['Event_Time'], 'SWT_LINE', [floc_bus,dist_bus, element_id],1) # Switch the line on/off
                    elif event_list[i]['Event_Type'] == 'OFF':
                        add_event(event_list[i]['Event_Time'], 'SWT_LINE', [floc_bus,dist_bus, element_id],0) # 
                    else:
                        print("check the switching status")

                elif(event_list[i]['Event_Element']=='Tx_2w'):
                    floc_bus = int(event_list[i]['i_bus'])
                    dist_bus= int(event_list[i]['j_bus'])
                    element_id = str(event_list[i]['id'])                        
                    if event_list[i]['Event_Type'] == 'ON':
                        add_event(event_list[i]['Event_Time'], 'SWT_TX_2W', [floc_bus,dist_bus, element_id],1) # 
                    elif event_list[i]['Event_Type'] == 'OFF':
                        add_event(event_list[i]['Event_Time'], 'SWT_TX_2W', [floc_bus,dist_bus, element_id],0) # 
                    else:
                        print("check the switching status")
                        
                elif(event_list[i]['Event_Element']=='Tx_3w'):
                    floc_bus = int(event_list[i]['i_bus'])
                    dist_bus= int(event_list[i]['j_bus'])
                    k_bus= int(event_list[i]['k_bus'])
                    element_id = str(event_list[i]['id'])                        
                    if event_list[i]['Event_Type'] == 'ON':
                        add_event(event_list[i]['Event_Time'], 'SWT_TX_3W', [floc_bus,dist_bus,k_bus, element_id],1) # 
                    elif event_list[i]['Event_Type'] == 'OFF':
                        add_event(event_list[i]['Event_Time'], 'SWT_TX_3W', [floc_bus,dist_bus,k_bus, element_id],0) # 
                    else:
                        print("check the switching status")
                        
                elif(event_list[i]['Event_Element']=='Bus'):
                    floc_bus = int(event_list[i]['i_bus'])
                    if event_list[i]['Event_Type']=='ON':
                        add_event(event_list[i]['Event_Time'], 'SWT_BUS', [floc_bus],1) # bus connected
                    elif event_list[i]['Event_Type']=='OFF':
                        add_event(event_list[i]['Event_Time'], 'SWT_BUS', [floc_bus],0) # bus disconnected
                    else:
                        print("check the switching status")
                        
                elif(event_list[i]['Event_Element']=='Machine'):
                    floc_bus = int(event_list[i]['i_bus'])
                    element_id = str(event_list[i]['id'])                        
#                    add_event(event_list[i]['Event_Time'], 'line', [floc_bus,dist_bus, line_id],0) # Switch off the Tx    
                    if event_list[i]['Event_Type'] == 'ON':
                        add_event(event_list[i]['Event_Time'], 'SWT_MAC', [floc_bus, element_id],1) # 
                    elif event_list[i]['Event_Type'] == 'OFF':
                        add_event(event_list[i]['Event_Time'], 'SWT_MAC', [floc_bus, element_id],0) #             
                    else:
                        print("check the switching status")
                        
                elif(event_list[i]['Event_Element']=='Shunt'):
                    floc_bus = int(event_list[i]['i_bus'])
                    element_id = str(event_list[i]['id'])                        
#                    add_event(event_list[i]['Event_Time'], 'line', [floc_bus,dist_bus, line_id],0) # Switch off the Tx    
                    if event_list[i]['Event_Type'] == 'ON':
                        add_event(event_list[i]['Event_Time'], 'SWT_SHT', [floc_bus, element_id],1) # 
                    elif event_list[i]['Event_Type'] == 'OFF':
                        add_event(event_list[i]['Event_Time'], 'SWT_SHT', [floc_bus, element_id],0) #    
                    else:
                        print("check the switching status")
                        
            #####################################################################################
            # MultiFault event:
            elif(event_list[i]['Test Type']=='Multi_fault'):
                if(event_list[i]['Event_Element']=='Line'):
                    if( ('F resistance' in event_list[i].keys()) and ('F reactance'in event_list[i].keys()) ):
                        imp_R=0.000000
                        imp_X=0.000001 #imp_X=-0.2E+10 #admittance in MVA from impedance in ohms
                    floc_bus = int(event_list[i]['i_bus'])
                    dist_bus= int(event_list[i]['j_bus'])
                    line_id = str(event_list[i]['id'])
                    
                    if event_list[i]['location'] !='':
                        flt_loc = float(event_list[i]['location'])
                    else:
                        flt_loc = 0.05 # 5% from the near bus 0.0001
                        
                    psspy.ltap(floc_bus, dist_bus, str(line_id), flt_loc, cb_bus, 'FLT_BUS', _f)
                    standalone_script+="psspy.ltap("+str(floc_bus)+", "+str(dist_bus) +", '" +str(line_id)+"', " +str(flt_loc)+", "+str(cb_bus)+", 'FLT_BUS', _f)\n"
                   
                    floc_clear, dist_clear = determine_clearing_times(event_list[i])
                    max_clear_time=max(floc_clear, dist_clear)
                    min_clear_time=min(floc_clear, dist_clear)
                                
                    #####initialise and add items to event queue, depending on event
                    
                    #3PHG fault
                    if(event_list[i]['Event_Type']=='3PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_3PHG',[cb_bus,i],1,[imp_R,imp_X]) #here the fault impedance parameters could come from the fault dictionary instead, provided they are listed in the Excel and the input function reads them
                        add_event(event_list[i]['Event_Time']+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],0) # Switch off the line from fault bus to far location
#                        add_event(fault_list[i]['Ftime']+dist_clear, 'bus', cb_bus,0) #fault bus disconnected
                        add_event(event_list[i]['Event_Time']+dist_clear, 'FLT_BUS_3PHG', [cb_bus, i], 0) #clear fault
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],1) # Switch on the line from fault bus to far location
                        
                    #2PHG fault
                    if(event_list[i]['Event_Type']=='2PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_2PHG',[cb_bus,i],1,[imp_R,imp_X]) #here the fault impedance parameters could come from the fault dictionary instead, provided they are listed in the Excel and the input function reads them
                        add_event(event_list[i]['Event_Time']+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],0) # Switch off the line from fault bus to far location
#                        add_event(fault_list[i]['Ftime']+dist_clear, 'bus', cb_bus,0) #fault bus disconnected
                        add_event(event_list[i]['Event_Time']+dist_clear, 'FLT_BUS_2PHG', [cb_bus, i], 0) #clear fault
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],1) # Switch on the line from fault bus to far locati
                        
                    #1PHG fault
                    if(event_list[i]['Event_Type']=='1PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_1PHG',[cb_bus,i],1,[imp_R,imp_X]) #here the fault impedance parameters could come from the fault dictionary instead, provided they are listed in the Excel and the input function reads them
                        add_event(event_list[i]['Event_Time']+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],0) # Switch off the line from fault bus to far location
#                        add_event(fault_list[i]['Ftime']+dist_clear, 'bus', cb_bus,0) #fault bus disconnected
                        add_event(event_list[i]['Event_Time']+dist_clear, 'FLT_BUS_1PHG', [cb_bus, i], 0) #clear fault
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],1) # Switch on the line from fault bus to far locati

            #####################################################################################
            # fault event
            elif(event_list[i]['Test Type']=='Fault'): #If the fault event
                
                # Line fault
                if(event_list[i]['Event_Element']=='Line'):
                    if( ('F resistance' in event_list[i].keys()) and ('F reactance'in event_list[i].keys()) ):
                        imp_R=0.000000 # imp_R=0.1
                        imp_X=0.000001 # imp_X=0.3 imp_X=-0.2E+10 #admittance in MVA from impedance in ohms

                    floc_bus = int(event_list[i]['i_bus'])
                    dist_bus= int(event_list[i]['j_bus'])
                    line_id = str(event_list[i]['id'])
                    if(event_list[i]['arc_time']!='' and event_list[i]['arc_time']!=None):
                        autorec_t=float(event_list[i]['arc_time'])
                    else:
                        autorec_t=-1
                    if(event_list[i]['arc_success']!='' and event_list[i]['arc_success']!=None):
                        autorec_suc = float(event_list[i]['arc_success'])
                    else:
                        autorec_suc=0
                    if event_list[i]['location'] !='':
                        flt_loc = float(event_list[i]['location'])
                    else:
                        flt_loc = 0.0001
                    psspy.ltap(floc_bus, dist_bus, str(line_id), flt_loc, cb_bus, 'FLT_BUS', _f)
                    standalone_script+="psspy.ltap("+str(floc_bus)+", "+str(dist_bus) +", '" +str(line_id)+"', " +str(flt_loc)+", "+str(cb_bus)+", 'FLT_BUS', _f)\n"
                   
                    floc_clear, dist_clear = determine_clearing_times(event_list[i])
                    max_clear_time=max(floc_clear, dist_clear)
                    min_clear_time=min(floc_clear, dist_clear)
                                
                    #####initialise and add items to event queue, depending on event

                    #3PHG fault
                    if(event_list[i]['Event_Type']=='3PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_3PHG',[cb_bus,i],1,[imp_R,imp_X]) #here the fault impedance parameters could come from the fault dictionary instead, provided they are listed in the Excel and the input function reads them
                        add_event(event_list[i]['Event_Time']+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],0) # Switch off the line from fault bus to far location
#                        add_event(fault_list[i]['Ftime']+dist_clear, 'bus', cb_bus,0) #fault bus disconnected
    #                    add_event(fault_list[i]['Ftime']+dist_clear, '3PHG', [cb_bus, i], 0) #clear fault
                        max_reclose_time = 0
                        if (autorec_t <0): 
                            print('done adding event')
                        elif (autorec_t>=0):#    autoreclose
                            if(autorec_suc>0): #autoreclose successful
                                print('successful autoreclosure')
    #                            add_event(fault_list[i]['Ftime']+min_clear_time+autorec_t -0.2, '3PHG', [cb_bus, i], 0) #clear fault 200 ms before auto reclosure
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t -0.2, 'FLT_BUS_3PHG', [cb_bus, i], 0) #clear fault 200 ms before auto reclosure
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_BUS', [cb_bus],1) #fault bus back on
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_LINE', [cb_bus,dist_bus, line_id],1) # Switch on the line from fault bus to near location
                            else:
                                for autorec_cnt in range (0,int(abs(autorec_suc)) ): 
                                    print('unsuccessful autoreclosure')
    #                                add_event(fault_list[i]['Ftime']+(min_clear_time+autorec_t)*(autorec_cnt+1), '3PHG', [cb_bus, i], 0) #clear fault
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1), 'SWT_BUS', [cb_bus],1) #fault bus back on
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1), 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1),'FLT_BUS_3PHG',[cb_bus,i],1,[imp_R,imp_X])
                                    print('opening breakers again')
#                                    add_event(fault_list[i]['Ftime']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'bus', cb_bus,0) #fault bus disconnected
#                                    add_event(fault_list[i]['Ftime']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, '3PHG', [cb_bus, i], 0) #clear fault
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
#                                    add_event(fault_list[i]['Ftime']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'bus', cb_bus,0) #fault bus disconnected
                                    max_reclose_time = (min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear
                        both_lines_disconnect_time = max(max_reclose_time, dist_clear)
                        add_event(event_list[i]['Event_Time']+both_lines_disconnect_time, 'SWT_BUS', [cb_bus],0) #fault bus disconnected

                    #2PHG fault
                    if(event_list[i]['Event_Type']=='2PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_2PHG',[cb_bus,i],1,[imp_R,imp_X]) #here the fault impedance parameters could come from the fault dictionary instead, provided they are listed in the Excel and the input function reads them
                        add_event(event_list[i]['Event_Time']+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],0) # Switch off the line from fault bus to far location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_BUS', [cb_bus],0) #fault bus disconnected
    #                    add_event(fault_list[i]['Ftime']+dist_clear, '2PHG', [cb_bus, i], 0) #clear fault
                        if (autorec_t <0): 
                            print('done adding event')
                        elif (autorec_t>=0):#    autoreclose
                            if(autorec_suc>0): #autoreclose successful
                                print('successful autoreclosure')
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t -0.2, 'FLT_BUS_2PHG', [cb_bus, i], 0) #clear fault 200 ms before auto reclosure
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_BUS', [cb_bus],1) #fault bus back on
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_LINE', [cb_bus,dist_bus, line_id],1) # Switch on the line from fault bus to near location
                            else:
                                for autorec_cnt in range (0,int(abs(autorec_suc)) ):                                    
                                    print('unsuccessful autoreclosure') 
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1), 'SWT_BUS', [cb_bus],1) #fault bus back on
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1), 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1),'FLT_BUS_2PHG',[cb_bus,i],1,[imp_R,imp_X])
                                    
                                    print('opening breakers again')
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'SWT_BUS', [cb_bus],0) #fault bus disconnected
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'FLT_BUS_2PHG', [cb_bus, i], 0) #clear fault
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                    
                    #1PHG fault
                    if(event_list[i]['Event_Type']=='1PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_1PHG',[cb_bus,i],1,[imp_R,imp_X]) #here the fault impedance parameters could come from the fault dictionary instead, provided they are listed in the Excel and the input function reads them
                        add_event(event_list[i]['Event_Time']+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_LINE', [cb_bus,dist_bus, line_id],0) # Switch off the line from fault bus to far location
                        add_event(event_list[i]['Event_Time']+dist_clear, 'SWT_BUS', [cb_bus],0) #fault bus disconnected
    #                    add_event(fault_list[i]['Ftime']+dist_clear, '2PHG', [cb_bus, i], 0) #clear fault
                        if (autorec_t <0): 
                            print('done adding event')
                        elif (autorec_t>=0):#    autoreclose
                            if(autorec_suc>0): #autoreclose successful
                                print('successful autoreclosure')
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t -0.2, 'FLT_BUS_1PHG', [cb_bus, i], 0) #clear fault 200 ms before auto reclosure
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_BUS', [cb_bus],1) #fault bus back on
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                                add_event(event_list[i]['Event_Time']+min_clear_time+autorec_t, 'SWT_LINE', [cb_bus,dist_bus, line_id],1) # Switch on the line from fault bus to near location
                            else:
                                for autorec_cnt in range (0,int(abs(autorec_suc)) ):                                    
                                    print('unsuccessful autoreclosure') 
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1), 'SWT_BUS', [cb_bus],1) #fault bus back on
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1), 'SWT_LINE', [floc_bus,cb_bus, line_id],1) # Switch on the line from fault bus to near location
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1),'FLT_BUS_1PHG',[cb_bus,i],1,[imp_R,imp_X])
                                    
                                    print('opening breakers again')
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'SWT_BUS', [cb_bus],0) #fault bus disconnected
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'FLT_BUS_1PHG', [cb_bus, i], 0) #clear fault
                                    add_event(event_list[i]['Event_Time']+(min_clear_time+autorec_t)*(autorec_cnt+1)+floc_clear, 'SWT_LINE', [floc_bus,cb_bus, line_id],0) # Switch off the line from fault bus to near location
                    
                    #LL fault
                #Transformer fault at i_bus and then transformer being switched off and fault being cleared
                elif(event_list[i]['Event_Element']=='Tx_2w'):
                    print('2 windings transformer fault')
                    add_event(event_list[i]['Event_Time'], 'FLT_TX_2W_3PHG', [event_list[i]['i_bus'], event_list[i]['j_bus'], str(event_list[i]['id'])],1) # apply fault
                    add_event(event_list[i]['Event_Time']+(float(event_list[i]['trip_near'])), 'SWT_TX_2W', [event_list[i]['i_bus'], event_list[i]['j_bus'], str(event_list[i]['id'])],0) # deactivate the Tx
                    
                elif(event_list[i]['Event_Element']=='Tx_3w'):
                    print('3 windings transformer fault')
                    add_event(event_list[i]['Event_Time'], 'FLT_TX_3W_3PHG', [event_list[i]['i_bus'], event_list[i]['j_bus'], event_list[i]['k_bus'], str(event_list[i]['id'])],1)
                    add_event(event_list[i]['Event_Time']+(float(event_list[i]['trip_near'])), 'SWT_TX_3W', [event_list[i]['i_bus'], event_list[i]['j_bus'], event_list[i]['k_bus'], str(event_list[i]['id'])],0)
                    
                #BUS fault--> no auto reclosure permitted
                elif(event_list[i]['Event_Element']=='Bus'):
                    if(event_list[i]['F reactance']!=None and event_list[i]['F resistance']!=None):
                        imp_X=float(event_list[i]['F reactance'])
                        imp_R=float(event_list[i]['F resistance'])
                    else:
                        imp_R=0.1
                        imp_X=0.5
                    print('bus fault')
                    if(event_list[i]['Event_Type']=='3PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_3PHG',[event_list[i]['i_bus']],1,[imp_R,imp_X]) 
                        add_event(event_list[i]['Event_Time']+(float(event_list[i]['trip_near'])), 'SWT_BUS', [event_list[i]['i_bus']],1,0) #disconnect faulted bus after clearing time

                    if(event_list[i]['Event_Type']=='2PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_2PHG',[event_list[i]['i_bus']],1,[imp_R,imp_X]) 
                        add_event(event_list[i]['Event_Time']+(float(event_list[i]['trip_near'])), 'SWT_BUS', [event_list[i]['i_bus']],1,0) #disconnect faulted bus after clearing time

                    if(event_list[i]['Event_Type']=='1PHG'):
                        add_event(event_list[i]['Event_Time'],'FLT_BUS_1PHG',[event_list[i]['i_bus']],1,[imp_R,imp_X]) 
                        add_event(event_list[i]['Event_Time']+(float(event_list[i]['trip_near'])), 'SWT_BUS', [event_list[i]['i_bus']],1,0) #disconnect faulted bus after clearing time

            #####################################################################################
            # Transformer tap event                        
            elif(event_list[i]['Test Type']=='Tx_tap_profile'): # If the voltage setpoint test in network model
                if event_list[i]['i_bus'] != '' and event_list[i]['j_bus'] != '':
                    floc_bus = int(event_list[i]['i_bus'])
                    dist_bus= int(event_list[i]['j_bus'])
                    element_id = str(event_list[i]['id'])                        
                    add_event(event_list[i]['Event_Time'], 'tap_change_rel', [floc_bus,dist_bus, element_id],1,event_list[i]['Test profile']) # 
                else:
                    pass
                
            #####################################################################################
            # Control event                        
            elif(event_list[i]['Test Type']=='V_stp_profile'): # If the voltage setpoint test in network model
                
                profile=ProfilesDict[event_list[i]['Test profile']]
                scaling_factor=profile['scaling_factor_PSSE']
                if(not is_number(scaling_factor)):
                    scaling_factor=1.0
                offset=profile['offset_PSSE']
                if(not is_number(offset)):
                    offset=0.0   
                profile=interpolate(profile=profile, TimeStep=0.001, density=20.0, scaling=scaling_factor, offset=offset )  

                Vset_params=[]
                for key in PSSEmodelDict.keys():
                    if('Vset' in key):
                        Vset_params.append(key)
                Vset_cnt=1
                Vset_dict={}
                while ( any( 'Vset'+str(Vset_cnt) in key for key in Vset_params)):
                    Vset_dict[Vset_cnt]={}
                    for param_cnt in range(0, len(Vset_params)):                
                        param=Vset_params[param_cnt]
                        if('Vset'+str(Vset_cnt) in param):
                            Vset_dict[Vset_cnt][param.replace('Vset'+str(Vset_cnt)+'_', '')]=PSSEmodelDict[param]
                             
                    Vset_cnt+=1
                for Vset_inst in Vset_dict.keys(): #iterate over all instances of voltage setpoints requiring to be changes (e.g. across different machines/control systems)
                    
                    if('model' in Vset_dict[Vset_inst].keys()): #It means the setpoint that needs to be changes is a variable or constant
                        #L = psspy.mdlind(322813, '1 ', 'EXC', 'VAR')[1]
                        #L=psspy.mdlind(Vset_dict[Vset_inst]['bus'], Vset_dict[Vset_inst]['mac'], Vset_dict[Vset_inst]['type'], 'VAR')[1]
                        if('var' in Vset_dict[Vset_inst].keys()):
                            if(profile['scaling']=='relative'):
                                for cnt in range(0, len(profile['x_data'])):
                                    event_queue.append({'time':profile['x_data'][cnt], 'event_type':'var_change_rel', 'rel_id': Vset_dict[Vset_inst]['var'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                            elif(profile['scaling']=='absolute'):
                                for cnt in range(0, len(profile['x_data'])):
                                    event_queue.append({'time':profile['x_data'][cnt], 'event_type':'var_change_abs', 'rel_id': Vset_dict[Vset_inst]['var'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                        elif('con' in Vset_dict[Vset_inst].keys()): 
                            if(profile['scaling']=='relative'):
                                for cnt in range(0, len(profile['x_data'])):
                                    event_queue.append({'time':profile['x_data'][cnt], 'event_type':'con_change_rel', 'rel_id': Vset_dict[Vset_inst]['con'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                            elif(profile['scaling']=='absolute'):
                                for cnt in range(0, len(profile['x_data'])):
                                    event_queue.append({'time':profile['x_data'][cnt], 'event_type':'con_change_abs', 'rel_id': Vset_dict[Vset_inst]['con'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                            
                        #write every point from the profile vector into the event queue as a variable change
                        pass
                    else: #It means the setpoint that needs to be changed is in the PSS/E Vref vector
                        #write ever point from the profile vector into the event queue as a Vred change
                        if(profile['scaling']=='relative'):
                            for cnt in range(0, len(profile['x_data'])):
                                event_queue.append({'time':profile['x_data'][cnt], 'type':'VREF_change_rel', 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'chn':Vset_dict[Vset_inst]['chn'],'value':profile['y_data'][cnt]})
                        elif(profile['scaling']=='absolute'):
                            for cnt in range(0, len(profile['x_data'])):
                                event_queue.append({'time':profile['x_data'][cnt], 'type':'VREF_change_abs', 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})                    
                event_queue=order_event_queue(event_queue)
                total_duration=event_queue[-1]['time']+5

    return 0    

#initialise dynamic simulation
def init_dynamics(target_dir, dyr_file, dll_files): #maybe set to ESCO standards instead, using the library
    global standalone_script 
    subdir = target_dir
    #convert, factorise, tysl
    ############################################################################
#    #Convert generators from their power flow representation in preparation for switching studies and dynamic simulations
#    psspy.cong(0)
#    
#    # Define the Tasmanian subsystem, initialise for load conversion, convert loads, post-processing housekeeping
#    psspy.bsys(sid=0,numarea=1,areas=[7])
#    psspy.conl(apiopt=1)
#    psspy.conl(0,0,2,[0,0],[91.27, 19.36,-126.88, 188.43])
#    
#    # Define a subsystem for APD loads and convert loads
#    psspy.bsys(sid=1,numbus=2,buses=[37580,38588])
#    psspy.conl(1,0,2,[0,0],[52.75, 58.13, 5.97, 95.52])
#    
#    # Define a subsystem for Tomago loads and convert loads
#    psspy.bsys(sid=1,numbus=1,buses=[21790])
#    psspy.conl(1,0,2,[0,0],[86.63, 25.19, -378.97, 347.97])
#    
#    # Define a subsystem for Boyne Island loads and convert loads
#    psspy.bsys(sid=1,numbus=3,buses=[45081,45082,45088])
#    psspy.conl(1,0,2,[0,0],[51.36, 59.32,-228.04, 254.01])
#    
#    # Convert remaining loads in NEM
#    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,-306.02, 303.0])
#    psspy.conl(apiopt=3)
#    
#    # Preserve sparsity of network matrices, factorise network admittance matrix, adjust iteration limit and acceleration factor
#    psspy.ordr(0)
#    psspy.fact()
#    #psspy.solution_parameters_4(intgar3=600, realar7=0.2)
#    #psspy.dynamics_solution_param_2(intgar1=600, realar1=0.2, realar3=0.001, realar4=0.008)
    
    ################################################################
    psspy.cong(0)
    standalone_script+='psspy.cong(0)\n'
    psspy.bsys(0,0,[0.0, 500.],1,[7],0,[],0,[],0,[])
    standalone_script+='psspy.bsys(0,0,[0.0, 500.],1,[7],0,[],0,[],0,[])\n'
#    psspy.bsys(0,0,[0.0, 500.],0,[],0,[],0,[],0,[])
#    standalone_script+='psspy.bsys(0,0,[0.0, 500.],0,[],0,[],0,[],0,[])\n'
    psspy.conl(0,0,1,[0,0],[ 91.27, 19.36,-126.88, 188.43])
    standalone_script+='psspy.conl(0,0,1,[0,0],[ 91.27, 19.36,-126.88, 188.43])\n'
    psspy.conl(0,0,2,[0,0],[ 91.27, 19.36,-126.88, 188.43])
    standalone_script+='psspy.conl(0,0,2,[0,0],[ 91.27, 19.36,-126.88, 188.43])\n'
    # start powercor loads
    psspy.bsys(1,0,[0.0,0.0],0,[],8,[37703,37704,37705,37706,37707,37708,37714,37747],0,[],0,[])
    standalone_script+='psspy.bsys(1,0,[0.0,0.0],0,[],8,[37703,37704,37705,37706,37707,37708,37714,37747],0,[],0,[])\n'
    psspy.conl(1,0,2,[0,0],[0.0, 40.0,0.0, 40.0])
    standalone_script+='psspy.conl(1,0,2,[0,0],[0.0, 40.0,0.0, 40.0])\n'
    # end powercor loads
    psspy.bsys(1,0,[0.0,0.0],0,[],6,[37600,37601,37602,37580,37584,38588],0,[],0,[])
    standalone_script+='psspy.bsys(1,0,[0.0,0.0],0,[],6,[37600,37601,37602,37580,37584,38588],0,[],0,[])\n'
    psspy.conl(1,0,2,[0,0],[ 52.75, 58.13, 5.97, 95.52])
    standalone_script+='psspy.conl(1,0,2,[0,0],[ 52.75, 58.13, 5.97, 95.52])\n'
    psspy.bsys(1,0,[0.0,0.0],0,[],1,[21790],0,[],0,[])
    standalone_script+='psspy.bsys(1,0,[0.0,0.0],0,[],1,[21790],0,[],0,[])\n'
    psspy.conl(1,0,2,[0,0],[ 86.63, 25.19, -378.97, 347.97])
    standalone_script+='psspy.conl(1,0,2,[0,0],[ 86.63, 25.19, -378.97, 347.97])\n'
    psspy.bsys(1,0,[0.0,0.0],0,[],1,[45082],0,[],0,[])
    standalone_script+='psspy.bsys(1,0,[0.0,0.0],0,[],1,[45082],0,[],0,[])\n'
    psspy.conl(1,0,2,[0,0],[ 51.36, 59.32,-228.04, 254.01])
    standalone_script+='psspy.conl(1,0,2,[0,0],[ 51.36, 59.32,-228.04, 254.01])\n'
    psspy.bsys(1,0,[0.0,0.0],0,[],9,[40320,40340,40350,40970,40980,40990,41050,41071,41120],0,[],0,[])
    standalone_script+='psspy.bsys(1,0,[0.0,0.0],0,[],9,[40320,40340,40350,40970,40980,40990,41050,41071,41120],0,[],0,[])\n'
    psspy.conl(1,0,2,[0,0],[ 100.0,0.0,0.0, 100.0])
    standalone_script+='psspy.conl(1,0,2,[0,0],[ 100.0,0.0,0.0, 100.0])\n'
    psspy.bsys(0,0,[0.0, 500.],0,[],0,[],0,[],0,[])
    standalone_script+='psspy.bsys(0,0,[0.0, 500.],0,[],0,[],0,[],0,[])\n'
    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,-306.02, 303.0])
    standalone_script+='psspy.conl(0,1,2,[0,0],[ 100.0,0.0,-306.02, 303.0])\n'
    psspy.conl(0,1,3,[0,0],[ 100.0,0.0,-306.02, 303.0])
    standalone_script+='psspy.conl(0,1,3,[0,0],[ 100.0,0.0,-306.02, 303.0])\n'
    psspy.ordr(0)
    standalone_script+='psspy.ordr(0)\n'
    psspy.fact()
    standalone_script+='psspy.fact()\n'
    psspy.tysl(0)    
    standalone_script+='psspy.tysl(0)\n'
    
    
    ###########################################################################
    
    # Use present voltage vector as starting point in network solution
    #psspy.tysl(0)
    
    # Add DYR - dynamic data loading
    psspy.dyre_new([1,1,1,1], subdir+'\\'+dyr_file, "","","")
    standalone_script+="psspy.dyre_new([1,1,1,1],'"+str(os.path.split(dyr_file)[1])+"', '','', '')\n"

    # Add more DYR
    target_dir_DYR=target_dir+"\\"+"DYRs"
    for root, dirs, files in os.walk(target_dir_DYR):
        for file in files:
            if '.dyr' in file:
                psspy.dyre_add([_i,_i,_i,_i],os.path.join(root, file),"","")   
                standalone_script+="psspy.dyre_add([_i,_i,_i,_i],'"+str('DYRs\\'+file)+"', '','')\n"
                
    # Add the DLL                
    for dll_cnt in range(0, len(dll_files)):
        if os.path.isfile(subdir+'\\'+dll_files[dll_cnt]):
            psspy.addmodellibrary(subdir+'\\'+dll_files[dll_cnt])
            standalone_script+="psspy.addmodellibrary('"+str('DLLs\\'+dll_files[dll_cnt])+"')\n"
    # additional DLL from committed generators
    target_dir_DLL=target_dir+"\\"+"DLLs"
    for root, dirs, files in os.walk(target_dir_DLL):
        for file in files:
            if '.dll' in file:
                psspy.addmodellibrary(os.path.join(root, file))  
                standalone_script+="psspy.addmodellibrary('"+str('DLLs\\'+file)+"')\n"
                try:
    #                shutil.copyfile(zingen, model_copy_dir+'\\dsusr_zingen.dll')
                    shutil.copyfile(os.path.join(root, file), target_dir_DLL+'\\'+file) #copy dll file to the main folder for the standalone_script
                except:
                    pass
    
    #set output file format to old type
    #psspy.set_chnfil_type(0)
    
    #dynamic solution parameters
    #psspy.dynamics_solution_param_2([200,_i,_i,_i,_i,_i,_i,_i],[1.0,0.0001, 0.001, 0.008, 0.06, 0.14, 1.0, 0.0005])
    #psspy.solution_parameters_4(intgar3=600, realar7=0.2)
    #psspy.dynamics_solution_param_2(intgar1=600, realar1=0.2, realar3=0.001, realar4=0.008)

    psspy.dynamics_solution_param_2([999,_i,_i,_i,_i,_i,_i,_i],[ 0.2,_f, 0.001,_f,_f,_f, 0.2,_f])
    #psspy.dynamics_solution_param_2([999,_i,_i,_i,_i,_i,_i,_i],[ 0.2,_f, 0.001,0.016,_f,_f, 0.2,_f])
    standalone_script+="psspy.dynamics_solution_param_2([999,_i,_i,_i,_i,_i,_i,_i],[0.2,_f,0.001,_f,_f,_f,0.2,_f])\n"
    
    psspy.set_netfrq(1)
    standalone_script+="psspy.set_netfrq(1)\n"

    swing_buses = [233201, 233202, 233203, 233204] #2ERARNG 1, 2, 3, 4 # each case may use one of these buses as swing bus
    s_bus = 233202 # asumming 233202 being the swing bus
    for s_bus_loop in swing_buses: # loop through other bus numer if that is the swing bus, then get that one.
        ierr, bus_type = psspy.busint(s_bus_loop,'TYPE')
        if bus_type == 3: 
            s_bus = s_bus_loop
            break
    psspy.set_relang(1, s_bus, '1') #2ERARNG___G323.000 20101 233202
    standalone_script+="psspy.set_relang(1, "+str(s_bus)+", '1')\n"
    
def init_simulation():
    global standalone_script
    MYOUTFILE = psspy.sfiles()[0].rstrip(".sav")+".out"
    print 'MYOUTFILE='
    print(MYOUTFILE)
    
    psspy.strt_2([0,1],MYOUTFILE)
#    standalone_script+="psspy.strt_2([0,1],+"+str(MYOUTFILE)+")\n"
    standalone_script+="psspy.strt_2([0,1],'"+str(os.path.split(MYOUTFILE)[1])+"')\n"
#    psspy.strt(0,MYOUTFILE)
#    psspy.strt(0,MYOUTFILE)
#    psspy.strt(0,MYOUTFILE)
#    psspy.strt(0,MYOUTFILE)
    return MYOUTFILE

def add_channels():
    global standalone_script    
    psspy.delete_all_plot_channels()
    standalone_script+="psspy.delete_all_plot_channels()\n"
    #modify routine to generate channels from Scenario definition spreadsheet instead.
    #add P and Q signals for all branches in Branch_Lib
    #Add V, F, Ang signals for all buses in Bus Lib
    #add P, Q and F, V, Ang for buses and connecting branches from Locations of interest (such as POC) as listed in PSS/E sheet, 
    #Add all other plot channels from ModelDetailsPSSE tab in spreadsheet
  

    meas_locs={}
    for key in PSSEmodelDict.keys():
        if('fromBus' in key):
            loc=key.replace('fromBus','')
            meas_locs[loc]={'fromBus':PSSEmodelDict[key], 'toBus':PSSEmodelDict[key.replace('fromBus', 'toBus')], 'measBus':PSSEmodelDict[key.replace('fromBus', 'measBus')]}
    psspy.delete_all_plot_channels()
#    standalone_script+="psspy.delete_all_plot_channels()\n"
    chn_idx=1
    for loc in meas_locs.keys():
        psspy.voltage_and_angle_channel([chn_idx,-1,-1,meas_locs[loc]['measBus']], ['U_'+loc, 'ANG_'+loc])
        standalone_script+="psspy.voltage_and_angle_channel(["+str(chn_idx)+",-1,-1,"+str(meas_locs[loc]['measBus'])+"], ['U_"+str(loc)+"', 'ANG_"+str(loc)+"'])\n"
        plot_channels['U_'+loc]=chn_idx
        plot_channels['ANG_'+loc]=chn_idx+1
        chn_idx+=2
        psspy.branch_p_and_q_channel([chn_idx,-1,-1,meas_locs[loc]['fromBus'], meas_locs[loc]['toBus']], r"""1""", ['P_'+loc, 'Q_'+loc])
        standalone_script+="psspy.branch_p_and_q_channel(["+str(chn_idx)+",-1,-1,"+str(meas_locs[loc]['fromBus'])+", "+str(meas_locs[loc]['toBus'])+"], '1', ['P_"+str(loc)+"', 'Q_"+str(loc)+"'])\n"
        plot_channels['P_'+loc]=chn_idx
        plot_channels['Q_'+loc]=chn_idx+1
        chn_idx+=2
        psspy.bus_frequency_channel([chn_idx,meas_locs[loc]['measBus']], 'F_'+loc)
        standalone_script+="psspy.bus_frequency_channel(["+str(chn_idx)+","+str(meas_locs[loc]['measBus'])+"], 'F_"+str(loc)+"')\n"
        plot_channels['F_'+loc]=chn_idx
        chn_idx+=1

    #----------------------------USER INPUT REQUIRED---------------------------

    ierr, L_PPC = psspy.cctmind_buso(9920,'SMAHYCF14','VAR')
    ierr, K_PPC = psspy.cctmind_buso(9920,'SMAHYCF14','STATE')
    ierr, L_INV = psspy.mdlind(9942,'1','GEN','VAR')
    ierr, L_INV2 = psspy.mdlind(9944,'1','GEN','VAR')

    try:
        psspy.var_channel([-1,L_INV+102],'INV1_FRT_FLAG' ) # FRT detection flag
        psspy.var_channel([-1,L_INV+163],'INV1_FRT_STATE' ) # Internal signal FRT detection: Frt_State
        psspy.var_channel([-1,L_INV+186],'INV1_LVRT' )
        psspy.var_channel([-1,L_INV+187],'INV1_HVRT' )
        psspy.var_channel([-1,L_INV+9],'INV1_IQ_COMMAND' ) #Iq command before dynamic limitation block
        psspy.var_channel([-1,L_INV+40],'INV1_IP_COMMAND' )
        psspy.var_channel([-1,L_INV+174],'INV1_P_COMMAND' )
        psspy.var_channel([-1,L_INV+175],'INV1_Q_COMMAND' )
        
        psspy.var_channel([-1,L_INV2+200],'INV2_Frequency' )
    #    psspy.var_channel([-1,L_INV2+83],'INV2_Vq' )    
        psspy.var_channel([-1,L_INV2+82],'INV2_Vd' )
        psspy.var_channel([-1,L_INV2+83],'INV2_Vq' )
        psspy.var_channel([-1,L_INV2+86],'INV2_Id' )
        psspy.var_channel([-1,L_INV2+87],'INV2_Iq' )
        psspy.var_channel([-1,L_INV2+16],'INV2_FRT_FLAG' ) # FRTDetect – Flag to Check whether FRT is enabled or not
        psspy.var_channel([-1,L_INV2+79],'INV2_ANGLE' ) # Inverter_Voltage_Angle
        psspy.var_channel([-1,L_INV2+17],'VI_X' ) # Reactive part (Virtual Impedance when FRT)
        psspy.var_channel([-1,L_INV2+18],'VI_R' ) # Real part (Virtual Impedance when FRT)
        
        psspy.var_channel([-1,L_PPC+14],'QREF_POC' ) # POI_Var_Spt
        psspy.var_channel([-1,L_PPC+15],'PFREF' ) # POI_Var_Spt
        psspy.var_channel([-1,L_PPC+16],'VREF_POC' ) # POI_Vol_Spt
        psspy.var_channel([-1,L_PPC+17],'HZREF_POC' ) # POI_Hz_Spt
        psspy.var_channel([-1,L_PPC+18],'PREF_POC' ) # PwrAtLimSales
        psspy.var_channel([-1,L_PPC+1],'VOLT_RB' )
        psspy.var_channel([-1,L_PPC+2],'P_PCC' )
        psspy.var_channel([-1,L_PPC+3],'Q_PCC' )
        psspy.var_channel([-1,L_PPC+4],'S_PCC' )
        psspy.var_channel([-1,L_PPC+56],'P_CMD_PV' )
        psspy.var_channel([-1,L_PPC+57],'Q_CMD_PV' )
        psspy.var_channel([-1,L_PPC+58],'P_CMD_BESS' )
        psspy.var_channel([-1,L_PPC+59],'Q_CMD_BESS' )
        psspy.var_channel([-1,L_PPC+75],'FrtActive' )
        psspy.var_channel([-1,L_PPC+76],'FRT_ExitTm' )
        
    except: 
        pass
    

    ######################################################################
    # PPC
#    psspy.var_channel([-1,L_PPC+14],'QREF_POC' ) # POI_Var_Spt
#    psspy.var_channel([-1,L_PPC+15],'PFREF' ) # POI_Var_Spt
#    psspy.var_channel([-1,L_PPC+16],'VREF_POC' ) # POI_Vol_Spt
#    psspy.var_channel([-1,L_PPC+17],'HZREF_POC' ) # POI_Hz_Spt
#    psspy.var_channel([-1,L_PPC+18],'PREF_POC' ) # PwrAtLimSales
#    psspy.var_channel([-1,L_PPC+1],'VOLT_RB' )
#    psspy.var_channel([-1,L_PPC+2],'P_PCC' )
#    psspy.var_channel([-1,L_PPC+3],'Q_PCC' )
#    psspy.var_channel([-1,L_PPC+4],'S_PCC' )
#    psspy.var_channel([-1,L_PPC+56],'P_CMD_PV' )
#    psspy.var_channel([-1,L_PPC+57],'Q_CMD_PV' )
#    psspy.var_channel([-1,L_PPC+58],'P_CMD_BESS' )
#    psspy.var_channel([-1,L_PPC+59],'Q_CMD_BESS' )
#    psspy.var_channel([-1,L_PPC+75],'FrtActive' )
#    psspy.var_channel([-1,L_PPC+76],'FRT_ExitTm' )
#
#    # INV1
#    psspy.var_channel([-1,L_INV+102],'INV1_FRT_FLAG' ) # FRT detection flag
#    psspy.var_channel([-1,L_INV+163],'INV1_FRT_STATE' ) # Internal signal FRT detection: Frt_State
#    psspy.var_channel([-1,L_INV+186],'INV1_LVRT' )
#    psspy.var_channel([-1,L_INV+187],'INV1_HVRT' )
#    psspy.var_channel([-1,L_INV+9],'INV1_IQ_COMMAND' ) #Iq command before dynamic limitation block
##    psspy.var_channel([-1,L_INV+7],'INV1_IP_COMMAND' ) #Id command in non-FRT situations
#    psspy.var_channel([-1,L_INV+40],'INV1_IP_COMMAND' )
#    psspy.var_channel([-1,L_INV+174],'INV1_P_COMMAND' )
#    psspy.var_channel([-1,L_INV+175],'INV1_Q_COMMAND' )
#
#    # INV2
#    psspy.var_channel([-1,L_INV2+200],'INV2_Frequency' )
##    psspy.var_channel([-1,L_INV2+83],'INV2_Vq' )    
#    psspy.var_channel([-1,L_INV2+82],'INV2_Vd' )
#    psspy.var_channel([-1,L_INV2+83],'INV2_Vq' )
#    psspy.var_channel([-1,L_INV2+86],'INV2_Id' )
#    psspy.var_channel([-1,L_INV2+87],'INV2_Iq' )
#    psspy.var_channel([-1,L_INV2+16],'INV2_FRT_FLAG' ) # FRTDetect – Flag to Check whether FRT is enabled or not
#    psspy.var_channel([-1,L_INV2+79],'INV2_ANGLE' ) # Inverter_Voltage_Angle
##    psspy.var_channel([-1,L_INV2+186],'INV2_LVRT' )
##    psspy.var_channel([-1,L_INV2+187],'INV2_HVRT' )
##    ierr = chsb(sid, all, status)
##    ierr = machine_array_channel(status, id, ident)
##    ierr = psspy.chsb(9944, 0, [-1, -1, -1, 1, 2, 1]) #STATUS(5)=2 PELEC
##    ierr = psspy.machine_array_channel(2, '1', "PELEC") #STATUS(2)=2 PELEC, machine electrical power (pu on SBASE)

    ######################################################################    
    
#
    #psspy.machine_array_channel([chn_idx,5,100001],'1','Q_CMD_TO_INV')
    #plot_channels['Q_CMD_TO_INV']=chn_idx
    #chn_idx+=1
    #psspy.machine_array_channel([chn_idx,8,100001],'1','P_CMD_TO_INV')
    #plot_channels['P_CMD_TO_INV']=chn_idx
    #chn_idx+=1    
    #psspy.machine_array_channel([chn_idx,10,100001],'1','F_CMD_TO_INV')
    #plot_channels['F_CMD_TO_INV']=chn_idx
    #chn_idx+=1
    

    
#    psspy.var_channel([-1,L_INV+102],'INV FRT flag' ) # FRT detection flag
#    psspy.var_channel([-1,L_INV+186],'INV LVRT' )
#    psspy.var_channel([-1,L_INV+187],'INV HVRT' )
#    psspy.var_channel([-1,L_INV+9],'Iq Command' )
#    psspy.var_channel([-1,L_INV+40],'Ip Command' )
#
 
    
#    psspy.var_channel([-1,L_INV2+10],'INV2_Vd' )
#    psspy.var_channel([-1,L_INV2+11],'INV2_Vq' )
#    psspy.var_channel([-1,L_INV2+12],'INV2_Id' )
#    psspy.var_channel([-1,L_INV2+13],'INV2_Iq' )




    


#    LSM_330     = 250490
#    COFF_330    = 226893
#    LSM_132     = 250401
#    KOLK_132    = 245641
#    COFF_132    = 226843
#    LSM_132_B   = 250040
#    CASN_132    = 294840
#    ARM_132     = 211640
#    SUM_POC_DUM = 800009
#    SUM_POC     = 9910

#    LSM_330     = 250490 # Lismore 330kV_A
#    COFF_330    = 226893 # Coffs Harbour 330kV
#    LSM_132     = 250401 # Lismore 132kV
#    KOLK_132    = 245641 # Koolkhan 132kV
#    COFF_132    = 226843 # Coffs Harbour 132kV
#    LSM_132_B   = 250040 # Lismore 132kV_2
#    CASN_132    = 294840 # Casino 132kV
#    ARM_132     = 211640 # Armidale 132kV
#    SUM_POC_DUM = 800009 # SUMSF DM 132kV
#    SUM_POC     = 9910 # SUM_POC 132kV

#    psspy.voltage_channel([-1,-1,-1,LSM_330],r"""V_LSM_330""")
#    psspy.voltage_channel([-1,-1,-1,COFF_330],r"""V_COFF_330""")
#    psspy.voltage_channel([-1,-1,-1,LSM_132],r"""V_LSM_132""")
#    psspy.voltage_channel([-1,-1,-1,KOLK_132],r"""V_KOLK_132""")
#    psspy.voltage_channel([-1,-1,-1,COFF_132],r"""V_COFF_132""")
#    psspy.voltage_channel([-1,-1,-1,CASN_132],r"""V_CASN_132""")
#    psspy.voltage_channel([-1,-1,-1,ARM_132],r"""V_ARM_132""")
#    psspy.voltage_channel([-1,-1,-1,SUM_POC_DUM],r"""V_SUM_POC""")

    LSM_330     = 250490 # Lismore 330kV_A
    COFF_330    = 226893 # Coffs Harbour 330kV
    LSM_132     = 250401 # Lismore 132kV
    KOLK_132    = 245641 # Koolkhan 132kV
    COFF_132    = 226843 # Coffs Harbour 132kV
    LSM_132_B   = 250040 # Lismore 132kV_2
    CASN_132    = 294840 # Casino 132kV
    ARM_132     = 211640 # Armidale 132kV
    SUM_POC_DUM = 800009 # SUMSF DM 132kV
    SUM_POC     = 9910 # SUM_POC 132kV
                       
                       
                      
        
    psspy.voltage_channel([-1,-1,-1,LSM_330],r"""V_LSM_330""")
    psspy.voltage_channel([-1,-1,-1,COFF_330],r"""V_COFF_330""")
    psspy.voltage_channel([-1,-1,-1,LSM_132],r"""V_LSM_132""")
    psspy.voltage_channel([-1,-1,-1,KOLK_132],r"""V_KOLK_132""")
    psspy.voltage_channel([-1,-1,-1,COFF_132],r"""V_COFF_132""")
    psspy.voltage_channel([-1,-1,-1,CASN_132],r"""V_CASN_132""")
    psspy.voltage_channel([-1,-1,-1,ARM_132],r"""V_ARM_132""")
    psspy.voltage_channel([-1,-1,-1,SUM_POC_DUM],r"""V_SUM_POC""")  
    
#    #Power from neighbouring SFs:
#    psspy.voltage_channel([-1,-1,-1,36716],r"""V_NSF_66""")
#    psspy.branch_p_and_q_channel([-1,-1,-1,36713, 36716], r"""1""", ['P_NSF1', 'Q_NSF1'])# Numurkah SF
#    psspy.branch_p_and_q_channel([-1,-1,-1,36712, 36716], r"""1""", ['P_NSF2', 'Q_NSF2'])
#    
#    psspy.voltage_channel([-1,-1,-1,1000],r"""V_WSF_66""")
#    psspy.branch_p_and_q_channel([-1,-1,-1,1000, 88888], r"""1""", ['P_WSF', 'Q_WSF']) # Wunghnu SF
#    
#    psspy.voltage_channel([-1,-1,-1,99999],r"""V_GSF_66""")
#    psspy.branch_p_and_q_channel([-1,-1,-1,99900, 99999], r"""1""", ['P_GSF', 'Q_GSF']) # Girgarre SF
#    
#    psspy.voltage_channel([-1,-1,-1,106],r"""V_GlSF_66""")
#    psspy.branch_p_and_q_channel([-1,-1,-1,105, 106], r"""1""", ['P_GlSF', 'Q_GlSF']) # Glenrowan SF
#    
#    psspy.voltage_channel([-1,-1,-1,1810],r"""V_GlwSF_66""")
#    psspy.branch_p_and_q_channel([-1,-1,-1,1820, 1810], r"""1""", ['P_GlwSF', 'Q_GlwSF']) # Glenrowan West SF
#    
#    psspy.voltage_channel([-1,-1,-1,36250],r"""V_WanSF_66""")
#    psspy.branch_p_and_q_channel([-1,-1,-1,1106, 36250], r"""1""", ['P_WanSF', 'Q_WanSF']) # Wangaratta SF
#    
#    psspy.voltage_channel([-1,-1,-1,36250],r"""V_WinSF_66""")
#    psspy.branch_p_and_q_channel([-1,-1,-1,800, 36250], r"""1""", ['P_WinSF', 'Q_WinSF']) # Wangaratta SF        
#        
#    for i in range (0, len(graph_info)):
#        psspy.voltage_and_angle_channel([-1,-1,-1,int(graph_info[i][0])], ['U_'+str(graph_info[i][2]), 'ANG_'+str(graph_info[i][2])])
#        psspy.branch_p_and_q_channel([-1,-1,-1,int(graph_info[i][0]), int(graph_info[i][1])], r"""1""", ['P_'+str(graph_info[i][2]), 'Q_'+str(graph_info[i][2])  ])
#        psspy.bus_frequency_channel([-1,int(graph_info[i][0])], 'F_'+str(graph_info[i][2]))
#    
#    
#    ierr, ppc_var_id = psspy.mdlind(500,'1','EXC','VAR')
#    ierr, ppc_state_id = psspy.mdlind(500,'1','EXC','STATE')
#    ierr, inv_var_id = psspy.mdlind(500,'1','GEN','VAR')
#    
#    psspy.var_channel([-1,ppc_var_id+17],'PPC_LVRT' ) #PPC lvrt 
#    psspy.var_channel([-1,ppc_var_id+18],'PPC_HVRT' ) #PPC lvrt flag
#    psspy.var_channel([-1,inv_var_id+186], 'INV_LVRT') #INV Lvrt flag
#    psspy.var_channel([-1,inv_var_id+187], 'INV_HVRT') #INV Lvrt flag
#    
#    psspy.var_channel([-1,inv_var_id+102], 'INV_SPIKE_SUP') #INV Lvrt flag
#    psspy.var_channel([-1,inv_var_id+163], 'INV_FRT_STATE') #INV Lvrt flag
#    
#    psspy.machine_array_channel([-1,5,500],'1','Q_CMD_TO_INV')
#    psspy.machine_array_channel([-1,8,500],'1','P_CMD_TO_INV')
#    
#    psspy.state_channel([-1,ppc_state_id+10],'F_FILTER_OUT' ) # frequency filter output at PPC
    
#    for chan_if in range (0,199):
#        psspy.var_channel([-1,ppc_var_id])
    
    
    print('channels added')


def tune_parameters(PSSEmodelDict):
    # Only activated in tunning process. Should be commented out when finishing this process
#    busgen1 = 9942
#    busgen2 = 9944
#    pccmodel = r"""EMSPCI2_1"""
#    invmodel = r"""ING1BI2_1"""
#    psspy.change_wnmod_con(busgen1,r"""1""",invmodel,83,0)  
#    psspy.change_wnmod_con(busgen2,r"""1""",invmodel,83,0) 
    pass
    
    
def save_case_file(testname):
    global standalone_script   
    dirname, flnm = os.path.split(psspy.sfiles()[0].rstrip(".sav"))
    psspy.save(dirname+"\\"+flnm+'_'+str(testname)+".sav")
    
#    psspy.save(dirname+"\\"+"loadflow_after_init.sav")
    
#creates output directory to save results in 
#def create_output_dir(subdir, set_id): 
#    try:
#        os.mkdir(subdir+"\\FT_"+str(set_id))
#    except OSError:
#        print("Creation of the directory %s failed" % subdir+"\\FT_"+str(set_id))
#    else:
#        print("Successfully created the directory %s " % subdir+"\\FT_"+str(set_id))


def save_results(MYOUTFILE, csv_file):
    outfile=dyntools.CHNF(MYOUTFILE)
    short_title, chanName_dict, chandata_dict = outfile.get_data()
    
    dataDict={}
    
    df_time = pd.Series(chandata_dict['time'], name = chanName_dict['time'])
    df=df_time
    for i in sorted(chanName_dict.keys()):
        if i !="time":
            dataDict[chanName_dict[i]]=chandata_dict[i]
            temp = pd.Series(chandata_dict[i], name = chanName_dict[i])
            df=pd.concat([df,temp],axis=1)
            
    dirname, flnm=os.path.split(psspy.sfiles()[0].rstrip(".sav"))
    
    df.to_csv("{}.csv".format(csv_file), sep=',', index=False)
    
    
#def save_results(MYOUTFILE):
#    outfile=dyntools.CHNF(MYOUTFILE)
#    short_title, chanName_dict, chandata_dict = outfile.get_data()
#    
#    dataDict={}
#    
#    df_time = pd.Series(chandata_dict['time'], name = chanName_dict['time'])
#    df=df_time
#    for i in sorted(chanName_dict.keys()):
#        if i !="time":
#            dataDict[chanName_dict[i]]=chandata_dict[i]
#            temp = pd.Series(chandata_dict[i], name = chanName_dict[i])
#            df=pd.concat([df,temp],axis=1)
#            
#    dirname, flnm=os.path.split(psspy.sfiles()[0].rstrip(".sav"))
#    csv_filename=dirname+"\\FT_"+str(set_id)+"\\"+'FT_'+flnm+'_set_'+str(set_id)
#    df.to_csv("{}.csv".format(MYOUTFILE), sep=',', index=False)
  
    
#address any network bugs This is project_specific and should normally be done in the sav. files before using this script. 
    
def implement_droop_LF(droop_value, droop_base, vol_deadband, V_POC, Q_POC):
#    droop_value = 3.325
#    droop_base = 31.6
#    V_POC = setpoint['V_POC']
#    Q_actual = setpoint['Q']
#    vol_deadband = 0
    
    deltaV_comp = (Q_POC/droop_base)*(droop_value/100)
    if Q_POC<0:
        deltaV =  deltaV_comp + vol_deadband
    else: 
        deltaV =  deltaV_comp - vol_deadband
    Vspnt = deltaV + V_POC
    return Vspnt    
    
def precondition_network():

    # From Patrick
    # fault_studies.py -> initialisation
    # Reason: These machine does not have the model, and they dont deliver power in Load Flow -> causing error message in initialisation
#    for gen in [30711, 30712, 30713, 30718, 30716, 30720, 30721, 30722, 
#                30724, 30725, 30726, 30728, 30730, 30732, 30733, 30734, 
#                30735, 30736, 30737, 30740, 30741, 30743, 30744, 30745]:
#        psspy.machine_chng_2(gen,r"""1""",[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#
#    
#    # Error: POC active power out of limits: P_POC > Maximum P POC 
##    psspy.machine_chng_2(367176,r"""1""",[_i,_i,_i,_i,_i,_i],[ 40.55,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
##    psspy.machine_chng_2(367179,r"""1""",[_i,_i,_i,_i,_i,_i],[ 40.55,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#    # psspy.fnsl([1,0,0,1,1,0,99,0])
#
##     psspy.machine_chng_2(99903,r"""1""",[_i,_i,_i,_i,_i,_i],[ 77.6,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
##     psspy.fnsl([1,0,0,1,1,0,99,0])
#    psspy.machine_chng_2(99903,r"""1""",[_i,_i,_i,_i,_i,_i],[ 77.1,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#    psspy.fnsl([1,0,0,1,1,0,99,0])
#    psspy.fnsl([1,0,0,1,1,0,99,0])
    pass
        
#    return 0
    
#def save_test_description(testInfoDir, scenario, scenario_params, setpoint_params, ProfilesDict):
#    try:
#        os.mkdir(testInfoDir)
#    except OSError:
#        print("Creation of the directory %s failed" % testInfoDir)
#    else:
#        print("Successfully created the directory %s " % testInfoDir)
#    testInfo=shelve.open(testInfoDir+"\\"+str(scenario), protocol=2)
#    testInfo['scenario']=scenario
#    testInfo['scenario_params']=scenario_params
#    testInfo['setpoint']=setpoint_params
#    if('Test profile' in scenario_params.keys()):
#        if( (scenario_params['Test profile']!= None) and (scenario_params['Test profile']!='') ):
#            testInfo['profile']=ProfilesDict[scenario_params['Test profile']]
#    testInfo.close() 
#    pass





#    #create list of generator components to be initialised using the setpoint info keys
#    gen_list={}
#    for key in setpoint.keys():
#        if ('P_' in key):
#            gen_name=key[2:]
#            loc_ID=setpoint['LOC_'+gen_name]
#            frombus=PSSEmodelDict[loc_ID[0:2]+'fromBus'+loc_ID[2:]]
#            tobus=PSSEmodelDict[loc_ID[0:2]+'toBus'+loc_ID[2:]]
#            gen_list[gen_name]={'P':setpoint[key], 'Q':setpoint['Q_'+gen_name], 'fromBus':frombus, 'toBus':tobus, 'genBus':setpoint['BUS_'+gen_name]}
#            
#    pass
#
#    #disconnect offline machines
#    offline_machines=[]
#    if('offline_machines' in setpoint.keys()):
#        if( (setpoint['offline_machines']!=None) and(setpoint['offline_machines']!='') ):
#            offline_machines=ast.literal_eval(setpoint['offline_machines']) 
#
#    for gen in  gen_list.keys():
#        if(not (gen_list[gen]['genBus'] in offline_machines)):
#            #set output to  and switch gen on
##            psspy.plant_data_3(gen_list[gen]['genBus'],0,_i,[ setpoint['V_POC'],_f]) #Update the setpoint voltage of the generator - may not needed as the Q value is fixed with the loop below
#            ############################################################    
#            # 01/9/2022: Initialise droop characteristic - Lancaster only: Update voltage setpoint base on the actual voltage and Q at POC
#            droop_value = 3.325 #%
#            droop_base = 31.6
#            vol_deadband = 0  
#            Vspnt = implement_droop_LF(droop_value, droop_base, vol_deadband, setpoint['V_POC'], setpoint['Q'])
#            psspy.plant_data_3(gen_list[gen]['genBus'],0,_i,[ Vspnt,_f]) #Update the setpoint voltage of the generator - for setpoint variable initialisation
#            ############################################################ 


def checkBusNum(faultbus):
    ibus = faultbus
    flag = True
    while flag:
        ierr = psspy.busexs(ibus) 
        if ierr == 0: #if ierr = 0: bus found
            ibus -= 1 # change faultbus and continue the loop
        else: # if bus not found then it can be used
            faultbus = ibus
            flag = False
    return faultbus

    
###############################################################################
# Network Fault simulation routine
###############################################################################


def main():
    global main_folder
    global ModelCopyDir
    global ResultsDir
    global active_faults
    global reset_index
    global standalone_script
#        prompt user to select batch name (date and case name as default)   
#        print("Batch name (will default to current date):")
#        batchName = sys.stdin.readline()
#    batchName=str(datetime.datetime.now().strftime("%Y%m%d-%H%M"))+"_Network_faults"
#    try: 
#        batchName=str(input("Batch name (will default to current date):"))
#    except: 
#        batchName=str(datetime.datetime.now().strftime("%Y%m%d-%H%M"))+"_Network_faults - "+str(ProjectDetailsDict['NameShrt'])
    print(testRun)
    try: os.mkdir(ModelCopyDir+"\\"+testRun)
    except: ("'"+ModelCopyDir+"\\"+testRun+"' already exists")
    NEM_models=str(SimulationSettingsDict['model(s) for PSSE network studies']).split(",")
    for NEM_model in NEM_models:  
        try: os.mkdir(ModelCopyDir+"\\"+testRun+"\\"+NEM_model)
        except: ("'"+ModelCopyDir+"\\"+testRun+"\\"+NEM_model+"' already exists")
        
        scenarios=NetworkScenDict.keys()
        scenarios.sort(key = lambda x: x[3:] )
        
#        active_scenarios=[]
#        for scenario in scenarios:
#            if(NetworkScenDict[scenario][0]['run in PSS/E?']=="yes"):
#                active_scenarios.append(scenario)
                
        active_scenarios=[]
        for scenario in scenarios:
            if(NetworkScenDict[scenario][0]['run in PSS/E?']=='yes'):
                if('simulation batch' in NetworkScenDict[scenario][0].keys()):
                    if( (simulation_batches==[]) or (NetworkScenDict[scenario][0]['simulation batch'] in simulation_batches) ):
                        active_scenarios.append(scenario)
                else:
                    active_scenarios.append(scenario)
                
        for scenario in active_scenarios:
            event_list=NetworkScenDict[scenario]
#            try: os.mkdir(ModelCopyDir+"\\"+batchName+"\\"+NEM_model+"\\"+scenario)
#            except: ("'"+ModelCopyDir+"\\"+batchName+"\\"+NEM_model+"\\"+scenario+"' already exists")
            #create copy of model folder
            target_dir=ModelCopyDir+"\\"+testRun+"\\"+NEM_model+"\\"+scenario
            try:
                shutil.copytree(main_folder+"\\base_model\\"+NEM_model, target_dir)
            except OSError:
               print("Creation of the directory %s failed" % target_dir)
            else:
               print("Successfully created the directory %s" % target_dir)
               
               
            filenames = next(os.walk(target_dir), (None, None, []))[2]
            dll_files=[]
            for filename in filenames:
                if len(filename)>4:
#                    if(filename[-4:]==".sav"):
                    if(filename[-4:]==".sav") and (str(scenario) not in filename):
                        sav_file=filename
                    elif(filename[-4:]==".dll"):
                        dll_files.append(filename)
                    elif(filename[-4:]==".dyr"):
                        dyr_file=filename
                    
#           run network fault simulation on that copy, using channels specified in "ModelDetailsPSSE" and channels from Bus Lib and Branch Lib            
            active_faults = {} #dict to store ids of all faults. This is especially useful for multi-fault scenarios. internal PSS/E ID's may differ.
                            # The keys are the fault IDs in this script, the values are the PSS/E fautl IDs            
            reset_index = 1
            print(datetime.datetime.fromtimestamp(time.time()))
            #print('simulating fault set '+str(set_id))
            
            global breakers
            breakers = []
            global event_queue
            event_queue=[]
#            dirname, flnm=os.path.split(psspy.sfiles()[0].rstrip(".sav"))
            psspy.case(target_dir+'\\'+sav_file)
            psspy.fnsl([1,0,0,1,1,0,0,0])

            precondition_network() # for debugging purposes only            
            
#            initialise_network_and_build_event_queue(fault_list, breakers, target_dir, sav_file)    
            save_case_file(scenario) #save network for test purposes (to allow to see breakers that have been added)
#            standalone_script+="psspy.case('loadflow_after_init.sav')\n"
            standalone_script+="psspy.case('"+sav_file.rstrip('.sav')+"_"+scenario+".sav')\n"
            standalone_script+='psspy.base_frequency(50.0)\n'
            
#            psspy.startrecording(1,target_dir+'\\'+sav_file.rstrip(".sav")+'_'+scenario+"_Recorded"+".py")
            init_dynamics(target_dir, dyr_file, dll_files)
            add_channels()
            

#            stdout=open(target_dir+'\\'+sav_file.rstrip(".sav")+'_'+scenario+'.log','w')
#            with silence(stdout):

#            psspy.progress_output(2,os.path.join(workspace_folder, OUTPUT_name)[0:-4]+".LOG",[0,0])
#            standalone_script+="psspy.progress_output(2,'"+str(OUTPUT_name[0:-4]+".LOG")+"',[0,0])\n"
    
            psspy.progress_output(2,os.path.join(target_dir, sav_file)[0:-4]+'_'+str(scenario)+".LOG",[0,0])
            standalone_script+="psspy.progress_output(2,'"+str(sav_file[0:-4]+'_'+str(scenario)+".LOG")+"',[0,0])\n"
    
#                psspy.startrecording(1,target_dir+'\\'+sav_file.rstrip(".sav")+'_'+scenario+"_Script"+".py")
 
            MYOUTFILE=init_simulation()
            
#                psspy.startrecording(1,target_dir+'\\'+sav_file.rstrip(".sav")+'_'+scenario+"_Script"+".py")
            initialise_network_and_build_event_queue(event_list, breakers, target_dir, sav_file) 

            event_queue = order_event_queue(event_queue)
            final_event_time = event_queue[-1]['time']
            while (event_queue!=[]):
                event_queue = execute_event(event_queue) #execute next event in the queue

            # If it is fault test, run an additional 5secs
            if event_list[-1]['Test Type'] in ["Fault", "Multi_fault", "Switching", "Tx_tap_profile"]:
                psspy.run(0, final_event_time+5, 100,1,0)
                standalone_script+="psspy.run(0,"+str(final_event_time+5)+", 100, 1, 0)\n"

##                psspy.stoprecording()
                
#            stdout.close()

            text_file = open(target_dir+'\\'+sav_file.rstrip(".sav")+'_'+str(scenario)+".py", "w")
            text_file.write(standalone_script)
            text_file.close()            
            
            OutputDir =ResultsDir
            testRun_ = testRun
            csv_file=OutputDir+"\\"+testRun_+"\\"+NEM_model+"\\"+scenario+"\\"+scenario+'_results'
            try:
                    os.mkdir(OutputDir+"\\"+testRun_)
            except:
                print("testRun result folder already exists")
            else:
                print("testRun directory created")
            try:
                    os.mkdir(OutputDir+"\\"+testRun_+"\\"+NEM_model)
            except:
                print("NEM_model result folder already exists")
            else:
                print("NEM_model directory created")                
            try:
                    os.mkdir(OutputDir+"\\"+testRun_+"\\"+NEM_model+"\\"+scenario)
            except:
                print("scenario folder already exists")
            else:
                print("scenario results folder created")
        
            save_results(MYOUTFILE, csv_file) # save csv files for every run in properly organised folder structure. Save metadata about test itself in folder structure along with the result data
            scenario_params = event_list
            save_test_description(OutputDir+"\\"+testRun_+"\\"+NEM_model+"\\"+scenario+"\\testInfo", scenario, scenario_params)
#            save_test_description(OutputDir+"\\"+testRun_+"\\"+scenario+"\\testInfo", scenario, fault_list, SetpointsDict[scenario_params['setpoint ID']], ProfilesDict)# write test metadata to human-readable txt or csv file in same folder as results.
            
            import gc
            gc.collect()
#            del my_array
#            gc.collect()
            
#save_result_summary()

if __name__ == "__main__":
    ###########################################################################
    #USER INPUT
    ###########################################################################    
    main()

    
    
    
                                
                            
                                
                                
                        
                        
        

            
            
                

    
    
                            
        