# -*- coding: utf-8 -*-
"""
Created on Tue Feb 15 09:08:16 2022

@author: ESCO

FUNCTIONALYTY:
The script will run a steady state analysis, using the study inputs provided in the TestInfor spreadsheet

COMMENTS:
        The script reads inputs from common excel spreadsheet located in folder: test_scenario_definitions. 
            + Contingency scenarios are defined in tab: SteadyStateStudies
            + Buses and Branches to be monitored are definded in tab: MonitorBuses and MonitorBranches
        The script then runs through differnt contingency scenarios, does the load flow analysis. Results return for each snapshot:
            + Voltage Level: voltage level at interested (monitored) buses
            + Line Loadings: line loading in base case and contingencies
            

@NOTE: 
        if the script is located on sharepoint folder, it will create an equivalent folder path locally for storing results -> reduce syncing burden
        In the initialised funtion e.g. init_gens_vdc, power flow is considered reversed from POC, so ibus will be from the SF side.

"""
# import sys
import os, sys
import getpass
import pandas as pd
import numpy as np
from datetime import datetime
import time
from contextlib import contextmanager
from win32com.client import Dispatch
timestr=time.strftime("%Y%m%d-%H%M%S")
#start_time = datetime.now()

###############################################################################
#USER INPUTS
###############################################################################

TestDefinitionSheet = r'20230828_SUM_TESTINFO_V1.xlsx'

# allign the input model name between the main analysis and reporting scripts
input_models={
                'HighLoad':{'on':'HighLoad_genon\\SUMSF_high_genon.sav',
                           'off':'HighLoad_genoff\\SUMSF_high_genoff.sav'},
                'LowLoad': {'on':'LowLoad_genon\\SUMSF_low_genon.sav',
                            'off':'LowLoad_genoff\\SUMSF_low_genoff.sav'},
             }
                
'''
input_models={
                'HighLoad':{'on':'HighLoad\\SUMSF_high_genon.sav',
                           'off':'HighLoad\\SUMSF_high_genoff.sav'},
                'LowLoad': {'on':'LowLoad\\SUMSF_low_genon.sav',
                            'off':'LowLoad\\SUMSF_low_genoff.sav'},
             }
'''

# define the name for data frame and associated excel sheet output
df_to_sheet = {'volt_levels':{'df':'volt_levels', 'sht':'Volt Levels'},
                'line_loadings':{'df':'line_loadings', 'sht':'Line Loadings'},
                'volt_fluc_gen_chng':{'df':'volt_fluc_gen_chng', 'sht':'Volt Fluc GenChg'},
                'volt_fluc_lol':{'df':'volt_fluc_lol', 'sht':'Volt Fluc Lol'},
                'fault_levels':{'df':'fault_levels', 'sht':'Fault Levels'},
               }



simulation_batches=['S52512_SS']

try: testRun = timestr + '_' + simulation_batches[0] #define a test name for the batch or configuration that is being tested -> link to time stamp for auto update
except: testRun = timestr
    
# Different control mode if applicable. 
gens_with_pf = {}
gens_with_vdc = {}
gens_with_pf_vc = {}
statcom = {}

# generators with PF control mode
gens_with_pf = {'gens':[
                         {'gen_bus':1,'gen_id':'1','poc_bus':401,'ibus':201,'poc_pf':-0.98,'poc_q_max':23.7,'poc_p_gen':49.5}, #Argyle SF

                         ]}
#
# generators with voltage droop control mode - note that the power flow is measured from POC back to SF and then get the negative value

gens_with_vdc = {'gens':[ # One generator controlling POC voltage
#                            {'gen_bus':[273501, 273502],'gen_id':['1','1'],'poc_bus':273540,'ibus':273541,'poc_v_spt':1.02,'poc_v_dbn':0.0,'poc_q_max':35.55,'poc_drp_pct':5.0165,'gen_q_max':[60.48,60.48]},
                            {'gen_bus':2000,'gen_id':'1','poc_bus':5400,'ibus':5300,'poc_v_spt':0.99,'poc_v_dbn':0.0,'poc_q_max':45.425,'poc_drp_pct':5.0165,'gen_q_max':81.0}, #Mez SF
                            {'gen_bus':2500,'gen_id':'1','poc_bus':2200,'ibus':2300,'poc_v_spt':1.0,'poc_v_dbn':0.0,'poc_q_max':43.45,'poc_drp_pct':5.0165,'gen_q_max':129.44}, #Gunnedah SF
                            {'gen_bus':100,'gen_id':'1','poc_bus':106,'ibus':105,'poc_v_spt':1.0,'poc_v_dbn':0.0,'poc_q_max':25.675,'poc_drp_pct':5.0165,'gen_q_max':33.0}, # Tamworth SF

                        ],
    
                'gens2':[ # Two generators controlling voltage at same POC point
                            {'gen_bus':[273501, 273502],'gen_id':['1','1'],'poc_bus':273540,'ibus':273541,'poc_v_spt':1.02,'poc_v_dbn':0.0,'poc_q_max':35.55,'poc_drp_pct':5.0165,'gen_q_max':[60.48,45.36],'gen_q_pct':[0.5,0.5]},
                            
                        ]
                }

## generators with hybrid control mode (PF and direct voltage control)
#gens_with_pf_vc = {'gens':[{'gen_bus':302,'gen_id':'1','poc_bus':1,'ibus':102,'poc_pf':-0.995,'poc_q_max':30.0276,'poc_p_gen':76.0,'gen_q_max':55.0},
##                         {'gen_bus':30714,'gen_id':'1','poc_bus':36711,'ibus':36712,'poc_pf':-0.96,'poc_q_max':21.1857,'poc_p_gen':66.00,'gen_q_max':21.1857},
##                         {'gen_bus':30715,'gen_id':'1','poc_bus':36711,'ibus':36713,'poc_pf':-0.96,'poc_q_max':17.6143,'poc_p_gen':34.00,'gen_q_max':17.6143},
#                         
#                         #{'gen_bus':100001,'gen_id':'1','poc_bus':37703,'ibus':100001,'poc_pf':0.0,'gen_q_max':3.5,'gen_p_gen':0.0}
#                         #{'gen_bus':367169,'gen_id':'1','poc_bus':36721,'ibus':367167,'poc_pf':-0.985,'gen_q_max':15.8,'gen_p_gen':-40},
#                         #{'gen_bus':367166,'gen_id':'1','poc_bus':36721,'ibus':367164,'poc_pf':-0.985,'gen_q_max':15.8,'gen_p_gen':-40}
#                         ]}
## Statcom
#statcom =         {'gens':[{'gen_bus':100001,'gen_id':'1','poc_bus':37703,'ibus':100001,'gen_q_max':3.5,'gen_q_min':-3.5,'gen_q_ini':0.0}]}

###############################################################################
# Supporting functions
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

def make_dir(dir_path, dir_name=""):
    dir_make = os.path.join(dir_path, dir_name)
    try:
        os.mkdir(dir_make)
    except OSError:
        pass
    return dir_make
        
def createPath(main_folder_out):
    path = os.path.normpath(main_folder_out)
    path_splits = path.split(os.sep) # Get the components of the path
    child_folder = r""+ path_splits[0] # Build up the output path from base drive
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
                for bus_no, bus_name in zip(mntd_buses[0],mntd_buses[1]):
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
                for bus_no, bus_name in zip(mntd_buses[0],mntd_buses[1]):
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
    for ii,jj,kk in zip(mntd_brches[0],mntd_brches[1],mntd_brches[2]):
        info = af.get_branch_info(ii,jj,kk)
        brch_info.append(info)
    brch_df1 = brch_df.append(pd.DataFrame(data=brch_info))
    brch_df1 = brch_df1[0].apply(pd.Series)
    brch_df = brch_df.append(brch_df1)
    brch_df['brch_name']= mntd_brches[3]
    #bus results
    bus_df['bus_name'] = []
    for bus_no,bus_name in zip(mntd_buses[0],mntd_buses[1]):
        temp_bus_info = af.get_bus_info(bus_no,['TYPE','PU','ANGLED'])
        temp_bus_info[bus_no].update({'bus_name':bus_name})
        bus_df1 = pd.DataFrame.from_dict(temp_bus_info,orient = 'index')
        bus_df = bus_df.append(bus_df1)
         

    return brch_df, bus_df

def execute_event(event, applied_status = 'ini_status', applied_value = 'ini_value'):
    global standalone_script
    if (event['TestType']=='Trip'): # if trip event, donot consider the input applied_value
        if(event['Element']=='Line'): # If fault elememt is a line
            psspy.branch_chng_3(event['i_bus'],event['j_bus'],str(event['id']),[event[applied_status],_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
            standalone_script+="psspy.branch_chng_3("+str(event['i_bus'])+","+str(event['j_bus'])+",'"+str(event['id'])+"',["+str(event[applied_status])+",_i,_i,_i,_i,_i],"+"[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)\n"
        elif(event['Element']=='Shunt'): # If fault elememt is a Shunt
            psspy.shunt_chng(event['i_bus'],str(event['id']),event[applied_status],[_f,_f])
            standalone_script+="psspy.shunt_chng("+str(event['i_bus'])+",'"+str(event['id'])+"',"+str(event[applied_status])+","+"[_f,_f])\n"
        elif(event['Element']=='Machine'): # If fault elememt is a Machine
            psspy.machine_chng_2(event['i_bus'],str(event['id']),[event[applied_status],1,0,0,0,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            standalone_script+="psspy.machine_chng_2("+str(event['i_bus'])+",'"+str(event['id'])+"',"+str(event[applied_status])+",1,0,0,0,0],"+"[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
        elif(event['Element']=='Tx_2w'): # If fault elememt is a two windig transformer
            psspy.two_winding_chng_6(event['i_bus'],event['j_bus'],str(event['id']),[event[applied_status],_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
            standalone_script+="psspy.two_winding_chng_6("+str(event['i_bus'])+","+str(event['j_bus'])+",'"+str(event['id'])+"',"+str(event[applied_status])+",_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],"+"[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s\n"
        elif(event['Element']=='Bus'):
            if event[applied_status] == 0:
                psspy.dscn(event['i_bus'])
                standalone_script+="psspy.dscn("+str(event['i_bus'])+")\n"
            else:
                psspy.recn(event['i_bus'])
                standalone_script+="psspy.recn("+str(event['i_bus'])+")\n"
                
    if (event['TestType']=='ChgMW'):
        if(event['Element']=='Machine'):
            if event[applied_value] != '': # only change if the applied_value is provided 
                psspy.machine_chng_2(event['i_bus'],str(event['id']),[_i,1,0,0,0,0],[event[applied_value],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #if status is 0, trhen machine will be  switched off.
                standalone_script+="psspy.machine_chng_2("+str(event['i_bus'])+",'"+str(event['id'])+"',"+"[_i,1,0,0,0,0],"+"["+str(event[applied_value])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
        elif(event['Element']=='Load'):
            if event[applied_value] != '': # only change if the applied_value is provided 
                psspy.load_chng_5(event['i_bus'],str(event['id']),[_i,_i,_i,_i,_i,_i,_i],[event[applied_value],_f,_f,_f,_f,_f,_f,_f])
                standalone_script+="psspy.load_chng_5("+str(event['i_bus'])+",'"+str(event['id'])+"',"+"[_i,_i,_i,_i,_i,_i,_i],"+"["+str(event[applied_value])+",_f,_f,_f,_f,_f,_f,_f])\n"

    if (event['TestType']=='ChgTap'): #change tap of the transoformer
        pass
    if (event['TestType']=='ChgVol'): #change voltage setpoint of the machine 
        pass
    if (event['TestType']=='ChgMVAr'): #change Reactive power setpoint of a machine or Statcom    
        pass
        

                
def pre_event_condition(sav_file,case_num):
    global statcom, gens_with_pf, gens_with_pf_vc, gens_with_vdc, standalone_script
    # Initialise Statcom:
    if statcom != {}: 
        err_code, auto_script = mf.ini_statcom(statcom,err_code = 0,auto_script = '')
        if err_code == 1:  #if load flow does not converge, note it down and raise an error
            error_file.write(str(sav_file) + ' - Case: '+str(case_num) + ' - Load flow is not converged when initialising the statcom')
            error_file.close() # Save the note
            raise
        else:
    #        standalone_script+=auto_script # for debuging purpose
            pass
    
    # Initialise the gens with Power factor control mode
    if gens_with_pf != {}: 
        err_code, auto_script = mf.init_gens_pf(gens_with_pf,err_code = 0,auto_script = '')
        if err_code == 1:  #if load flow does not converge, note it down and raise an error
            error_file.write(str(sav_file) + ' - Case: '+str(case_num) + ' - Load flow is not converged when initialising power factor control mode')
            error_file.close() # Save the note
            raise
        else:
    #        standalone_script+=auto_script # for debuging purpose
            pass
    
    # Initialise the gens with Voltage droop control mode
    if gens_with_vdc != {}: 
        err_code, auto_script = mf.init_gens_vdc(gens_with_vdc,err_code = 0,auto_script = '')
        if err_code == 1:  #if load flow does not converge, note it down and raise an error
            error_file.write(str(sav_file) + ' - Case: '+str(case_num) + ' - Load flow is not converged when initialising voltage droop control mode')
            error_file.close() # Save the note
            raise
        else:
            standalone_script+=auto_script # for debuging purpose
            pass
        
def post_event_condition(sav_file,case_num):
    global statcom, gens_with_pf, gens_with_pf_vc, gens_with_vdc, standalone_script
    
    # Statcom
    if statcom != {}: 
        err_code, auto_script = mf.lckd_statcom(statcom)
        if err_code == 1:  #if load flow does not converge, note it down and raise an error
            error_file.write(str(sav_file) + ' - Case: '+str(case_num) + ' - Load flow is not converged when updating the statcom at post event')
            error_file.close() # Save the note
            raise
        else:
            standalone_script+=auto_script # for debuging purpose
            pass

    # Hybrid control gens
    if gens_with_pf_vc != {}: err_code, auto_script = mf.lckd_gens_hc(gens_with_pf_vc)
    
    # Voltage droop control gens
    if gens_with_vdc != {}: 
        err_code, auto_script = mf.lckd_gens_vdc(gens_with_vdc)
        if err_code == 1:  #if load flow does not converge, note it down and raise an error
            error_file.write(str(sav_file) + ' - Case: '+str(case_num) + ' - Load flow is not converged when update voltage droop control mode at post event')
            error_file.close() # Save the note
            raise
        else:
            standalone_script+=auto_script # for debuging purpose
            pass
#    if gens_with_pf != {}: mf.lckd_gens_pf(gens_with_pf) # if pf control is in use, pq flow does not change at contingency
  
    
def copy_file(ori_file, target_folder):
    try:
        shutil.copy2(ori_file, target_folder)
    except OSError:
       print("Creation of the directory %s failed" % target_folder)
    else:
       print("Successfully created the directory %s" % target_folder)
    
    file_name = os.path.basename(ori_file) # name of the file coppied
    
    return file_name
    
def run_scenarios(sav_file, active_scenarios):
    global ModelCopyDir, standalone_script, sim_log, error_file
    sim_log=''
    #generatorStatus=check_gen_status(gen_bus) #function that checks generator status at given bus and returns 1 or 0.
    
    #Create dataframe to store voltage levels, Contingency results, Generation Changes results, fault levels
    vltg_lvls_df = pd.DataFrame()
    line_ldng_df = pd.DataFrame()
    bus_df_lol = pd.DataFrame()
    brch_df_lol = pd.DataFrame()
    bus_df_Pchange = pd.DataFrame()
    brch_df_Pchange = pd.DataFrame()
    fault_df = pd.DataFrame()
    
    ContRating = 'RATING1' #'RATING2'
    
    # Load the base case, apply the initialised condition, solve load flow and save it as the base results - Normal State
    with silence():

        # Load the base case implement initial condition and save it as before genchange
        psspy.case(base_model +"\\"+ sav_file) # load the case
        sav_file_name = os.path.basename(base_model +"\\"+ sav_file) # get the file name
        case_num = "basecase"
        pre_event_condition(sav_file,case_num) # voltage droop or power factor control initialisation if applicable
        wrkg_spc = createPath(ModelCopyDir+"\\"+testRun+"\\"+'Normal State') # working space folder
        psspy.save(wrkg_spc +"\\" +sav_file_name) #save the case with intial conditions
#        standalone_script+="psspy.save(\\"+str('Normal State')+"\\"+str(sav_file_name)+")\n"
            
        # Monitor variables for the base case
        brch_df0, bus_df0 = monitoring()
                                                   
        bus_df0['CaseNr']= "case0 (base)"
        bus_df0['VolDev (%)']= 0
        brch_df0['CaseNr']= "case0 (base)"
        brch_df0['Case_Code']= "Network normal"
        brch_df0['Loading (%)']= np.round((brch_df0['MVA'] / brch_df0[ContRating])*100,2)
        vltg_lvls_df = vltg_lvls_df.append(bus_df0)
        line_ldng_df = line_ldng_df.append(brch_df0)
#        bus_df_final = bus_df_final.append(bus_df0)
#        brch_df_final = brch_df_final.append(brch_df0)
        
    # Contingency cases 
    for case_num in active_scenarios: # run all the active case
        scenario=SteadyStateDict[case_num]
        sim_log+=case_num
        case_num_str = str(case_num)
        standalone_script = "# This script has been auto-generated by the PSSE Network test tool to allow for debugging of individual test cases.\n"
        
        # Fault Level caluclation. This requires only one model which is either given by the NSP or use the high load scenario with genon settings
        if(scenario[0]['TestType']=='FaultCal' and scenario[0]['Case_Code'] in sav_file): # No event cases - calculate Fault levels
            fault_results = fault_lvls(scenario)
            fault_df = fault_df.append(pd.DataFrame.from_dict(fault_results))

        else: # Event cases: need to execute all the events listed under the associated scenario 
            psspy.case(base_model +"\\"+ sav_file) # load the case - base case
            sav_file_name = os.path.basename(base_model +"\\"+ sav_file) # get the file name
            wrkg_spc = make_dir(ModelCopyDir + "\\" + testRun, dir_name=case_num_str) # working space folder
#            psspy.save(wrkg_spc +"\\" +  sav_file_name) # copy the base case over to the working space             #ONLY FOR DEBUGING THE INITIALISATION PROCESS
#            standalone_script+= "# Load the original case and apply pre-test conditions (if applicable).\n"        #ONLY FOR DEBUGING THE INITIALISATION PROCESS
#            standalone_script+="psspy.case('"+str(sav_file_name)+"')\n"                                            #ONLY FOR DEBUGING THE INITIALISATION PROCESS
            for i in range(0, len(scenario)): #loop through all the events in each scenario
                if scenario[i]['ini_value'] != '': #only execute the event if ini_value is provided
                    execute_event(scenario[i], applied_status = 'ini_status', applied_value = 'ini_value') # apply initial value for genchange cases
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
#                    if(af.test_convergence(method="fnsl", taps="step")>1):
                    mismatch=psspy.sysmsm()
                    if(mismatch>1):
                        error_file.write('\n' + str(sav_file)+ ' '+ str(case_num)+' System did not converge - when applying ini_value')# THROW ERROR IF MISMATCH >1
                        error_file.close() # Save the note
                        raise
            pre_event_condition(sav_file,case_num) # implement initialisation condition if applicable e.g. implement voltage droop control, pf control
            psspy.fnsl([1,0,0,1,1,0,0,0])
            psspy.fnsl([1,0,0,1,1,0,0,0])
            psspy.save(wrkg_spc +"\\" +  'Before_' +sav_file_name) #save the case with intial conditions - before event
#            standalone_script+="psspy.save('Before_"+str(sav_file_name)+"')\n"                                     #ONLY FOR DEBUGING THE INITIALISATION PROCESS
            standalone_script+= "# Load the case before contingencies....\n"
            standalone_script+="psspy.case('Before_"+str(sav_file_name)+"')\n"

            # Monitor variables BEFORE applying the event if needed
            brch_df_bf, bus_df_bf = monitoring()
            bus_df_bf['CaseNr']= case_num_str
            bus_df_bf['Case_Code']= scenario[0]['Case_Code']
            brch_df_bf['CaseNr']= case_num_str
            brch_df_bf['Case_Code']= scenario[0]['Case_Code']
            brch_df_bf['Loading_ini (%)']= np.round((brch_df_bf['MVA'] / brch_df_bf[ContRating])*100,2)

            # Apply the contingencies/genchanges, implement post event condition and save the case. This may need to adapt to different scenarios. Sometime solving the case right after applying the contingency make the network unconverged.
            standalone_script+= "# Apply the contingencies....\n"
            for i in range(0, len(scenario)):
                                                   
                execute_event(scenario[i], applied_status = 'end_status', applied_value = 'end_value') # apply end value for genchange cases

            psspy.fnsl([0,0,0,1,0,0,0,0])
            psspy.fnsl([0,0,0,1,0,0,0,0])
#            if(af.test_convergence(method="fnsl", taps="locked")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) +' System did not converge')# THROW ERROR IF MISMATCH >1
            mismatch=psspy.sysmsm()
            if(mismatch>1):
                error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) +' System did not converge - when executing the contingency/event')# THROW ERROR IF MISMATCH >1
                error_file.close() # Save the note
                raise

            post_event_condition(sav_file,case_num) # implement post event conditions after the contingency. In some cases, this need to be done right after applying the contingency, before solving the load flow with locked tap -> to avoid the not converging load flow
            psspy.fnsl([0,0,0,1,0,0,0,0])
            psspy.fnsl([0,0,0,1,0,0,0,0])
            
            psspy.save(wrkg_spc +"\\" +  'After_' +sav_file_name)  #save the case after the contingency with locked taps
                                                                                                         
            standalone_script+= "# Save the case after contingencies applied\n"
            standalone_script+="psspy.save('After_"+str(sav_file_name)+"')\n\n"
            text_file = open(ModelCopyDir+"\\"+testRun+"\\"+str(case_num_str)+'\\'+str(sav_file_name[0:-4])+".py", "w")
            text_file.write(standalone_script)
            text_file.close()

            # separate monitoring for contingencies and genchange
            brch_df_af, bus_df_af = monitoring() # monitor parameters after the event applied
            if (scenario[0]['TestType']=='Trip'): # if it is a contingency case update the voltage and loading data
                bus_df_Lupdate = bus_df_bf
                                          
                                                             
                bus_df_Lupdate['VolDev (%)']= np.round((bus_df_af['PU'] - bus_df_bf['PU'])*100,4) #calculates deviaton compared to voltage level before contingency
                bus_df_Lupdate['PU_final'] = bus_df_af['PU'] 
                bus_df_Lupdate['ANGLED_final'] = bus_df_af['ANGLED']
                bus_df_lol = bus_df_lol.append(bus_df_Lupdate)
                brch_df_Lupdate = brch_df_bf
                brch_df_Lupdate['Loading (%)']= np.round((brch_df_af['MVA'] / brch_df_af[ContRating])*100,2)
#                brch_df_lol = brch_df_lol.append(brch_df_Lupdate)
                line_ldng_df = line_ldng_df.append(brch_df_Lupdate)
                
            else: # if it is a setpoint change test: dont need to check the line loading as it would be less than full generation
                bus_df_Pupdate = bus_df_bf
                bus_df_Pupdate['VolDev (%)'] = np.round((bus_df_af['PU'] - bus_df_bf['PU'])*100,4)
                bus_df_Pupdate['PU_final'] = bus_df_af['PU'] 
                bus_df_Pupdate['ANGLED_final'] = bus_df_af['ANGLED']
                                                                                                                   
                bus_df_Pchange = bus_df_Pchange.append(bus_df_Pupdate)
#                brch_df_Pupdate = brch_df_bf
#                brch_df_Pupdate['Loading_final (%)'] = np.round((brch_df_af['MVA'] / brch_df_af[ContRating])*100,2)
#                brch_df_Pupdate['MVA_final'] = brch_df_af['MVA']
#                brch_df_Pupdate['P_final'] = brch_df_af['P']
#                brch_df_Pupdate['Q_final'] = brch_df_af['Q']
#                brch_df_Pupdate['PCTMVARATE_final'] = brch_df_af['PCTMVARATE']
#                brch_df_Pchange = brch_df_Pchange.append(brch_df_Pupdate)
    results={df_to_sheet['volt_levels']['df']:vltg_lvls_df, df_to_sheet['line_loadings']['df']:line_ldng_df, df_to_sheet['volt_fluc_gen_chng']['df']:bus_df_Pchange, df_to_sheet['volt_fluc_lol']['df']:bus_df_lol, df_to_sheet['fault_levels']['df']: fault_df }
    return results, sim_log

###############################################################################
# Define Project Paths
###############################################################################

overwrite = False # 
max_processes = 8 #set to the number of cores on my machine. Needs to be >= scenarioPerGroup --> increase for PSCAD machine
         
            
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
else: # if the output location is same as input location, then delete the links
    try:os.remove(main_folder + "\\model_copies.lnk")
    except: pass
    try: os.remove(main_folder + "\\result_data.lnk")
    except: pass
# Locating the existing folders
testDefinitionDir= os.path.abspath(os.path.join(main_folder, os.pardir))+"\\test_scenario_definitions"
base_model = main_folder+"\\base_model" #parent directory of the workspace folder
base_model_workspace = main_folder+"\\base_model\\SMIB" #path of the workspace folder, formerly "workspace_folder" --> in case the workspace is located in a subdirectory of the model folder (as is the case with MUL model for example)
zingen=main_folder+"\\zingen\\dsusr_zingen.dll"
libpath = os.path.abspath(main_folder) + "\\scripts\\Libs"
sys.path.append(libpath)
# print ("libpath = " + libpath)

# Directory to store Steady State/Dynamic result
ResultsDir = OutputDir+"\\steady_state"
make_dir(ResultsDir)

#outputResultPath = ''
#outputResultPath=ResultsDir+"\\"+testRun
#make_dir(outputResultPath)

###############################################################################
# Import additional functions, # Initialise PSSE
###############################################################################
import misc_functions as mf
import auxiliary_functions as af
import readtestinfo as readtestinfo
import psspy
import redirect
import shutil

sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
sys.path.append(sys_path_PSSE)
os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE
# Start PSSE
redirect.psse2py()
psspy.psseinit(10000)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()
redirect.psse2py()

###############################################################################
# Contingencies Infor
###############################################################################
# SteadyStateDict =  readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet)
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['SteadyStateStudies', 'ModelDetailsPSSE', 'MonitorBuses', 'MonitorBranches'])
SteadyStateDict = return_dict['SteadyStateStudies']
ModelDetailsPSSE = return_dict['ModelDetailsPSSE']
mntd_buses = return_dict['MonitorBuses']
mntd_brches = return_dict['MonitorBranches']
standalone_script = ''

# Create the folder to store analysis results
outputResultPath=ResultsDir+"\\"+testRun
make_dir(outputResultPath)
error_file = open(outputResultPath + '\\' + 'error_log.txt','a') #Text file to log errors
out_file_name = 'Steady State Analysis Results.xlsx' # for recording the results
writer = pd.ExcelWriter(outputResultPath+'\\'+out_file_name,engine = 'xlsxwriter')
        
def main():
    global ModelCopyDir, SteadyStateDict, input_models


    
    start_time = datetime.now()
    
    # Prepare active scenarios to be run:
    scenarios=SteadyStateDict.keys()
    scenarios.sort(key = lambda x: x[3:] )
    active_scenarios=[]
#    active_scenarios_Des = []
    for scenario in scenarios:
        if(SteadyStateDict[scenario][0]['run in PSS/E?']=='yes')or(SteadyStateDict[scenario][0]['run in PSS/E?']==1):
            active_scenarios.append(scenario)
    
    # loop through the active scenraios, run the analysis and record the results
    if len(active_scenarios) > 0:
#        # Create the folder to store analysis results
#        outputResultPath=ResultsDir+"\\"+testRun
#        make_dir(outputResultPath)
#        error_file = open(outputResultPath + '\\' + 'error_log.txt','a') #Text file to log errors
#        out_file_name = 'Steady State Analysis Results.xlsx' # for recording the results
#        writer = pd.ExcelWriter(outputResultPath+'\\'+out_file_name,engine = 'xlsxwriter')
    
        for snapshot_name in input_models.keys(): #loop the snapshots
            for config in ['on','off']: # two modes gen on and off
                #run case and get results
                results, sim_log=run_scenarios(input_models[snapshot_name][config], active_scenarios)
                
                # Export results:
                if not results[df_to_sheet['volt_levels']['df']].empty: results[df_to_sheet['volt_levels']['df']].to_excel(writer,sheet_name = df_to_sheet['volt_levels']['sht']+'_'+snapshot_name+'_'+config, index=True )
                if not results[df_to_sheet['line_loadings']['df']].empty: results[df_to_sheet['line_loadings']['df']].to_excel(writer,sheet_name = df_to_sheet['line_loadings']['sht']+'_'+snapshot_name+'_'+config, index=True )
                if not results[df_to_sheet['volt_fluc_gen_chng']['df']].empty: results[df_to_sheet['volt_fluc_gen_chng']['df']].to_excel(writer,sheet_name = df_to_sheet['volt_fluc_gen_chng']['sht']+'_'+ snapshot_name+'_'+config, index=True )
                if not results[df_to_sheet['volt_fluc_lol']['df']].empty: results[df_to_sheet['volt_fluc_lol']['df']].to_excel(writer,sheet_name = df_to_sheet['volt_fluc_lol']['sht']+'_'+ snapshot_name+'_'+config, index=True )
                if not results[df_to_sheet['fault_levels']['df']].empty: results[df_to_sheet['fault_levels']['df']].to_excel(writer,sheet_name = df_to_sheet['fault_levels']['sht']+'_'+ snapshot_name+'_'+config, index=True )

        writer.save()
        error_file_size = os.path.getsize(outputResultPath + '\\' + 'error_log.txt')
        if(error_file_size ==0):error_file.write('The system converged in all models and all scenarios')
        error_file.close()

    # Calculate time spent
    end_time = datetime.now()
    print('Duration: {}'.format(end_time - start_time))


if __name__ == '__main__':
    main()