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

## generators with PF control mode
#gens_with_pf = {'gens':[{'gen_bus':30714,'gen_id':'1','poc_bus':36711,'ibus':36712,'poc_pf':-0.96,'poc_q_max':21.1857,'poc_p_gen':66.00}, #NSF
#                         {'gen_bus':30715,'gen_id':'1','poc_bus':36711,'ibus':36713,'poc_pf':-0.96,'poc_q_max':17.6143,'poc_p_gen':34.00},
#                         {'gen_bus':30709,'gen_id':'1','poc_bus':36716,'ibus':36717,'poc_pf':-0.990001,'poc_q_max':19.25625,'poc_p_gen':48.75}, #WNSF
#                         {'gen_bus':30710,'gen_id':'1','poc_bus':36716,'ibus':36718,'poc_pf':-0.990001,'poc_q_max':10.36875,'poc_p_gen':26.25},
#                         {'gen_bus':302,'gen_id':'1','poc_bus':1,'ibus':102,'poc_pf':-0.995,'poc_q_max':30.0276,'poc_p_gen':76}, #GSF
#                         #{'gen_bus':367169,'gen_id':'1','poc_bus':36721,'ibus':367167,'poc_pf':-0.96,'gen_q_max':15.8,'gen_p_gen':-40},
#                         #{'gen_bus':367166,'gen_id':'1','poc_bus':36721,'ibus':367164,'poc_pf':-0.96,'gen_q_max':15.8,'gen_p_gen':-40}
#                         ]}
#
# generators with voltage droop control mode
#gens_with_vdc = {'gens':[{'gen_bus':9942,'gen_id':'1','poc_bus':9920,'ibus':9930,'poc_v_spt':1.02,'poc_v_dbn':0.0,'poc_q_max':35.5,'poc_drp_pct':5.02},
#                     ]}
                   
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
    for i,j,k in zip(mntd_brches[0],mntd_brches[1],mntd_brches[2]):
        info = af.get_branch_info(i,j,k)
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

def execute_event(event, applied_value = 'ini_value'):
    global standalone_script
    if (event['TestType']=='Trip'): # if trip event, donot consider the input applied_value
        if(event['Element']=='Line'): # If fault elememt is a line
            psspy.branch_chng_3(event['i_bus'],event['j_bus'],str(event['id']),[event['end_status'],_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
            standalone_script+="psspy.branch_chng_3("+str(event['i_bus'])+","+str(event['j_bus'])+",'"+str(event['id'])+"',["+str(event['end_status'])+",_i,_i,_i,_i,_i],"+"[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)\n"
        elif(event['Element']=='Shunt'): # If fault elememt is a Shunt
            psspy.shunt_chng(event['i_bus'],str(event['id']),event['end_status'],[_f,_f])
            standalone_script+="psspy.shunt_chng("+str(event['i_bus'])+",'"+str(event['id'])+"',"+str(event['end_status'])+","+"[_f,_f])\n"
        elif(event['Element']=='Machine'): # If fault elememt is a Machine
            psspy.machine_chng_2(event['i_bus'],str(event['id']),[event['end_status'],1,0,0,0,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            standalone_script+="psspy.machine_chng_2("+str(event['i_bus'])+",'"+str(event['id'])+"',"+str(event['end_status'])+",1,0,0,0,0],"+"[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
        elif(event['Element']=='Tx_2w'): # If fault elememt is a two windig transformer
            psspy.two_winding_chng_6(event['i_bus'],event['j_bus'],str(event['id']),[event['end_status'],_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
            standalone_script+="psspy.two_winding_chng_6("+str(event['i_bus'])+","+str(event['j_bus'])+",'"+str(event['id'])+"',"+str(event['end_status'])+",_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],"+"[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s\n"
        elif(event['Element']=='Bus'):
            psspy.dscn(event['i_bus'])
            standalone_script+="psspy.dscn("+str(event['i_bus'])+")\n"
            
    if (event['TestType']=='ChgMW'):
        if event[applied_value] != '': # only change if the applied_value is provided 
            psspy.machine_chng_2(event['i_bus'],str(event['id']),[_i,1,0,0,0,0],[event[applied_value],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #if status is 0, trhen machine will be  switched off.
            standalone_script+="psspy.machine_chng_2("+str(event['i_bus'])+",'"+str(event['id'])+"',"+"[_i,1,0,0,0,0],"+"["+str(event[applied_value])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"

    if (event['TestType']=='ChgTap'): #change tap of the transoformer
        pass
    if (event['TestType']=='ChgVol'): #change voltage setpoint of the machine 
        pass
    if (event['TestType']=='ChgMVAr'): #change Reactive power setpoint of a machine or Statcom    
        pass
        
# Initialise gens with voltage droop control mode:
def init_gens_vdc(gens_with_vdc): 
    global standalone_script
    if len(gens_with_vdc['gens']) != 0:
        for i in range(0,5): #If multiple generator participate in the QV droop, then repeat the process to make sure the actual voltage settle well
            for gen in gens_with_vdc['gens']:
                ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
                if ival == 0:
                    print('GEN is OFF')
                else:
                    poc_v_actual = af.get_bus_info(gen['poc_bus'],'PU')
                    poc_v_actual = poc_v_actual[gen['poc_bus']]['PU'] #poc volt level
                     
                    delta_v = poc_v_actual - gen['poc_v_spt']
                    vol_deadband = gen['poc_v_dbn']/2
                    if delta_v > vol_deadband:
                        delta_v = delta_v - vol_deadband #compensate for only part outside deadband/2
                    elif delta_v < -vol_deadband:
                        delta_v = delta_v + vol_deadband
                    else: delta_v = 0 # within the deadband
                    
                    q_poc_req = -((gen['poc_q_max']/(gen['poc_drp_pct']/100.0)) * delta_v) # Consider Qmax as the base for droop
                    
                    if q_poc_req > gen['poc_q_max']: # Limit the compensation to Qcorner at POC
                        q_poc_req = gen['poc_q_max']
                    if q_poc_req < -gen['poc_q_max']: 
                        q_poc_req = -gen['poc_q_max']
    
                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    q_poc = -s_poc.imag # poc q MVAr
                    delta_q = q_poc_req - q_poc
                     
                    tol_q = 1.0 #MVAr
                    k_factor = 0.55 # regression factor
                    iter_num = 15
                     
                    while abs(delta_q) > tol_q and iter_num > 0:
                         
                        ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                        q_gen_new = q_gen + delta_q*k_factor
                         
                        psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new,q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                        psspy.fnsl([1,0,0,1,1,0,0,0])
                        psspy.fnsl([1,0,0,1,1,0,0,0])
                        psspy.fnsl([1,0,0,1,1,0,0,0])
                        standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                        standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                        standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                        af.print_volts(lower=0.9, upper=1.1) #for debugging
#                        step_solv_paras()
                        if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
    
                        s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                        q_poc = -s_poc.imag # poc q MVAr
                        delta_q = q_poc_req - q_poc
                        iter_num -=1
                        
# Initialise gens with fix power factor control mode:
def init_gens_pf(gens_with_pf):
    global standalone_script
    if len(gens_with_pf['gens']) != 0: 
        for gen in gens_with_pf['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            if ival == 0:
                print('GEN is OFF')
            else:
                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #Update the Control Mode to Not a wind machine so Qmax, Qmin can be updated
                standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1] #power measured at POC, from POC back to the SF (reversed power flow)
                p_poc = -s_poc.real # poc p MW generated from SF
                q_poc = -s_poc.imag # poc q MVAr generated from SF
                q_poc_req = np.tan(np.arccos(abs(gen['poc_pf'])))*np.sign(gen['poc_pf']) * gen['poc_p_gen']
#                 q_poc_req = np.tan(np.arccos(abs(gen['poc_pf'])))*np.sign(gen['poc_pf']) * p_poc
                if q_poc_req > gen['poc_q_max']: # Limit the compensation to Qcorner at POC
                    q_poc_req = gen['poc_q_max']
                if q_poc_req < -gen['poc_q_max']: 
                    q_poc_req = -gen['poc_q_max']
                        
                delta_p = gen['poc_p_gen'] - p_poc
                delta_q = q_poc_req - q_poc

                tol_p = 0.5 #MW
                tol_q = 1.0 #MVAr
                k_factor = 0.25 # regression factor
                iter_num = 15
                 
#                while (abs(delta_p) > tol_p or abs(delta_q) > tol_q) and iter_num > 0:
                while abs(delta_q) > tol_q and iter_num > 0:
                    ierr,p_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'P') # gen p MW
                    ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                    p_gen_new = p_gen + delta_p*k_factor
                    q_gen_new = q_gen + delta_q*k_factor
                     
#                    if q_gen_new > gen['gen_q_max']: q_gen_new = gen['gen_q_max']
#                    if q_gen_new < -gen['gen_q_max']: q_gen_new = -gen['gen_q_max']
                         
                    psspy.machine_chng_2(int(gen['gen_bus']),r"""1""",[_i,_i,_i,_i,_i,_i],[p_gen_new,_f, q_gen_new, q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],["+str(p_gen_new)+",_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    af.print_volts(lower=0.9, upper=1.1) #for debugging
#                    step_solv_paras()
                    if(af.test_convergence(method='fnsl',taps='step')>1.0):
                        af.print_volts(lower=0.9, upper=1.1)
                        raise

                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    p_poc = -s_poc.real # poc p MW
                    q_poc = -s_poc.imag # poc q MVAr
#                     q_poc_req = np.tan(np.arccos(abs(gen['poc_pf'])))*np.sign(gen['poc_pf']) * p_poc
                    delta_p = gen['poc_p_gen'] - p_poc
                    delta_q = q_poc_req - q_poc
                    iter_num -=1

#                # set the voltage setpoint of the generator to actual voltage at POC -> this will not impact the PQ level in pf control; but will prepare for Vcontrol after contingency when plant in hygrid control
#                poc_v_actual = af.get_bus_info(gen['poc_bus'],'PU')
#                poc_v_spnt = poc_v_actual[gen['poc_bus']]['PU'] #poc volt level
#                psspy.plant_data_4(gen['gen_bus'],0,[_i,_i],[ poc_v_spnt,gen['poc_bus']])
                    
# Initialise gens with fix reactive power control mode:
def init_gens_qfix(gens_with_qfix):
    global standalone_script
    if len(gens_with_qfix['gens']) != 0: 
        for gen in gens_with_qfix['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            if ival == 0:
                print('GEN is OFF')
            else:
                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #Update the Control Mode to Not a wind machine so Qmax, Qmin can be updated
                s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1] #power measured at POC, from POC back to the SF (reversed power flow)
                q_poc = -s_poc.imag # poc q MVAr generated from SF
                q_poc_req = gen['poc_q_gen']
                delta_q = q_poc_req - q_poc
                tol_q = 1.0 #MVAr
                k_factor = 0.85 # regression factor
                iter_num = 15
                while abs(delta_q) > tol_q and iter_num > 0:
                    ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                    q_gen_new = q_gen + delta_q*k_factor
                    psspy.machine_chng_2(int(gen['gen_bus']),r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new, q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    af.print_volts(lower=0.9, upper=1.1) #for debugging
#                    step_solv_paras()
                    if(af.test_convergence(method='fnsl',taps='step')>1.0):
                        af.print_volts(lower=0.9, upper=1.1)
                        raise

                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    q_poc = -s_poc.imag # poc q MVAr
                    delta_q = q_poc_req - q_poc
                    iter_num -=1

# Initialise statcom - for LSF but can be used for other Statcorms
def ini_statcom(statcom):
    global standalone_script
    if len(statcom['gens']) != 0: 
        for gen in statcom['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            if ival == 0:
                print('STATCOM is OFF')
            else:

                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f,gen['gen_q_ini'],gen['gen_q_ini'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #Initialise statcom at 0MVAr
                standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(gen['gen_q_ini'])+","+str(gen['gen_q_ini'])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                vlt_lvl = af.get_bus_info(gen['poc_bus'],'PU')
                vlt_lvl = vlt_lvl[gen['poc_bus']]['PU']
                psspy.plant_data_4(gen['gen_bus'],0,[_i,_i],[vlt_lvl,gen['poc_bus']])
                psspy.fnsl([1,0,0,1,1,0,0,0])
                psspy.fnsl([1,0,0,1,1,0,0,0])
                standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                standalone_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
#                step_solv_paras()
                if(af.test_convergence(method='fnsl',taps='step')>1.0):raise


# Responsse of gens with voltage droop control mode when contingency occurs:
def lckd_gens_vdc(gens_with_vdc): 
    global standalone_script
    if len(gens_with_vdc['gens']) != 0:
        for i in range(0,5): #If multiple generator participate in the QV droop, then repeat the process to make sure the actual voltage settle well
            for gen in gens_with_vdc['gens']:
                ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
                if ival == 0:
                    print('GEN is OFF')
                else:
                    poc_v_actual = af.get_bus_info(gen['poc_bus'],'PU')
                    poc_v_actual = poc_v_actual[gen['poc_bus']]['PU'] #poc volt level
                     
                    delta_v = poc_v_actual - gen['poc_v_spt']
                    vol_deadband = gen['poc_v_dbn']/2
                    if delta_v > vol_deadband:
                        delta_v = delta_v - vol_deadband #compensate for only part outside deadband/2
                    elif delta_v < -vol_deadband:
                        delta_v = delta_v + vol_deadband
                    else: delta_v = 0 # within the deadband
                    
                    q_poc_req = -((gen['poc_q_max']/(gen['poc_drp_pct']/100.0)) * delta_v) # Consider Qmax as the base for droop
                    
                    if q_poc_req > gen['poc_q_max']: # Limit the compensation to Qcorner at POC
                        q_poc_req = gen['poc_q_max']
                    if q_poc_req < -gen['poc_q_max']: 
                        q_poc_req = -gen['poc_q_max']

                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    q_poc = -s_poc.imag # poc q MVAr
                    delta_q = q_poc_req - q_poc
                     
                    tol_q = 1.0 #MVAr
                    k_factor = 0.25 # regression factor
                    iter_num = 15
                    while abs(delta_q) > tol_q and iter_num > 0:
                         
                        ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                        q_gen_new = q_gen + delta_q*k_factor
                         
                        psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new,q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                        psspy.fnsl([0,0,0,1,0,0,0,0])
                        psspy.fnsl([0,0,0,1,0,0,0,0])
                        psspy.fnsl([0,0,0,1,0,0,0,0])
                        standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                        standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                        standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
#                        lckd_solv_paras()
                        if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                        s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                        q_poc = -s_poc.imag # poc q MVAr
                        delta_q = q_poc_req - q_poc
                        iter_num -=1
                        
                        
# Responsse of gens with hybrid (PF and V_PI) or direct voltage control mode when contingency occurs:
def lckd_gens_hc(gens_with_pf_vc):
    global standalone_script
    if len(gens_with_pf_vc['gens']) != 0:
        for gen in gens_with_pf_vc['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            if ival == 0:
                print('GEN is OFF')
            else:
                gen_q_max = gen['gen_q_max']
                gen_q_min = -gen['gen_q_max']
                ierr, mc_q_max=psspy.macdat(gen['gen_bus'],gen['gen_id'],'QMAX') #current maximum reactive power of the machine
                ierr, mc_q_min=psspy.macdat(gen['gen_bus'],gen['gen_id'],'QMIN')
                if gen_q_max < mc_q_max: gen_q_max = mc_q_max # keep original value if it provide a wider range
                if gen_q_min > mc_q_min: gen_q_min = mc_q_min
                
                iter_num = 5
                for i in range(1, iter_num):
                    gen_q_max_ = gen_q_max * i / iter_num
                    gen_q_min_ = gen_q_min * i / iter_num
                    psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, gen_q_max_, gen_q_min_,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #release the Q capability of the plant
                    standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(gen_q_max_)+","+str(gen_q_min_)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([0,0,0,1,0,0,0,0])
    ##                af.print_volts(lower=0.9, upper=1.1) #for debugging
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
    #                lckd_solv_paras()
                    if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise

def lckd_gens_qfix(gens_with_qfix):
    global standalone_script
    if len(gens_with_qfix['gens']) != 0: 
        for gen in gens_with_qfix['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            if ival == 0:
                print('GEN is OFF')
            else:
                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #Update the Control Mode to Not a wind machine so Qmax, Qmin can be updated
                s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1] #power measured at POC, from POC back to the SF (reversed power flow)
                q_poc = -s_poc.imag # poc q MVAr generated from SF
                q_poc_req = gen['poc_q_gen']
                delta_q = q_poc_req - q_poc
                tol_q = 1.0 #MVAr
                k_factor = 0.85 # regression factor
                iter_num = 15
                while abs(delta_q) > tol_q and iter_num > 0:
                    ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                    q_gen_new = q_gen + delta_q*k_factor
                    psspy.machine_chng_2(int(gen['gen_bus']),r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new, q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#                    psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[p_gen_new,_f, q_gen_new, q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    af.print_volts(lower=0.9, upper=1.1) #for debugging
#                    lckd_solv_paras()
                    if(af.test_convergence(method='fnsl',taps='locked')>1.0):
                        af.print_volts(lower=0.9, upper=1.1)
                        raise

                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    q_poc = -s_poc.imag # poc q MVAr
                    delta_q = q_poc_req - q_poc
                    iter_num -=1

# Reponse of the statcom
def lckd_statcom(statcom):
    global standalone_script
    if len(statcom['gens']) != 0: 
        for gen in statcom['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            if ival == 0:
                print('STATCOM is OFF')
            else:
                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[1,_i,_i,_i,_i,_i],[_f,_f,gen['gen_q_max'],gen['gen_q_min'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) # release the capacity of the statcom for voltage regulation
                standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[1,_i,_i,_i,_i,_i],[_f,_f,"+str(gen['gen_q_max'])+","+str(gen['gen_q_min'])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                psspy.fnsl([0,0,0,1,0,0,0,0])
                psspy.fnsl([0,0,0,1,0,0,0,0])
                standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                standalone_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
#                lckd_solv_paras()
                if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                
def pre_event_condition():
    global statcom, gens_with_pf, gens_with_pf_vc, gens_with_vdc
    if statcom != {}: ini_statcom(statcom)
    if gens_with_pf != {}: init_gens_pf(gens_with_pf)
#    if gens_with_pf_vc != {}: init_gens_hc(gens_with_pf_vc)
    if gens_with_vdc != {}: init_gens_vdc(gens_with_vdc)
   

def post_event_condition():
    global statcom, gens_with_pf, gens_with_pf_vc, gens_with_vdc
#    if statcom != {}: mf.lckd_statcom(statcom)
#    if gens_with_pf != {}: mf.lckd_gens_pf(gens_with_pf)
#    if gens_with_pf_vc != {}: mf.lckd_gens_hc(gens_with_pf_vc)
#    if gens_with_vdc != {}: mf.lckd_gens_vdc(gens_with_vdc)
    if statcom != {}: lckd_statcom(statcom)
#    if gens_with_pf != {}: lckd_gens_pf(gens_with_pf) # if pf control is in use, pq flow does not change at contingency
    if gens_with_pf_vc != {}: lckd_gens_hc(gens_with_pf_vc)
    if gens_with_vdc != {}: lckd_gens_vdc(gens_with_vdc)
    
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
    global ModelCopyDir, standalone_script, log
    log=''
    #generatorStatus=check_gen_status(gen_bus) #function that checks generator status at given bus and returns 1 or 0.
    
    #Create dataframe to store voltage levels, Contingency results, Generation Changes results, fault levels
    vltg_lvls_df = pd.DataFrame()
    bus_df_final = pd.DataFrame()
    brch_df_final = pd.DataFrame()
    bus_df_Pchange = pd.DataFrame()
    brch_df_Pchange = pd.DataFrame()
    fault_df = pd.DataFrame()
    
    ContRating = 'RATING1' #'RATING2'
    
    # Load the base case, apply the initialised condition, solve load flow and save it as the base results - Normal State
    with silence():

        # Load the base case implement initial condition and save it as before genchange
        psspy.case(base_model +"\\"+ sav_file) # load the case
        sav_file_name = os.path.basename(base_model +"\\"+ sav_file) # get the file name
        pre_event_condition() # voltage droop or power factor control initialisation if applicable
        wrkg_spc = createPath(ModelCopyDir+"\\"+testRun+"\\"+'Normal State') # working space folder
        psspy.save(wrkg_spc +"\\" +sav_file_name) #save the case with intial conditions
#        standalone_script+="psspy.save(\\"+str('Normal State')+"\\"+str(sav_file_name)+")\n"
            
            
        # Monitor variables and add to final DataFrame
        brch_df0, bus_df0 = monitoring()
        vltg_lvls_df = vltg_lvls_df.append(bus_df0)
        bus_df0['CaseNr']= "case0 (base)"
        bus_df0['VolDev (%)']= 0
        brch_df0['CaseNr']= "case0 (base)"
        brch_df0['Loading (%)']= np.round((brch_df0['MVA'] / brch_df0[ContRating])*100,2)
        bus_df_final = bus_df_final.append(bus_df0)
        brch_df_final = brch_df_final.append(brch_df0)
        

    # Contingency cases 
    for case_num in active_scenarios: # run all the active case
        scenario=SteadyStateDict[case_num]
        log+=case_num
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
                    execute_event(scenario[i], applied_value = 'ini_value') # apply initial value for genchange cases
                    if(af.test_convergence(method="fnsl", taps="step")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num)+' System did not converge')# THROW ERROR IF MISMATCH >1
            pre_event_condition() # implement initialisation condition if applicable e.g. implement voltage droop control, pf control
            psspy.save(wrkg_spc +"\\" +  'Before_' +sav_file_name) #save the case with intial conditions - before event
#            standalone_script+="psspy.save('Before_"+str(sav_file_name)+"')\n"                                     #ONLY FOR DEBUGING THE INITIALISATION PROCESS
            standalone_script+= "# Load the case before contingencis....\n"
            standalone_script+="psspy.case('Before_"+str(sav_file_name)+"')\n"


            # Monitor variables before event if needed
#            brch_df_P1, bus_df_P1 = monitoring()
#            bus_df_P1['CaseNr']= case_num_str
#            bus_df_P1['Case_Code']= scenario[0]['Case_Code']
#            brch_df_P1['CaseNr']= case_num_str
#            brch_df_P1['Case_Code']= scenario[0]['Case_Code']
#            brch_df_P1['Loading_ini (%)']= np.round((brch_df_P1['MVA'] / brch_df_P1[ContRating])*100,2)
#            bus_df_Pupdate = bus_df_P1
#            brch_df_Pupdate = brch_df_P1
                
            # separate monitoring for contingencies and genchange
            if (scenario[0]['TestType']!='Trip'): # if genchange event, then monitor the parameters for comparision.
                brch_df_P1, bus_df_P1 = monitoring()
                bus_df_P1['CaseNr']= case_num
                bus_df_P1['Case_Code']= scenario[0]['Case_Code']
                brch_df_P1['CaseNr']= case_num
                brch_df_P1['Case_Code']= scenario[0]['Case_Code']
                brch_df_P1['Loading_ini (%)']= np.round((brch_df_P1['MVA'] / brch_df_P1[ContRating])*100,2)
                bus_df_Pupdate = bus_df_P1
                brch_df_Pupdate = brch_df_P1
            
            # Apply the contingencies/genchanges, implement post event condition and save the case 
            standalone_script+= "# Apply the contingencies....\n"
            for i in range(0, len(scenario)):
#                if scenario[i]['end_value'] != '':
                execute_event(scenario[i], applied_value = 'end_value') # apply end value for genchange cases
                if(af.test_convergence(method="fnsl", taps="locked")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) +' System did not converge')# THROW ERROR IF MISMATCH >1
            post_event_condition() # implement post event conditions
            if(af.test_convergence(method="fnsl", taps="locked")>1):error_file.write('\n' + str(sav_file)+ ' '+ str(case_num) +' System did not converge')# THROW ERROR IF MISMATCH >1
            psspy.save(wrkg_spc +"\\" +  'After_' +sav_file_name)  #save the case after the contingency with locked taps
#            standalone_script+="psspy.save(\\"+str(case_num_str)+"\\After_"+str(sav_file_name)+")\n\n\n"
            standalone_script+= "# Save the case after contingencies applied\n"
            standalone_script+="psspy.save('After_"+str(sav_file_name)+"')\n\n"
            text_file = open(ModelCopyDir+"\\"+testRun+"\\"+str(case_num_str)+'\\'+str(sav_file_name[0:-4])+".py", "w")
            text_file.write(standalone_script)
            text_file.close()

            # Monitor variables after event and compare
#            brch_df_P2, bus_df_P2 = monitoring()
#            bus_df_Pupdate['VolDev (%)'] = np.round((bus_df_P2['PU'] - bus_df_P1['PU'])*100,4)
#            bus_df_Pupdate['PU_final'] = bus_df_P2['PU'] 
#            bus_df_Pupdate['ANGLED_final'] = bus_df_P2['ANGLED'] 
#            brch_df_Pupdate['Loading_final (%)'] = np.round((brch_df_P2['MVA'] / brch_df_P2[ContRating])*100,2)
#            brch_df_Pupdate['MVA_final'] = brch_df_P2['MVA']
#            brch_df_Pupdate['P_final'] = brch_df_P2['P']
#            brch_df_Pupdate['Q_final'] = brch_df_P2['Q']
#            brch_df_Pupdate['PCTMVARATE_final'] = brch_df_P2['PCTMVARATE']
#            bus_df_Pchange = bus_df_Pchange.append(bus_df_Pupdate)
#            brch_df_Pchange = brch_df_Pchange.append(brch_df_Pupdate)


            # separate monitoring for contingencies and genchange
            if (scenario[0]['TestType']=='Trip'): # if it is a contingency case
                brch_df, bus_df = monitoring()
                bus_df['CaseNr']= case_num
                bus_df['Case_Code']= scenario[0]['Case_Code']
                bus_df['VolDev (%)']= np.round((bus_df['PU'] - bus_df0['PU'])*100,4) #calculates deviaton compared to voltage level before contingency
                brch_df['CaseNr']= case_num
                brch_df['Case_Code']= scenario[0]['Case_Code']
                brch_df['Loading (%)']= np.round((brch_df['MVA'] / brch_df[ContRating])*100,2)
                bus_df_final = bus_df_final.append(bus_df)
                brch_df_final = brch_df_final.append(brch_df)
            else: # if it is a setpoint change test:
                brch_df_P2, bus_df_P2 = monitoring()
                bus_df_Pupdate['VolDev (%)'] = np.round((bus_df_P2['PU'] - bus_df_P1['PU'])*100,4)
                bus_df_Pupdate['PU_final'] = bus_df_P2['PU'] 
                bus_df_Pupdate['ANGLED_final'] = bus_df_P2['ANGLED'] 
                brch_df_Pupdate['Loading_final (%)'] = np.round((brch_df_P2['MVA'] / brch_df_P2[ContRating])*100,2)
                brch_df_Pupdate['MVA_final'] = brch_df_P2['MVA']
                brch_df_Pupdate['P_final'] = brch_df_P2['P']
                brch_df_Pupdate['Q_final'] = brch_df_P2['Q']
                brch_df_Pupdate['PCTMVARATE_final'] = brch_df_P2['PCTMVARATE']
                bus_df_Pchange = bus_df_Pchange.append(bus_df_Pupdate)
                brch_df_Pchange = brch_df_Pchange.append(brch_df_Pupdate)
    bus_df_final = bus_df_final[bus_df_final.CaseNr != 'case0 (base)']    
    results={df_to_sheet['volt_levels']['df']:vltg_lvls_df, df_to_sheet['line_loadings']['df']:brch_df_final, df_to_sheet['volt_fluc_gen_chng']['df']:bus_df_Pchange, df_to_sheet['volt_fluc_lol']['df']:bus_df_final, df_to_sheet['fault_levels']['df']: fault_df }
    return results, log

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


outputResultPath=ResultsDir+"\\"+testRun
make_dir(outputResultPath)

###############################################################################
#Text file to log errors
###############################################################################
error_file = open(outputResultPath + '\\' + 'error_log.txt','a')

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


def main():
    global ModelCopyDir, SteadyStateDict, outputResultPath, input_models, log
    out_file_name = 'Steady State Analysis Results.xlsx' # for recording the results
    writer = pd.ExcelWriter(outputResultPath+'\\'+out_file_name,engine = 'xlsxwriter')
    
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
        for snapshot_name in input_models.keys(): #loop the snapshots
            for config in ['on', 'off']: # two modes gen on and off
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