# -*- coding: utf-8 -*-
"""
Created on Thu Oct 22 15:33:37 2020

@author: Mervin Kall

01/04: update maximum number of iteration from 25(default) to 200 -> fix oscilation in Fprofile test
psspy.dynamics_solution_param_2([200,_i,_i,_i,_i,_i,_i,_i],[ 0.2,_f, 0.001,_f,_f,_f,_f,_f])

21/4/2022: Update Shunt bus as from input spreadsheet
#psspy.shunt_data(2,r"1",0,[0.0,Qcap])
28/4/2022: Update: added tune_parameters()
01/09/2022: Added the implement_droop_LF function to initinalise the model correctly with voltage droop characteristic.
            Note to update the droop_value, droop_base, and vol_deadband accordingly before using the script -> to be updated and called from the information spreadsheet in next change
29/11/2022: Update initialise_loadflow: one more loop to initilise at more exact value of Q
"""
#-----------------------------------------------------------------------------
# IMPORT PSSE
#-----------------------------------------------------------------------------
from __future__ import with_statement
from contextlib import contextmanager
import os, sys
import ntpath
from datetime import datetime

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
#import excelpy
import lntpy
import gicdata
#import pssexcel
import pssplot

redirect.psse2py()
with silence():
    psspy.psseinit(80000)
#_i=psspy.getdefaultint()
#_f=psspy.getdefaultreal()
#_s=psspy.getdefaultchar()


import dyntools

#### Kieran Kelly 2019 ####  
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()
#redirect.psse2py()


#-----------------------------------------------------------------------------
# IMPORT OTHER STUFF
#-----------------------------------------------------------------------------
import out_to_csv as conv
import math
import Vth_initialisation
import time
import ast
import shelve
import random
import TOV_calc
#import cmath
#-----------------------------------------------------------------------------
# GLOBAL VARIABLES
#-----------------------------------------------------------------------------
active_faults={}
event_queue=[] # types of events:  angle_change,var_change_abs, var_change_rel, change_VREF, const_change, apply_3phg, apply_2phg, apply_1phg, clear_3phg, clear_2phg, clear_1phg, add_cap, disconnect_cap
# event format: {'time': 1.0, 'type':'angle_change'}
#               {'time': 1.0, 'type':'apply_3phg', 'reactance': 0.0, 'resistance':0.1}
#               {'time': 1.0, 'type':'var_change', 'var_id': 112,'model':'AC7CU1','bus','mac':'1', 'value':2.3}
var_init_dict={}
chn_init_dict={} #not sure how to access machine array channel values directly, this is a workaround. 
plot_channels={}
standalone_script=""
#runtime_big = 0
#-----------------------------------------------------------------------------
# Functions
#-----------------------------------------------------------------------------
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
         
def get_bus_info(bus_list,param_list):
    bus_info={}
    if(type(bus_list)==list):
        for i in range (0,len(bus_list)):
            bus_info[bus_list[i]]={}
            ierr=psspy.bsys(sid=0,numbus=1, buses=bus_list[i])
            if ( (type(param_list)==list) and (ierr==0)):
                for param in param_list:
                    if((param=='AREA') or (param=='TYPE')):
                        ierr, param_val=psspy.abusint(0,2,param)
                    elif((param=='NAME')):
                        ierr, param_val=psspy.abuschar(0,2,param)
                    else:
                        ierr, param_val=psspy.abusreal(0,2,param)
                    if (ierr==0) and (param_val!=[[]]):
                        bus_info[bus_list[i]][param]=param_val[0][0]
            elif (type(param_list)==str):
                    if((param_list=='AREA') or (param_list=='TYPE')):
                        ierr, param_val=psspy.abusint(0,2,param_list)
                    elif((param_list=='NAME')):
                        ierr, param_val=psspy.abuschar(0,2,param_list)
                    else:
                        ierr, param_val=psspy.abusreal(0,2,param_list)
                    if (ierr == 0) and (param_val!=[[]]):
                        bus_info[bus_list[i]][param_list]=param_val[0][0]
    elif(type(bus_list)==int):
        bus_info[bus_list]={}
        ierr=psspy.bsys(sid=0,numbus=1, buses=bus_list)
        if ( (type(param_list)==list) and (ierr==0)):
            for param in param_list:
                if((param=='AREA') or (param=='TYPE')):
                    ierr, param_val=psspy.abusint(0,2,param)
                elif((param=='NAME')):
                    ierr, param_val=psspy.abuschar(0,2,param)                    
                else:
                    ierr, param_val=psspy.abusreal(0,2,param)
                if (ierr == 0) and (param_val!=[[]]):
                    bus_info[bus_list][param]=param_val[0][0]
        elif (type(param_list)==str):
                if((param_list=='AREA') or (param_list=='TYPE')):
                    ierr, param_val=psspy.abusint(0,2,param_list)
                elif((param_list=='NAME')):
                    ierr, param_val=psspy.abuschar(0,2,param_list)
                else:
                    ierr, param_val=psspy.abusreal(0,2,param_list)
                if (ierr == 0) and (param_val!=[[]]):
                    print('bus number is'+str(bus_list))
                    bus_info[bus_list][param_list]=param_val[0][0]
    return bus_info

def test_convergence(tree = 0):
    if(tree ==1):
        psspy.tree(1,0)
        psspy.tree(2,1)
        psspy.tree(2,1) 
        psspy.tree(2,1)
        
        
    psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([1,0,0,1,1,0,99,0])
    mismatch=psspy.sysmsm()
    print("the total mismatch is "+str(mismatch))
    if(abs(mismatch)>1):
        print("The system did not converge.")
    else:
        print("The system converged.")       
    return mismatch

#check if string is number
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def process_next_event(event_queue,start_offset, PSSEmodelDict): # added start_offset to delay the start of dynamic simulation 20/1/2022
    global standalone_script, runtime_big
    InfiniteBus=PSSEmodelDict['InfiniteBus']
    FaultBus=PSSEmodelDict['FaultBus']
    DummyTxBus=PSSEmodelDict['DummyTxBus']
    event=event_queue[0] #take always the first event only
    #psspy.run(0,event['time'], 500, 1, 0) #run until event occurs
    
#    if event['time'] - runtime_big > 300: #if runtime is too long, then reduce the resolution
#        psspy.run(0,event['time'] + start_offset, 500, 7, 0) #run until event occurs
#    else:
#        psspy.run(0,event['time'] + start_offset, 500, 1, 0) #run until event occurs
        
    psspy.run(0,event['time'] + start_offset, 500, 1, 0) #run until event occurs
    standalone_script+="psspy.run(0,"+str(event['time']+start_offset)+", 500, 1, 0)\n"
#    psspy.run(0,event['time'] + start_offset, 500, 7, 0) #run until event occurs # Use timestep of 7ms for long run (10min)
#    standalone_script+="psspy.run(0,"+str(event['time']+start_offset)+", 500, 7, 0)\n"
#    standalone_script+="psspy.run(0,"+str(event['time'])+", 500, 1, 0)\n"
    if(event['type']=='apply_3PHG'):
        psspy.dist_bus_fault(FaultBus,3,0.0,[event['resistance'],event['reactance']])#Apply 3PHG fault
        standalone_script+="psspy.dist_bus_fault("+str(FaultBus)+",3,0.0,["+str(event['resistance'])+","+str(event['reactance'])+"])\n"
#        standalone_script+="psspy.dist_bus_fault(2,3,0.0,["+str(event['resistance'])+","+str(event['reactance'])+"])\n"
    elif(event['type']=='apply_2PHG'):
        psspy.dist_scmu_fault_2([0,0,2,FaultBus,_i],[0.0,0.0,event['resistance'],event['reactance']])#apply 2PHG fault
        standalone_script+="psspy.dist_scmu_fault_2([0,0,2,"+str(FaultBus)+",_i],[0.0,0.0,"+str(event['resistance'])+","+str(event['reactance'])+"])\n"
    elif(event['type']=='apply_1PHG'):
        psspy.dist_scmu_fault_2([0,0,1,FaultBus,_i],[0.0,0.0,event['resistance'],event['reactance']])#apply 2PHG fault
        standalone_script+="psspy.dist_scmu_fault_2([0,0,1,"+str(FaultBus)+",_i],[0.0,0.0,"+str(event['resistance'])+","+str(event['reactance'])+"])\n"
    elif('clear' in event['type']):
        ierr=psspy.dist_clear_fault(1)
        standalone_script+="psspy.dist_clear_fault(1)\n"
    elif(event['type']=='angle_change'): #change voltage angle during runtime (using ideal transformer)
#        psspy.two_winding_chng_5(DummyTxBus,event['POC'],r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f, -1*event['angle'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
#        standalone_script+="psspy.two_winding_chng_5("+str(DummyTxBus) +","+str(event['POC'])+",'1',[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f, "+str(-1*event['angle'])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)\n"
        
        # Chagne the angle at the INF bus
        psspy.two_winding_chng_5(InfiniteBus,DummyTxBus,r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f, -1*event['angle'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
        standalone_script+="psspy.two_winding_chng_5("+str(InfiniteBus) +","+str(DummyTxBus)+",'1',[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f, "+str(-1*event['angle'])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)\n"
        
    elif(event['type']=='var_change_abs'): #change variable during runtime (e.g. for changing setpoint(s))        
        if(event['model_type']=='WAUX' or event['model_type']=='WGEN'):
            L=psspy.windmind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
            psspy.change_wnmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value']))
            standalone_script+="psspy.change_wnmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(event['value'])+")\n"
        elif(event['model_type']=='OTHER'):
            L = psspy.cctmind_buso(event['bus'],event['model'],'VAR')[1]
            psspy.change_cctbusomod_var(event['bus'],event['model'],event['rel_id']+1,float(event['value']))
            standalone_script+="psspy.change_cctbusomod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value']))+")\n"
        else:
            L=psspy.mdlind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
            psspy.change_plmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value']))
            standalone_script+="psspy.change_plmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value']))+")\n"
            
    elif(event['type']=='var_change_rel'): #change variable during runtime (e.g. for changing setpoint(s))
#        L=psspy.mdlind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#        standalone_script+="L=psspy.mdlind("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model_type'])+"', 'VAR')[1]\n"  #applies MDline int he actual script and sets variable L
#        event['abs_id']=L+event['rel_id']
#        #standalone_script+="abs_id=L+"str(event['rel_id'])+"\n"
#        var_value=init_val('VAR', event['abs_id']) #read previous value of variable and apply scaling to that
        
        if(event['model_type']=='WAUX' or event['model_type']=='WGEN'):
            L=psspy.windmind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#            standalone_script+="L=psspy.windmind("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model_type'])+"', 'VAR')[1]\n"  #applies MDline int he actual script and sets variable L
            event['abs_id']=L+event['rel_id']
        elif(event['model_type']=='OTHER'):
            L = psspy.cctmind_buso(event['bus'],event['model'],'VAR')[1]
#            standalone_script+="L=psspy.cctmind_buso("+str(event['bus'])+", '"+str(event['model_type'])+"', 'VAR')[1]\n"
            event['abs_id']=L+event['rel_id']
        else:
            L=psspy.mdlind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#            standalone_script+="L=psspy.mdlind("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model_type'])+"', 'VAR')[1]\n"  #applies MDline int he actual script and sets variable L
            event['abs_id']=L+event['rel_id']          
        var_value=init_val('VAR', event['abs_id']) #read previous value of variable and apply scaling to that        

        if(event['model_type']=='WAUX' or event['model_type']=='WGEN'):
            psspy.change_wnmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value'])*float(var_value))
            standalone_script+="psspy.change_wnmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value'])*float(var_value))+")\n"
        elif(event['model_type']=='OTHER'):
            psspy.change_cctbusomod_var(event['bus'],event['model'],event['rel_id']+1,float(event['value'])*float(var_value))
            standalone_script+="psspy.change_cctbusomod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value'])*float(var_value))+")\n"
        else:
            psspy.change_plmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value'])*float(var_value))
            standalone_script+="psspy.change_plmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value'])*float(var_value))+")\n"
        print(psspy.dsrval('VAR', event['abs_id'])[1])



#    elif(event['type']=='var_change_abs'): #change variable during runtime (e.g. for changing setpoint(s))        
#        if(event['model_type']=='WAUX' or event['model_type']=='WGEN'):
#            L=psspy.windmind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#            psspy.change_wnmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value']))
#            standalone_script+="psspy.change_wnmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(event['value'])+")\n"
#        else:
#            L=psspy.mdlind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#            psspy.change_plmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value']))
#            standalone_script+="psspy.change_plmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value']))+")\n"
#            
#    elif(event['type']=='var_change_rel'): #change variable during runtime (e.g. for changing setpoint(s))
#        if(event['model_type']=='WAUX' or event['model_type']=='WGEN'):
#            L=psspy.windmind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#            standalone_script+="L=psspy.windmind("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model_type'])+"', 'VAR')[1]\n"  #applies MDline int he actual script and sets variable L
#            event['abs_id']=L+event['rel_id']
#        else:
#            L=psspy.mdlind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#            standalone_script+="L=psspy.mdlind("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model_type'])+"', 'VAR')[1]\n"  #applies MDline int he actual script and sets variable L
#            event['abs_id']=L+event['rel_id']        
##        L=psspy.mdlind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
##        standalone_script+="L=psspy.mdlind("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model_type'])+"', 'VAR')[1]\n"  #applies MDline int he actual script and sets variable L
##        event['abs_id']=L+event['rel_id']
#        #standalone_script+="abs_id=L+"str(event['rel_id'])+"\n"
#        var_value=init_val('VAR', event['abs_id']) #read previous value of variable and apply scaling to that
#        # Issue: the var_value will be fixed to the first test case of the same test. e.g voltage setpoint test as the ID are remained. This is helpful when different type of event occuring in series for a test. but cause initialising condition not correct if the second test have same type of event but does not have same condition with first one.
#        # Need to reset this init_val after each test: clear var_init_dict at the end of the routin
#        if(event['model_type']=='WAUX' or event['model_type']=='WGEN'):
#            psspy.change_wnmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value'])*float(var_value))
#            standalone_script+="psspy.change_wnmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value'])*float(var_value))+")\n"
#        else:
#            psspy.change_plmod_var(event['bus'], event['mac'], event['model'], event['rel_id']+1, float(event['value'])*float(var_value))
#            standalone_script+="psspy.change_plmod_var("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(float(event['value'])*float(var_value))+")\n"
#        print(psspy.dsrval('VAR', event['abs_id'])[1])
#        #psspy.change_plmod_var(90104, r"""1""", r"""SMAHYC25""", 17, 1.01)
        pass        
    elif(event['type']=='con_change_abs'):
        psspy.change_plmod_con(event['bus'], event['mac'], event['model'], event['rel_id']+1, event['value'])
        standalone_script+="psspy.change_plmod_con("+str(event['bus'])+", '"+str(event['mac'])+"', '"+str(event['model'])+"', "+str(event['rel_id']+1)+", "+str(event['value'])+")\n"
    elif(event['type']=='VREF_change_abs'):
        psspy.change_vref(event['bus'], event['mac'], event['value']) 
        standalone_script+="psspy.change_vref("+str(event['bus'])+", "+str(event['mac'])+", "+str(event['value'])+")\n"
        #!!!!!!!!!standalone_script+=!!!!!!!!!!!
    elif(event['type']=='VREF_change_rel'):
        Vref_init=init_val('CHN',plot_channels[event['chn']])
        psspy.change_vref(event['bus'], event['mac'], float(event['value'])*float(Vref_init)) 
        standalone_script+="psspy.change_vref("+str(event['bus'])+", "+str(event['mac'])+", "+str(float(event['value'])*float(Vref_init))+")\n"
    elif(event['type']=='shunt_on'): #add capacitor at POC (typically causing overvoltage)
        psspy.shunt_chng(event['bus'],event['id'],1,[_f,_f])
        standalone_script+="psspy.shunt_chng("+str(event['bus'])+','+str(event['id'])+",1,[_f,_f])\n"
    elif(event['type']=='shunt_off'): #disconnect capacitor at POC
        psspy.shunt_chng(event['bus'],event['id'],0,[_f,_f])
        standalone_script+="psspy.shunt_chng("+str(event['bus'])+','+str(event['id'])+",0,[_f,_f])\n"
    elif(event['type']=='imp_change'): #change grid impedance to pre-set value
        psspy.branch_chng_3(event['fromBus'],event['toBus'],event['branchID'],[_i,_i,_i,_i,_i,_i],[ event['R'], event['X'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
        standalone_script+="psspy.branch_chng_3("+str(event['fromBus'])+','+str(event['toBus'])+','+str(event['branchID'])+',[_i,_i,_i,_i,_i,_i],['+str(event['R'])+', '+str(event['X'])+',_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)\n'
    else:
        return 1, event_queue #return error, if event does not fit any of the above categories.
    
#    runtime_big = event['time']
    
    del event_queue[0]
    return 0, event_queue #if successful, return 0
        
        
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

#interpolates pointwise-defined profile for use in PSS/E (profiles applied to variables in PSS/E need to be defined as discrete steps, with every variable change being explicitly called at a specific simulation time)
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
    
    
x_data=[0,5, 5.5, 15, 15.5, 25]
y_data=[50, 50, 52, 52, 50, 50]
                
def tune_parameters(PSSEmodelDict):
    # Only activated in tunning process. Should be commented out when finishing this process
#    busgen1 = 367176
#    busgen2 = 367179
#    pccmodel = r"""EMSPCI2_1"""
#    invmodel = r"""ING1BI2_1"""
#    psspy.change_wnmod_icon(busgen1,r"""1""",pccmodel,7, 2) # Control mode
    ###############################################################
    # Issue: casesmall80 -> plant get stuck in HVRT and then trip when voltage disturbance increase from 0.9pu to 1.1pu at inf bus
    #   -> plant does not recognise it is just one spike and need to control the Q down from PPC.
    
    # Check0: need a delay in relising the HVRT flag to be activated: not possible. Confirmed by David from Ingeteam: when voltage spike HVRT will be triggered
    # Check1: Increase kV -> not effective if only kV is changed
#    psspy.change_wnmod_con(busgen1,r"""1""",pccmodel,37, 30) # kV
    # check2: Update the protection settings: the plant can ride through
#    psspy.change_wnmod_con(busgen1,r"""1""",invmodel,44, 1.25)
#    psspy.change_wnmod_con(busgen2,r"""1""",invmodel,44, 1.25)
    # check 3: increase the Iq response from INV in HVRT:
#    psspy.change_wnmod_con(busgen1,r"""1""",invmodel,73, 1.30)
#    psspy.change_wnmod_con(busgen2,r"""1""",invmodel,73, 1.30)
#    psspy.change_wnmod_con(busgen1,r"""1""",invmodel,71, 1.17) #Propose to start inject Iq when voltage reach 1.17pu.
#    psspy.change_wnmod_con(busgen2,r"""1""",invmodel,71, 1.17)
    
    ###############################################################
    # Issue: casesmall35 -> Qsetpoint negative causing plant to expereience oscilation due to LVRT on/off
    # 
    # Check0: Q control gain? not avaiable
    # Check1; Adjust Q change level from 29MVAr to 24MVAr as requested
    # Check2: reduce threhold for LVRT support to 0.82
#    psspy.change_wnmod_con(busgen1,r"""1""",invmodel,63, 0.80)
#    psspy.change_wnmod_con(busgen2,r"""1""",invmodel,63, 0.80)
#    psspy.change_wnmod_con(busgen1,r"""1""",invmodel,81, 0.80)
#    psspy.change_wnmod_con(busgen2,r"""1""",invmodel,81, 0.80)
    pass
    
def set_channels(PSSEmodelDict):
    global standalone_script
    standalone_script+="#Add output channels\n"
    # in input excel sheet, bus numbers for HV, MV and LV shoudl be defined, --> the powers will be measured at the fromBus
    # in this section output channels will be added automatically for these points
    # more channels can be added here if required
    #add channels
    
    #detect measurement locations
    meas_locs={}
    for key in PSSEmodelDict.keys():
        if('fromBus' in key):
            loc=key.replace('fromBus','')
            meas_locs[loc]={'fromBus':PSSEmodelDict[key], 'toBus':PSSEmodelDict[key.replace('fromBus', 'toBus')], 'measBus':PSSEmodelDict[key.replace('fromBus', 'measBus')]}
    psspy.delete_all_plot_channels()
    standalone_script+="psspy.delete_all_plot_channels()\n"
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
    
    psspy.var_channel([-1,L_INV+102],'INV1_FRT_FLAG' ) # FRT detection flag
    psspy.var_channel([-1,L_INV+186],'INV1_LVRT' )
    psspy.var_channel([-1,L_INV+187],'INV1_HVRT' )
    psspy.var_channel([-1,L_INV+9],'INV1_IQ_COMMAND' )
    psspy.var_channel([-1,L_INV+40],'INV1_IP_COMMAND' )
    
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

    psspy.var_channel([-1,L_INV2+200],'INV2_Frequency' )
#    psspy.var_channel([-1,L_INV2+83],'INV2_Vq' )    
    psspy.var_channel([-1,L_INV2+82],'INV2_Vd' )
    psspy.var_channel([-1,L_INV2+83],'INV2_Vq' )
    psspy.var_channel([-1,L_INV2+86],'INV2_Id' )
    psspy.var_channel([-1,L_INV2+87],'INV2_Iq' )
    psspy.var_channel([-1,L_INV2+16],'INV2_FRT_FLAG' ) # FRT detection flag
    psspy.var_channel([-1,L_INV2+79],'INV2_ANGLE' )
#    psspy.var_channel([-1,L_INV2+186],'INV2_LVRT' )
#    psspy.var_channel([-1,L_INV2+187],'INV2_HVRT' )
#    ierr = chsb(sid, all, status)
#    ierr = machine_array_channel(status, id, ident)
#    ierr = psspy.chsb(9944, 0, [-1, -1, -1, 1, 2, 1]) #STATUS(5)=2 PELEC
#    ierr = psspy.machine_array_channel(2, '1', "PELEC") #STATUS(2)=2 PELEC, machine electrical power (pu on SBASE)

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
    
#    psspy.var_channel([-1,L_PPC+14],'POC Qref Stp' ) # POI_Var_Spt
#    psspy.var_channel([-1,L_PPC+16],'POC Vol Stp' ) # POI_Vol_Spt
#    psspy.var_channel([-1,L_PPC+17],'POC Hz Stp' ) # POI_Hz_Spt
#    psspy.var_channel([-1,L_PPC+18],'POC Pref Stp' ) # PwrAtLimSales
#    psspy.var_channel([-1,L_PPC+1],'Volt_RB' )
#    psspy.var_channel([-1,L_PPC+2],'P_pcc' )
#    psspy.var_channel([-1,L_PPC+3],'Q_pcc' )
#    psspy.var_channel([-1,L_PPC+4],'S_pcc' )
#    psspy.var_channel([-1,L_PPC+56],'Ppv_cmd_inv' )
#    psspy.var_channel([-1,L_PPC+57],'Qpv_cmd_inv' )
#    psspy.var_channel([-1,L_PPC+75],'FrtActive' )
#    psspy.var_channel([-1,L_PPC+76],'FRT_ExitTm' )
#    psspy.var_channel([-1,L_PPC+1],'Volt_RB' )
#    psspy.var_channel([-1,L_PPC+2],'P_pcc' )
#    psspy.var_channel([-1,L_PPC+3],'Q_pcc' )
#    psspy.var_channel([-1,L_PPC+3],'S_pcc' )
    
    
#    psspy.machine_iterm_channel([-1,-1,-1,367176],r"""1""",r"""INV_Itot""")
#    psspy.branch_mva_channel([-1,-1,-1,367176, 367175],r"""1""",r"""INV_MVA""")
#    psspy.machine_iterm_channel([-1,-1,-1,367179],r"""1""",r"""INV2_Itot""")
#    psspy.branch_mva_channel([-1,-1,-1,367179, 367178],r"""1""",r"""INV2_MVA""")
#
#    ierr, L_INV = psspy.windmind(367176,'1','WGEN','VAR')
#    ierr, L_INV2= psspy.windmind(367179,'1','WGEN','VAR')
#    ierr, K_INV = psspy.windmind(367176,'1','WGEN','STATE')
#    ierr, K_INV2 = psspy.windmind(367179,'1','WGEN','STATE')
#    psspy.state_channel([-1, K_INV], r"""P_Gen1_machine_array""") # (STATE K) Measured active power (Pgrid)    
##    standalone_script+="psspy.state_channel([-1, "+str(K_INV)+"], 'P_Gen1_machine_array')\n" #Exporting measurement cmd for AEMO reference
#    psspy.state_channel([-1, K_INV + 4], r"""INV1_Id""") # (STATE K+4) Id output of inverter (p.u. of inverter base)
#    psspy.state_channel([-1, K_INV + 5], r"""INV1_Iq""") #Iq (STATE K+5) output of inverter (p.u. of inverter base)
#    psspy.state_channel([-1, K_INV2 + 4], r"""INV2_Id""") # (STATE K+4) Id output of inverter (p.u. of inverter base)
#    psspy.state_channel([-1, K_INV2 + 5], r"""INV2_Iq""") #Iq (STATE K+5) output of inverter (p.u. of inverter base)
#    psspy.var_channel([-1, L_INV + 6], "INV1_Irradiance") # Voltage Disturbance Detection
#    psspy.var_channel([-1, L_INV + 7], "P_cmd") # P command
##    standalone_script+="psspy.var_channel([-1, "+str(L_INV + 7)+"], 'P_cmd')\n" #Exporting ref cmd for AEMO reference
#    psspy.var_channel([-1, L_INV2 + 6], "INV2_Irradiance") # Voltage Disturbance Detection
#    psspy.var_channel([-1, L_INV + 15], "INV1_VdFlag") # Voltage Disturbance Detection
#    psspy.var_channel([-1, L_INV2 + 15], "INV2_VdFlag") # Voltage Disturbance Detection
#    
#    ierr, L_PPC = psspy.windmind(367176,'1','WAUX','VAR')
#    ierr, K_PPC = psspy.windmind(367176,'1','WAUX','STATE')
#    psspy.var_channel([-1, L_PPC + 6], "Vref_POC") # 
#    psspy.var_channel([-1, L_PPC + 5], "PFref") # 
#    psspy.var_channel([-1, L_PPC + 4], "Qref_POC") # 
#    psspy.var_channel([-1, L_PPC + 1], "Pref_POC") # 
#    psspy.var_channel([-1, L_PPC + 20], "Var20") # 
#    psspy.var_channel([-1, L_PPC + 21], "Var21") # 
#    psspy.state_channel([-1, K_PPC + 2], "State2") # 
#    psspy.state_channel([-1, K_PPC + 5], "State5") # 
#    psspy.state_channel([-1, K_PPC + 6], "State6") # 
#    
#    # Get tap position of the main transformer:
#    ierr, rval = psspy.xfrdat(367171, 367174, '1', 'RATIO')
##    ierr, rval = psspy.xfrdat(367171, 367174, '1', 'RATIO2')
#    # XFRINT
#    # ATRNINT


    # Include all vars and states
    # ierr, L_PPC = psspy.windmind(367176,'1','WAUX','VAR')
    # ierr, K_PPC = psspy.windmind(367176,'1','WAUX','STATE')
    # ierr, L_INV = psspy.windmind(367176,'1','WGEN','VAR')
    # ierr, K_INV = psspy.windmind(367176,'1','WGEN','STATE')
    # ierr, L_PPC2 = psspy.windmind(367179,'1','WAUX','VAR')
    # ierr, K_PPC2 = psspy.windmind(367179,'1','WAUX','STATE')
    # ierr, L_INV2= psspy.windmind(367179,'1','WGEN','VAR')
    # ierr, K_INV2 = psspy.windmind(367179,'1','WGEN','STATE')
    # for l1 in range (55): #54
    #     psspy.var_channel([-1, L_INV + l1], "") #Inverter VAR
    # for k1 in range (26): #25
    #     psspy.state_channel([-1, K_INV + k1], "") #Inverter State
    # for l2 in range (28): #27
    #     psspy.var_channel([-1, L_PPC + l2], "") #PPC VAR
    # for k2 in range (27): #26
    #     psspy.state_channel([-1, K_PPC + k2], "") #PPC State


    # psspy.var_channel([-1,L_PPC+2],'PPC_L2' ) #f_ref(L+2)
    # psspy.var_channel([-1,L_PPC+16],'PPC_L16' ) #Pg(L+16)
    # psspy.var_channel([-1,L_PPC+17],'PPC_L17' ) #Pred(L+17)
    # psspy.state_channel([-1,K_PPC+3],'PPC_K3' ) #f_grid(K+3)
    # psspy.state_channel([-1,K_PPC+4],'PPC_K4' ) #Pref(K+4)
    # psspy.state_channel([-1,K_PPC+5],'PPC_K5' ) #Qref(K+5)
    # psspy.state_channel([-1,K_PPC+6],'PPC_K6' ) #Vref(K+6)
    # psspy.state_channel([-1,K_PPC+8],'PPC_K8' ) #deltaP(K+8)

    # psspy.var_channel([-1,L_INV+15],'INV1_L15' ) #VDdetection
    # psspy.var_channel([-1,L_INV+27],'INV1_L27' ) # Trip signal
    # psspy.var_channel([-1,L_INV+28],'INV1_L28' ) #Voltage Event Detection Flag
    # psspy.var_channel([-1,L_INV+29],'INV1_L29' ) # Freq Event Detection Flag
    # psspy.state_channel([-1,K_INV+0],'INV1_K0' ) # Pgrid_measured
    # psspy.state_channel([-1,K_INV+2],'INV1_K2' ) # Vgrid_measured
    # busNum = 1
    # psspy.voltage_and_angle_channel([-1,-1,-1,busNum], ['U_INF', 'ANG_INF'])




#     ierr, ppc_var_id = psspy.mdlind(100,'1','EXC','VAR')
#     ierr, ppc_state_id = psspy.mdlind(100,'1','EXC','STATE')
#     ierr, inv_var_id = psspy.mdlind(100,'1','GEN','VAR')
#     ierr, inv_state_id = psspy.mdlind(100,'1','GEN','STATE')

# #
#     psspy.machine_array_channel([chn_idx,5,100],'1','Q_CMD_TO_INV')
#     plot_channels['Q_CMD_TO_INV']=chn_idx
#     chn_idx+=1
#     psspy.machine_array_channel([chn_idx,8,100],'1','P_CMD_TO_INV')
#     plot_channels['P_CMD_TO_INV']=chn_idx
#     chn_idx+=1    
#     psspy.machine_array_channel([chn_idx,11,100],'1','SYNC_VREF')
#     plot_channels['SYNC_VREF']=chn_idx
#     chn_idx+=1
    
#     psspy.var_channel([-1,ppc_var_id+17],'PPC_LVRT' ) #PPC lvrt 
#     psspy.var_channel([-1,ppc_var_id+18],'PPC_HVRT' ) #PPC lvrt flag
#     var_channels=[  
                    
#                     ['INV_LVRT',186,inv_var_id],
#                     ['INV_HVRT',187,inv_var_id],
#                     ['INV_VREF', 11, ppc_var_id], 
#                     ['PMECH_old', 30, ppc_var_id],
#                     ['EDF_old', 31, ppc_var_id],
#                     ['XADIFD_old', 32, ppc_var_id],
# #                    ['INT_Q_CMD_TO_INV',174, inv_var_id],
# #                    ['INT_P_CMD_TO_INV', 175, inv_var_id],
#                     ['POC_Q_STP', 12, ppc_var_id],
#                     ['POC_P_STP', 13, ppc_var_id],
#                     ['PwrAtLm_L38', 38, ppc_var_id],
#                     ['VDNPwr_L38', 39, ppc_var_id],
#                     ['Pctrl_L+40', 40, ppc_var_id],
                    
#                     ['Pstp', 88, ppc_var_id],
#                     ['Qstp', 89, ppc_var_id],
#                     ['Vstp', 90, ppc_var_id],
                   
                    
#                     ['P_avail', 21, ppc_var_id],
# #                    ['POC_VOL_STP', 16, ppc_var_id],
# #                    ['P_STP', 20, ppc_var_id],
# #                    ['F_MEAS_POC', 28, ppc_var_id],
#                     ['PPRIM_INITIAL', 88, inv_var_id],
# #                    ['PLL_MEAS_INV',134, inv_var_id],
#                     ['FRT_STATE', 163, inv_var_id],
#                     ['INV_LVRT_STATE', 170, inv_var_id],
#                     ['INV_HVRT_STATE', 171, inv_var_id],
# #                    ['ID_STP', 40, inv_var_id],
# #                    ['IQ_STP_PRE_LIM', 9, inv_var_id],                                      
                    
#                     ]
#     state_channels=[
#                     ["IQ_ramp", 8, inv_state_id],
#                     ["IQ_Pt1", 14, inv_state_id],
                    
#             ]
                          
#     for var in var_channels:
#         psspy.var_channel([chn_idx, var[2]+var[1]], var[0])
#         plot_channels[var[0]]=chn_idx
#         chn_idx+=1
        
#     for state in state_channels:
#         psspy.state_channel([chn_idx, state[2]+state[1]], state[0])
#         plot_channels[state[0]]=chn_idx
#         chn_idx+=1
#    
#    psspy.var_channel([-1,inv_var_id+102], 'INV_SPIKE_SUP') #INV Lvrt flag
#    psspy.var_channel([-1,inv_var_id+163], 'INV_FRT_STATE') #INV Lvrt flag  
    
#   
#    psspy.state_channel([-1,ppc_state_id+10],'F_FILTER_OUT' ) # frequency filter output at PPC
    #--------------------------------------------------------------------------
    pass

#def calc_thv_volt(V_POC, GRID_X, GRID_R, P_POC, Q_POC, Vbase, MVA_base):
#    S_POC=(Q_POC*1j+P_POC)*1000000
#    I_POC=S_POC/V_POC/math.sqrt(3)
#    I_POC=I_POC.conjugate()
#    GRID_Z_=GRID_R+GRID_X*1j
#    Vth=GRID_Z_*I_POC*math.sqrt(3)+V_POC
#    return Vth

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
    
def initialise_loadflow(workspace_folder, ProjectDetailsDict, PSSEmodelDict, setpoint, Vbase_POC, Vth, ANG_POC, GRID_MVA, InfiniteBus, FaultBus, DummyTxBus, POC, F_dist ):
    global standalone_script
    test_convergence()
    #add Zingen machine, unless machine at bus 1 is alreay present. --> already present for most of the tests
    # add entry for Zingen machine to dyr file
    #add ideal transformer between grid and fault bus. 
#    ierr, base = psspy.busdat(1,'BASE')      
#    psspy.bus_data_4(DummyTxBus,0,[1,1,1,1],[base, 1.0,0.0, 1.1, 0.9, 1.1, 0.9],r"""DUM_TR_POC""")
#    psspy.movebrn(FaultBus,POC,r"""1""",DummyTxBus,r"""1""") #adds new bus for dummy transformer to connect to. This is the Dummy transformer that will be used for voltage angle change test etc.     
#    psspy.two_winding_data_5(DummyTxBus,POC,r"""1""",[1,DummyTxBus,1,0,0,0,33,0,DummyTxBus,0,1,0,1,1,1],[0.0, 0.000001, 100.0, 1.0,0.0,0.0, 1.0,0.0, 1.0, 1.0, 1.0, 1.0,0.0,0.0, 1.0, 1.0, 1.1, 0.9,0.0,0.0,0.0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"","")
#    #psspy.two_winding_chng_5(10,2,r"""1""",[_i,_i,_i,_i,_i,_i,3,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f, 1.0, 1.0,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)

    ierr, base = psspy.busdat(InfiniteBus,'BASE')
    psspy.bus_data_4(DummyTxBus,0,[1,1,1,1],[base, 1.0,0.0, 1.1, 0.9, 1.1, 0.9],r"""DUM_TR_INF""") # add dummy Tx next to the Tx instead of POC
    psspy.movebrn(FaultBus,InfiniteBus,r"""1""",DummyTxBus,r"""1""") #adds new bus for dummy transformer to connect to. This is the Dummy transformer that will be used for voltage angle change test etc.     
    psspy.two_winding_data_5(InfiniteBus,DummyTxBus,r"""1""",[1,InfiniteBus,1,0,0,0,33,0,InfiniteBus,0,1,0,1,1,1],[0.0, 0.000001, 100.0, 1.0,0.0,0.0, 1.0,0.0, 1.0, 1.0, 1.0, 1.0,0.0,0.0, 1.0, 1.0, 1.1, 0.9,0.0,0.0,0.0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"","")
    #psspy.two_winding_chng_5(10,2,r"""1""",[_i,_i,_i,_i,_i,_i,3,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f, 1.0, 1.0,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)


    test_convergence()
    GRID_Z=100.0/GRID_MVA #per unitisation of grid impedance on 100 MVA basis.
    GRID_R=GRID_Z/math.sqrt(1.0+math.pow(setpoint['X_R'],2.0))
    GRID_X=GRID_R*setpoint['X_R']
    #set correct grid impedance and initialise voltage and power flows #intended behaviour is to get defined power flow from every generator. 
    GRID_R1=(1.0-F_dist)*GRID_R
    GRID_X1=(1.0-F_dist)*GRID_X
    GRID_R2=(F_dist)*GRID_R
    GRID_X2=(F_dist)*GRID_X
#    

    
    #iteratively initialise loadflow
    # it must be specified for each portion of the plant how much active and reactive power it shoudl provide and in which location. 
    gen_list_main=[]
    for key in PSSEmodelDict.keys():
        if('Generator' in key):
            gen_list_main.append(PSSEmodelDict[key])

    #create list of generator components to be initialised using the setpoint info keys
    gen_list={}
    for key in setpoint.keys():
        if ('P_' in key):
            gen_name=key[2:]
            loc_ID=setpoint['LOC_'+gen_name]
            frombus=PSSEmodelDict[loc_ID[0:2]+'fromBus'+loc_ID[2:]]
            tobus=PSSEmodelDict[loc_ID[0:2]+'toBus'+loc_ID[2:]]
            gen_list[gen_name]={'P':setpoint[key], 'Q':setpoint['Q_'+gen_name], 'fromBus':frombus, 'toBus':tobus, 'genBus':setpoint['BUS_'+gen_name]}
            
    pass

    #disconnect offline machines
    offline_machines=[]
    if('offline_machines' in setpoint.keys()):
        if( (setpoint['offline_machines']!=None) and(setpoint['offline_machines']!='') ):
            offline_machines=ast.literal_eval(setpoint['offline_machines']) 
        
    for gen in gen_list_main:
        if(gen in offline_machines):
            psspy.machine_chng_2(gen,r"""1""",[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #switch off all generators that are not in service        
    
    #disconnect offline buses
    bTo = psspy.abrnint(sid=-1,owner=1,ties=3,flag=4,entry=1,string="TONUMBER")[1][0] 
    bId = psspy.abrnchar(sid=-1,owner=1,ties=3,flag=4,entry=1,string="ID")[1][0]
    bFrom = psspy.abrnint(sid=-1,owner=1,ties=3,flag=4,entry=1,string="FROMNUMBER")[1][0]
    offline_buses=[]
    if('disconnect_buses' in setpoint.keys()):
        if( (setpoint['disconnect_buses']!=None) and (setpoint['disconnect_buses']!='') ):
            offline_buses=ast.literal_eval(setpoint['disconnect_buses'])
        for bus in offline_buses:
            psspy.bus_chng_4(bus,0,[4,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s)   
            
        for branch_id in range(0, len(bId)):
            if( (bTo[branch_id] in offline_buses) or (bFrom[branch_id] in offline_buses) ):
                psspy.branch_chng_3(bFrom[branch_id],bTo[branch_id],bId[branch_id],[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
                psspy.two_winding_chng_5(bFrom[branch_id],bTo[branch_id],bId[branch_id],[0,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)

        test_convergence(tree=1)
        
    #enforce transformer tap if provided in config file, otherwise just let it adjust automatically
    
    tr_tap_info={}
    for tr_id in range(0,10):
        if ('TR'+str(tr_id)+'_from' in setpoint.keys()):
            if( (setpoint['TR'+str(tr_id)+'_from']!='') and (setpoint['TR'+str(tr_id)+'_to']!='') and (setpoint['TR'+str(tr_id)+'_tap']!='') ):
                tr_tap_info['TR'+str(tr_id)]={'from': setpoint['TR'+str(tr_id)+'_from'], 'to' :setpoint['TR'+str(tr_id)+'_to'], 'tap':setpoint['TR'+str(tr_id)+'_tap']}
    if(tr_tap_info!={}):
        for tr in tr_tap_info.keys():
            psspy.two_winding_chng_5(tr_tap_info[tr]['from'],tr_tap_info[tr]['to'],r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f, float(tr_tap_info[tr]['tap']), float(tr_tap_info[tr]['tap']),_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)

    
    #iterate over all generators in list, start with 0 output, check different to target, increase output in proportion to the difference (greedy)
    QMAX= ProjectDetailsDict['GenMVArMax']
    QMIN= ProjectDetailsDict['GenMVArMin']
    for gen in  gen_list.keys():
        if(not (gen_list[gen]['genBus'] in offline_machines)):
            #set output to  and switch gen on
#            psspy.plant_data_3(gen_list[gen]['genBus'],0,_i,[ setpoint['V_POC'],_f]) #Update the setpoint voltage of the generator - may not needed as the Q value is fixed with the loop below
            ############################################################    
#            # 01/9/2022: Initialise droop characteristic - Lancaster only: Update voltage setpoint base on the actual voltage and Q at POC
#            droop_value = 3.3 #%
#            droop_base = 31.6
#            vol_deadband = 0  
#            Vspnt = implement_droop_LF(droop_value, droop_base, vol_deadband, setpoint['V_POC'], setpoint['Q'])
#            psspy.plant_data_3(gen_list[gen]['genBus'],0,_i,[ Vspnt,_f]) #Update the setpoint voltage of the generator - for setpoint variable initialisation
            ############################################################ 
            
            psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[1,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])        
            test_convergence()
            # QMAX=psspy.macdat(gen_list[gen]['genBus'], '1', 'QMAX')[1]
            # # When the Generator is set at pf control mode, absorbing Q, then QMax = Qmin = fix value -> Q does not change with conditions from line 563 to line 570.
            # # -> may need to input Qmin Qmax from excel spreadsheet
            # if QMAX <0: # If Qmax is negative, then change the sign of the value
            #     QMAX = -QMAX 
            # QMIN=psspy.macdat(gen_list[gen]['genBus'], '1', 'QMIN')[1]

            PMAX=psspy.macdat(gen_list[gen]['genBus'], '1', 'PMAX')[1]
            PMIN=psspy.macdat(gen_list[gen]['genBus'], '1', 'PMIN')[1]
            psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ 0.0,_f, 0.0,0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            
            test_convergence()
            powers=psspy.brnflo(gen_list[gen]['fromBus'], gen_list[gen]['toBus'], '1')
            P_meas=powers[1].real
            Q_meas=powers[1].imag
            max_err=0.02 #0.05
            P_prev=0 
            Q_prev=0
            max_iter=50
            iter_cnt=0
            conv_coeff=0.95
            while ( ( (abs(P_meas-gen_list[gen]['P']) > max_err) or (abs(Q_meas-gen_list[gen]['Q']) > max_err) ) and (iter_cnt<max_iter) ) :
                P_err=gen_list[gen]['P']-P_meas
                Q_err=gen_list[gen]['Q']-Q_meas
                P_set=P_prev+conv_coeff*(P_err)
                Q_set=Q_prev+conv_coeff*(Q_err)
                P_prev=P_set
                Q_prev=Q_set
                
                if(P_set>PMAX):
                    P_set=PMAX
                elif(P_set<PMIN):
                    P_set=PMIN                
                if(Q_set>QMAX):
                    Q_set=QMAX
                elif(Q_set<QMIN):
                    Q_set=QMIN            
                
                psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ P_set,_f, Q_set,Q_set,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                
                test_convergence()
                test_convergence()
#                test_convergence()
                powers=psspy.brnflo(gen_list[gen]['fromBus'], gen_list[gen]['toBus'], '1')
                P_meas=powers[1].real
                Q_meas=powers[1].imag
                
                iter_cnt+=1
                
                #run load flow
                #measure power flow 
                #adjust gen_ouput (P and Q simultaneously but independently)
##############################
    # 22/8/2022: add another loop of power initialisation to minimise the differences between the two idendical branches if applicable.
    for gen in  gen_list.keys():
        if(not (gen_list[gen]['genBus'] in offline_machines)):
            #set output to  and switch gen on
#            psspy.plant_data_3(gen_list[gen]['genBus'],0,_i,[ setpoint['V_POC'],_f]) #Update the setpoint voltage of the generator - may not needed as the Q value is fixed with the loop below
#            psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[1,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])        
#            test_convergence()
            # QMAX=psspy.macdat(gen_list[gen]['genBus'], '1', 'QMAX')[1]
            # # When the Generator is set at pf control mode, absorbing Q, then QMax = Qmin = fix value -> Q does not change with conditions from line 563 to line 570.
            # # -> may need to input Qmin Qmax from excel spreadsheet
            # if QMAX <0: # If Qmax is negative, then change the sign of the value
            #     QMAX = -QMAX 
            # QMIN=psspy.macdat(gen_list[gen]['genBus'], '1', 'QMIN')[1]

#            PMAX=psspy.macdat(gen_list[gen]['genBus'], '1', 'PMAX')[1]
#            PMIN=psspy.macdat(gen_list[gen]['genBus'], '1', 'PMIN')[1]
#            psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ 0.0,_f, 0.0,0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            
            test_convergence()
            powers=psspy.brnflo(gen_list[gen]['fromBus'], gen_list[gen]['toBus'], '1')
            P_meas=powers[1].real
            Q_meas=powers[1].imag
            max_err=0.001 #0.05
#            P_prev=0 
#            Q_prev=0
            max_iter=100 #50
            iter_cnt=0
            conv_coeff=0.95
            while ( ( (abs(P_meas-gen_list[gen]['P']) > max_err) or (abs(Q_meas-gen_list[gen]['Q']) > max_err) ) and (iter_cnt<max_iter) ) :
                P_err=gen_list[gen]['P']-P_meas
                Q_err=gen_list[gen]['Q']-Q_meas
                P_set=P_prev+conv_coeff*(P_err)
                Q_set=Q_prev+conv_coeff*(Q_err)
                P_prev=P_set
                Q_prev=Q_set
                
                if(P_set>PMAX):
                    P_set=PMAX
                elif(P_set<PMIN):
                    P_set=PMIN                
                if(Q_set>QMAX):
                    Q_set=QMAX
                elif(Q_set<QMIN):
                    Q_set=QMIN            
                
                psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ P_set,_f, Q_set,Q_set,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                
                test_convergence()
                test_convergence()
#                test_convergence()
                powers=psspy.brnflo(gen_list[gen]['fromBus'], gen_list[gen]['toBus'], '1')
                P_meas=powers[1].real
                Q_meas=powers[1].imag
                
                iter_cnt+=1
                
                #run load flow
                #measure power flow 
                #adjust gen_ouput (P and Q simultaneously but independently)

#########################                      
                
                
#    #after load flows are adjusted, add grid impedance and adjust voltage of infinite source
#    psspy.branch_chng_3(InfiniteBus,FaultBus,r"""1""",[_i,_i,_i,_i,_i,_i],[ GRID_R1, GRID_X1,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
#    #impedance between fautl bus and dummy bus for dummy transformer that has been automatically added.
#    psspy.branch_chng_3(FaultBus,DummyTxBus,r"""1""",[_i,_i,_i,_i,_i,_i],[ GRID_R2, GRID_X2,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)  
#    psspy.plant_data_3(InfiniteBus,0,_i,[ Vth,_f])
    
    #after load flows are adjusted, add grid impedance and adjust voltage of infinite source
    # Note that transformer is inserted between infbus and faultbus -> grid impedance will be from DummyTxBus to FaultBus
    psspy.branch_chng_3(DummyTxBus,FaultBus,r"""1""",[_i,_i,_i,_i,_i,_i],[ GRID_R1, GRID_X1,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
    #impedance between fautl bus and dummy bus for dummy transformer that has been automatically added.
    psspy.branch_chng_3(FaultBus,POC,r"""1""",[_i,_i,_i,_i,_i,_i],[ GRID_R2, GRID_X2,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)  
    psspy.plant_data_3(InfiniteBus,0,_i,[ Vth,_f])
    
    test_convergence()
    
    #psspy.machine_chng_2(322821,r"""1""",[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    #read voltage here for debugging

#########################      
    #check voltage at POC. if target is not hit, reiterate infinite bus voltage to hit. 
    v_poc_check=get_bus_info(POC, 'PU')[POC]['PU']
#    while abs(setpoint['V_POC']-v_poc_check)>0.001:
    while abs(setpoint['V_POC']-v_poc_check)>0.000002:
        delta_v=setpoint['V_POC']-v_poc_check
        Vth=Vth+delta_v
        psspy.plant_data_3(InfiniteBus,0,_i,[ Vth,_f])
        test_convergence()
        v_poc_check=get_bus_info(POC, 'PU')[POC]['PU']
        
#    # 24/5/2022: Update voltage setpoint of generators, and set back the Qmax Qmin values to each Gen -> both generators generate same amout of P and Q:
#    for gen in  gen_list.keys():
#        if(not (gen_list[gen]['genBus'] in offline_machines)):    
#            psspy.plant_data_3(gen_list[gen]['genBus'],0,_i,[ setpoint['V_POC'],_f])
#            psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ _f,_f, QMAX,QMIN,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#    test_convergence()

##############################
    # 29/11/2022: add another loop of power initialisation to minimise the differences between the two idendical branches if applicable.
    for gen in  gen_list.keys():
        if(not (gen_list[gen]['genBus'] in offline_machines)):
            #set output to  and switch gen on
#            psspy.plant_data_3(gen_list[gen]['genBus'],0,_i,[ setpoint['V_POC'],_f]) #Update the setpoint voltage of the generator - may not needed as the Q value is fixed with the loop below
#            psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[1,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])        
#            test_convergence()
            # QMAX=psspy.macdat(gen_list[gen]['genBus'], '1', 'QMAX')[1]
            # # When the Generator is set at pf control mode, absorbing Q, then QMax = Qmin = fix value -> Q does not change with conditions from line 563 to line 570.
            # # -> may need to input Qmin Qmax from excel spreadsheet
            # if QMAX <0: # If Qmax is negative, then change the sign of the value
            #     QMAX = -QMAX 
            # QMIN=psspy.macdat(gen_list[gen]['genBus'], '1', 'QMIN')[1]

#            PMAX=psspy.macdat(gen_list[gen]['genBus'], '1', 'PMAX')[1]
#            PMIN=psspy.macdat(gen_list[gen]['genBus'], '1', 'PMIN')[1]
#            psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ 0.0,_f, 0.0,0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            
            test_convergence()
            powers=psspy.brnflo(gen_list[gen]['fromBus'], gen_list[gen]['toBus'], '1')
            P_meas=powers[1].real
            Q_meas=powers[1].imag
            max_err=0.001 #0.05
#            P_prev=0 
#            Q_prev=0
            max_iter=100 #50
            iter_cnt=0
            conv_coeff=0.95
            while ( ( (abs(P_meas-gen_list[gen]['P']) > max_err) or (abs(Q_meas-gen_list[gen]['Q']) > max_err) ) and (iter_cnt<max_iter) ) :
                P_err=gen_list[gen]['P']-P_meas
                Q_err=gen_list[gen]['Q']-Q_meas
                P_set=P_prev+conv_coeff*(P_err)
                Q_set=Q_prev+conv_coeff*(Q_err)
                P_prev=P_set
                Q_prev=Q_set
                
                if(P_set>PMAX):
                    P_set=PMAX
                elif(P_set<PMIN):
                    P_set=PMIN                
                if(Q_set>QMAX):
                    Q_set=QMAX
                elif(Q_set<QMIN):
                    Q_set=QMIN            
                
                psspy.machine_chng_2(gen_list[gen]['genBus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ P_set,_f, Q_set,Q_set,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                
                test_convergence()
#                test_convergence()
#                test_convergence()
                powers=psspy.brnflo(gen_list[gen]['fromBus'], gen_list[gen]['toBus'], '1')
                P_meas=powers[1].real
                Q_meas=powers[1].imag
                print Q_meas
                
                iter_cnt+=1
                
                #run load flow
                #measure power flow 
                #adjust gen_ouput (P and Q simultaneously but independently)
                
#    v_poc_check=get_bus_info(POC, 'PU')[POC]['PU']            
#    while abs(setpoint['V_POC']-v_poc_check)>0.000002:
#        delta_v=setpoint['V_POC']-v_poc_check
#        Vth=Vth+delta_v
#        psspy.plant_data_3(InfiniteBus,0,_i,[ Vth,_f])
#        test_convergence()
#        v_poc_check=get_bus_info(POC, 'PU')[POC]['PU']                
##############################
                
    #retrieve angle at POC - or at FaultBus
    psspy.bsys(sid=0,numbus=1, buses=FaultBus)
    POC_info=psspy.abusreal(0,1,['BASE', 'ANGLE'])[1]
#    Vbase_POC=POC_info[0][0]
    ang=POC_info[1][0]        
    
    psspy.save(workspace_folder+"\\loadflow_after_init.sav")
    #check=get_bus_info(322800, 'PU')
    
    #add entries to standalon_script
    standalone_script+="psspy.case('loadflow_after_init.sav')\n" # no need to include initialisation commands, given initialised case is saved separately already
    standalone_script+='psspy.base_frequency(50.0)\n'
    
    return 0, Vth, ang #add criterio to return different value if initialisaiton not successful.           


def calc_fault_impedance(Vresidual, Fault_X_R, Vpoc, Vbase, MVAbase, grid_SCR, Grid_MVA, grid_X_R):
    # calc grid resistance and reactance in Ohms
    if(abs(Vresidual)<0.0001):
        Vresidual=0.0001
    Zbase=(Vbase*Vbase*1000**2)/(MVAbase*1000000) # Zbase assuming MVA base is provided in MVA and Vbase in kV
    Zgrid_abs=1.0/grid_SCR*Zbase    
    
#    Zfault=Zgrid_abs/((Vpoc/Vresidual)-1) #calcualte Zfault based on voltage divider
#    X_R_fault=Fault_X_R    
#    return Zfault, X_R_fault
    ANG_Zgrid=math.atan(grid_X_R)
    ANG_Zfault=math.atan(Fault_X_R)
    phi=ANG_Zgrid-ANG_Zfault
    Zfault=Zgrid_abs/(-1*math.cos(phi)+math.sqrt(math.pow(math.cos(phi),2)+(1/(math.pow((Vresidual/Vpoc),2)))-1))    
    X_R_fault=Fault_X_R    
    
    return Zfault, X_R_fault

#def calc_fault_impedance(Vresidual, Fault_X_R, Vth, ang, Vbase, MVAbase, grid_SCR, Grid_MVA, grid_X_R):
#    # calc grid resistance and reactance in Ohms
#    if(abs(Vresidual)<0.0001):
#        Vresidual=0.0001
#    Zbase=(Vbase*Vbase*1000**2)/(MVAbase*1000000) # Zbase assuming MVA base is provided in MVA and Vbase in kV
#    Zgrid_abs=1.0/grid_SCR*Zbase    
#    
#    R=Zgrid_abs/math.sqrt((1+grid_X_R*grid_X_R))
#    X=R*grid_X_R*1j
#    ZgridC=R+X
#    phi=ang #voltage angle between POC and INF bus
#    ZfaultC=ZgridC * (Vresidual/((Vth*math.cos(phi) - Vresidual) - Vth*math.sin(phi)*1j))
#    Zfault = abs(ZfaultC)
#    X_R_fault=Fault_X_R
#    
#    return Zfault, X_R_fault
#init_dynamics
def initialise_dynamics(dyr_file, workspace_folder, PSSEmodelDict):
    global standalone_script
    InfiniteBus=PSSEmodelDict['InfiniteBus']
#    psspy.cong(0)
#    psspy.conl(0,1,1,[0,0],[ 100.0,0.0,0.0, 100.0])
#    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,0.0, 100.0])
#    psspy.conl(0,1,3,[0,0],[ 100.0,0.0,0.0, 100.0])
#    psspy.ordr(0)
#    psspy.fact()
#    psspy.tysl(0)
    
    psspy.fdns([0,0,0,1,1,0,99,0])
    standalone_script+='psspy.fdns([0,0,0,1,1,0,99,0])\n'
    standalone_script+="#Prepare for dynamic simulation\n"
    #read POC voltage again here for debugging
    check=get_bus_info(322800, 'PU')
    psspy.cong(0)
    standalone_script+='psspy.cong(0)\n'
    psspy.conl(0,1,1,[0,0],[ 100.0,0.0,0.0, 100.0])
    standalone_script+='psspy.conl(0,1,1,[0,0],[ 100.0,0.0,0.0, 100.0])\n'
    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,0.0, 100.0])
    standalone_script+='psspy.conl(0,1,2,[0,0],[ 100.0,0.0,0.0, 100.0])\n'
    psspy.conl(0,1,3,[0,0],[ 100.0,0.0,0.0, 100.0])
    standalone_script+='psspy.conl(0,1,3,[0,0],[ 100.0,0.0,0.0, 100.0])\n'
    psspy.fact()
    standalone_script+='psspy.fact()\n'
    psspy.tysl(0)
    standalone_script+='psspy.tysl(0)\n'
    #replace line in dyr file with entry required for ZinGen model
    #1 'USRMDL' 1 'ZINGEN' 1 0 1 0 1 1 1/
    psspy.addmodellibrary(workspace_folder+"\\dsusr_zingen.dll")
    standalone_script+="psspy.addmodellibrary('dsusr_zingen.dll')\n"
    for key in PSSEmodelDict.keys():
        if('dll' in key):
            psspy.addmodellibrary(workspace_folder+"\\"+PSSEmodelDict[key])
            standalone_script+="psspy.addmodellibrary('"+str(PSSEmodelDict[key])+"')\n"
    dyr_handle=open(dyr_file, 'r')
    dyr_content=dyr_handle.read()
    dyr_handle.close()
    index=0
    while(dyr_content[index:index+1]!='\n'):
        index+=1
    new_dyr_content=dyr_content.replace(dyr_content[0:index], str(InfiniteBus) + " 'USRMDL' 1 'ZINGEN' 1 0 1 0 1 1 1/") #add entry for Zingen file
    os.remove(dyr_file)
    dyr_handle=open(dyr_file, 'w+')
    dyr_handle.write(new_dyr_content)
    dyr_handle.close()
    
    
    
    psspy.dyre_new([1,1,1,1], dyr_file, "", "", "")
    standalone_script+="psspy.dyre_new([1,1,1,1],'"+str(os.path.basename(os.path.normpath(dyr_file)))+"', '','', '')\n"
    #psspy.dyre_new([1,1,1,1], dyr_file,r"""CONEC.FLX""",r"""CONET.FLX""",r"""COMPILE.BAT""")   

#def initialize_droop(droop_value=3.325, droop_base=31.6, setpoint):    
#def initialize_droop(droop_value, droop_base, vol_deadband, V_POC, Q_POC):
def implement_droop(droop_value, droop_base, vol_deadband, V_POC, Q_POC, PSSEmodelDict, ProfilesDict, scenario_params, event_queue):
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
    print Vspnt
#    Vspnt = 1.02

#    profile=ProfilesDict[scenario_params['Test profile']]
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
            if('var' in Vset_dict[Vset_inst].keys()):
                event_queue.append({'time':0, 'type':'var_change_abs', 'rel_id': Vset_dict[Vset_inst]['var'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':Vspnt})
            elif('con' in Vset_dict[Vset_inst].keys()): 
                event_queue.append({'time':0, 'type':'con_change_abs', 'rel_id': Vset_dict[Vset_inst]['con'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':Vspnt})
                
            #Up date the setpoint right at the beginning of the simulations
            pass
        else: #It means the setpoint that needs to be changed is in the PSS/E Vref vector
            event_queue.append({'time':0, 'type':'VREF_change_abs', 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':Vspnt})  
                        
    
#    L=psspy.windmind(event['bus'], event['mac'], event['model_type'], 'VAR')[1]
#    L=psspy.windmind(367176, '1', 'WAUX', 'VAR')[1]
#    psspy.change_wnmod_var(367176, '1', 'EMSPCI2_1', L+6, Vspnt) #Update votlage setpoint at initialisation -> make sure extreme Q cases will stay.
    
    
#    return Vspnt
    pass

def implement_Pavai(Pavai, PSSEmodelDict, ProfilesDict, scenario_params, event_queue):
    
#    Pavai = Pavai/1000
     
    Pprim_params=[]
    for key in PSSEmodelDict.keys():
        if('Pprim' in key):
            Pprim_params.append(key)
    Pprim_cnt=1
    Pprim_dict={}
    while ( any( 'Pprim'+str(Pprim_cnt) in key for key in Pprim_params)):
        Pprim_dict[Pprim_cnt]={}
        for param_cnt in range(0, len(Pprim_params)):                
            param=Pprim_params[param_cnt]
            if('Pprim'+str(Pprim_cnt) in param):
                Pprim_dict[Pprim_cnt][param.replace('Pprim'+str(Pprim_cnt)+'_', '')]=PSSEmodelDict[param]                     
        Pprim_cnt+=1
    for Pprim_inst in Pprim_dict.keys(): #iterate over all instances of voltage setpoints requiring to be changed (e.g. across different machines/control systems)


        if('model' in Pprim_dict[Pprim_inst].keys()): #It means the setpoint that needs to be changes is a variable or constant
            if('var' in Pprim_dict[Pprim_inst].keys()):
                event_queue.append({'time':0, 'type':'var_change_abs', 'rel_id': Pprim_dict[Pprim_inst]['var'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':Pavai})
            elif('con' in Pprim_dict[Pprim_inst].keys()): 
                event_queue.append({'time':0, 'type':'con_change_abs', 'rel_id': Pprim_dict[Pprim_inst]['con'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':Pavai})

        
##        if('var' in Pprim_dict[Pprim_inst].keys()): #It means the setpoint that needs to be changes is a variable
##            event_queue.append({'time':0, 'type':'con_change_abs', 'rel_id': Pprim_dict[Pprim_inst]['var'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':Pavai})
##
##        elif('con' in Pprim_dict[Pprim_inst].keys()):    
##            event_queue.append({'time':0, 'type':'con_change_abs', 'rel_id': Pprim_dict[Pprim_inst]['con'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':Pavai})
#        # -> not applied in LSF as the initial irradiance level is set at Con J+82
#        Pprim_dict[Pprim_inst]['con'] = 182
#        event_queue.append({'time':0, 'type':'con_change_abs', 'rel_id': Pprim_dict[Pprim_inst]['con'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':Pavai})
 
    
#    return Pavai
    pass
    
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

    

def set_profile(path, profile):
    zingen_file=open(path+'.dat', "w+")       
    for row_cnt in range(0,len(profile)-1):
        row=profile[row_cnt]
        zingen_file.write(str(row[0])+'\t'+str(row[1])+'\t'+str(row[2])+'\n')
    row=profile[-1]
    zingen_file.write(str(row[0])+'\t'+str(row[1])+'\t'+str(row[2]))
        
    zingen_file.close()

def save_test_description(testInfoDir, scenario, scenario_params, setpoint_params, ProfilesDict):
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
    pass

def run(OutputDir, scenario, scenario_params, workspace_folder, testRun_, ProjectDetailsDict, PSSEmodelDict, SetpointsDict, ProfilesDict):
    global standalone_script
    standalone_script="# This script has been auto-generated by the PSSE SMIB test tool v2.0 to allow for debugging of individual test cases.\n# Please leave folder configuration unchanged and run simulation by loading this script in the PSSE GUI to re-run the given test.\n# Depending on your PSS/E version you may need to link to different .dll files.\n"                                                                                                                                                                                                                                                                          
    os.environ['PATH'] += ';' + workspace_folder # THIS IS THE MOST IMPORTANT LINE IN THE WHOLE DAMN SCRIPT!!!!!
    os.chdir(workspace_folder)
    event_queue=[]
    #detect what type of test shall be carried out and set up scenario (create required test profile(s) etc.)
    P_POC=SetpointsDict[scenario_params['setpoint ID']]['P'] #scaled, because values provided in p.u. on Plant MW base
    Q_POC=SetpointsDict[scenario_params['setpoint ID']]['Q'] #scaled, because values provided in p.u. on Plant MW base
    V_POC=SetpointsDict[scenario_params['setpoint ID']]['V_POC']
    Pavai=SetpointsDict[scenario_params['setpoint ID']]['Pavai']
    GRID_MVA=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'] #SCR is expresased on Plant MW base, so needs scaling
    GRID_X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R']
    InfiniteBus=PSSEmodelDict['InfiniteBus']
    FaultBus=PSSEmodelDict['FaultBus']
    DummyTxBus=PSSEmodelDict['DummyTxBus']
    POC=PSSEmodelDict['POC']    
    setpoint_params=SetpointsDict[scenario_params['setpoint ID']]
    start_offset=PSSEmodelDict['start_offset'] # added start_offset to delay the start of dynamic simulation 20/1/2022                   
    
    if('F_dist' in scenario_params.keys()):
        F_dist=scenario_params['F_dist']
    else:
        F_dist=0.0001 #by default fault is assumed to be located at the point of connection
        
#    InfiniteBus=1
    
    psspy.case(workspace_folder+"\\"+PSSEmodelDict['savFileName'])  
    psspy.base_frequency(50.0)
#    psspy.set_netfrq(1) # Activate the frequency dependence
    #retrieve base voltage of Infinite bus
    psspy.bsys(sid=0,numbus=1, buses=InfiniteBus)
    POC_info=psspy.abusreal(0,1,['BASE', 'ANGLE'])[1]
    Vbase_POC=POC_info[0][0]
    ANG_POC=POC_info[1][0]
    
    calculated_SCR=SetpointsDict[scenario_params['setpoint ID']]['GridMVA']/ProjectDetailsDict['SCRMVA']
    #Vth=Vth_initialisation.calc_Vth_pu(GRID_X_R, calculated_SCR, ProjectDetailsDict['TotalMVA'], Vbase_POC, Q_POC, P_POC, V_POC)
    Vth, ang = Vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])

    #Vth=1.019
    
#    status, Vth=initialise_loadflow(workspace_folder, ProjectDetailsDict, PSSEmodelDict, SetpointsDict[scenario_params['setpoint ID']], Vbase_POC, Vth, ANG_POC, GRID_MVA, InfiniteBus, FaultBus, DummyTxBus, POC, F_dist)# initialise load flow --> initialise, P, Q and initial voltage at POC
    status, Vth, ang=initialise_loadflow(workspace_folder, ProjectDetailsDict, PSSEmodelDict, SetpointsDict[scenario_params['setpoint ID']], Vbase_POC, Vth, ANG_POC, GRID_MVA, InfiniteBus, FaultBus, DummyTxBus, POC, F_dist)# initialise load flow --> initialise, P, Q and initial voltage at POC
    #add empty profile file for Zingen machine that does nothing
    Vzin=Vbase_POC*Vth/math.sqrt(3)
    zingen_default= [[-0.004,	50.0000,	Vzin],
                     [-0.002,	50.0000,	Vzin],
                     [0.000,	50.0000,	Vzin],
                     [100.000,	50.0000,	Vzin],
                     ]
    set_profile(workspace_folder+"\\ZINGEN1", zingen_default)   
    
    total_duration=float(PSSEmodelDict['default_sim_duration'])

    ###########################################################    
#    # 31/8/2022: Initialise droop characteristic - Lancaster only: Update voltage setpoint base on the actual voltage and Q at POC
#    droop_value = 3.3 #%
#    droop_base = 31.6
#    vol_deadband = 0  
#    implement_droop(droop_value, droop_base, vol_deadband, V_POC, Q_POC, PSSEmodelDict, ProfilesDict, scenario_params, event_queue)
    
    ############################################################

#    ############################################################    
    # 07/09/2022: Implement Pavai to initialise model at correct irradiance level
#    implement_Pavai(Pavai, PSSEmodelDict, ProfilesDict, scenario_params, event_queue)
#    ############################################################  
    
    #add event to limit availabel power if explicitly specified in setpoint
    if('avail_P'in SetpointsDict[scenario_params['setpoint ID']].keys()):
        if(SetpointsDict[scenario_params['setpoint ID']]['avail_P']!=''):
            #add
             event_queue.append({'time':0.01, 'type':'con_change_abs', 'rel_id': PSSEmodelDict['Pprim1_con'], 'model_type':PSSEmodelDict['Pprim1_type'], 'model':PSSEmodelDict['Pprim1_model'], 'bus':PSSEmodelDict['Pprim1_bus'], 'mac':str(PSSEmodelDict['Pprim1_mac']), 'value':SetpointsDict[scenario_params['setpoint ID']]['avail_P'] })
    
    #Add the events required to change SCR during runtime to event-queue
    if('small' in scenario):
        if( (scenario_params['Secondary SCL time']!='') and (scenario_params['Secondary SCL']!='') and (scenario_params['Secondary X_R']!='') ):
            #determine Vth_new and set parameters in Grid source accordingly
            Vth_prev, ang = Vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
            Vth_new, ang_new = Vth_initialisation.calc_Vth_pu(X_R=scenario_params['Secondary X_R'], SCC=scenario_params['Secondary SCL'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])           
            secGRID_Z=100.0/scenario_params['Secondary SCL'] #per unitisation of grid impedance on 100 MVA basis.
            secGRID_R=secGRID_Z/math.sqrt(1.0+math.pow(scenario_params['Secondary X_R'],2.0))
            secGRID_X=secGRID_R*scenario_params['Secondary X_R']
            secGRID_R1=(1.0-F_dist)*secGRID_R
            secGRID_X1=(1.0-F_dist)*secGRID_X
            secGRID_R2=(F_dist)*secGRID_R
            secGRID_X2=(F_dist)*secGRID_X
            event_queue.append({'time': scenario_params['Secondary SCL time'], 'type':'imp_change', 'fromBus':InfiniteBus, 'toBus':FaultBus, 'branchID':'1', 'X':secGRID_X1, 'R':secGRID_R1 }) #impedance change of branch between grid source and fault bus
            event_queue.append({'time': scenario_params['Secondary SCL time'], 'type':'imp_change', 'fromBus':FaultBus, 'toBus':DummyTxBus, 'branchID':'1', 'X':secGRID_X2, 'R':secGRID_R2 }) #impedance change of branch between grid and 
            event_queue.append({'time':scenario_params['Secondary SCL time'], 'type':'angle_change', 'POC':POC, 'angle':ang-ang_new})
    
    elif ( ('large' in scenario) or ('tov' in scenario) ):
        if((scenario_params['SCL_post']!='') and (scenario_params['X_R_post']!='')):
            switchTime=scenario_params['Ftime']+scenario_params['Fduration']#detemrin time at which fault ends (and SCR switches)
            Vth_prev, ang = Vth_initialisation.calc_Vth_pu(X_R=SetpointsDict[scenario_params['setpoint ID']]['X_R'], SCC=SetpointsDict[scenario_params['setpoint ID']]['GridMVA'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
            Vth_new, ang_new=Vth_initialisation.calc_Vth_pu(X_R=scenario_params['X_R_post'], SCC=scenario_params['SCL_post'], Vbase=ProjectDetailsDict['VPOCkv']*1000, Qpoc=SetpointsDict[scenario_params['setpoint ID']]['Q'], Ppoc=SetpointsDict[scenario_params['setpoint ID']]['P'], Vpoc=SetpointsDict[scenario_params['setpoint ID']]['V_POC'])
            secGRID_Z=100.0/scenario_params['SCL_post'] #per unitisation of grid impedance on 100 MVA basis.
            secGRID_R=secGRID_Z/math.sqrt(1.0+math.pow(scenario_params['X_R_post'],2.0))
            secGRID_X=secGRID_R*scenario_params['X_R_post']
            secGRID_R1=(1.0-F_dist)*secGRID_R
            secGRID_X1=(1.0-F_dist)*secGRID_X
            secGRID_R2=(F_dist)*secGRID_R
            secGRID_X2=(F_dist)*secGRID_X
            event_queue.append({'time': switchTime, 'type':'imp_change', 'fromBus':InfiniteBus, 'toBus':FaultBus, 'branchID':'1', 'X':secGRID_X1, 'R':secGRID_R1 }) #impedance change of branch between grid source and fault bus
            event_queue.append({'time': switchTime, 'type':'imp_change', 'fromBus':FaultBus, 'toBus':DummyTxBus, 'branchID':'1', 'X':secGRID_X2, 'R':secGRID_R2 }) #impedance change of branch between grid and 
            event_queue.append({'time': switchTime, 'type':'angle_change', 'POC':POC, 'angle':ang-ang_new})
    
    if (scenario_params['Test Type']=='F_profile'):        
        # write profile to zingen file
        zingen_profile=[[-0.004,	50.0000,	Vzin],
                     [-0.002,	50.0000,	Vzin],
                     ]
        profile=ProfilesDict[scenario_params['Test profile']]
        if(profile['scaling']=='relative'):
            for cnt in range(0, len(profile['x_data'])):
                zingen_profile.append([profile['x_data'][cnt], 50.0*float(profile['y_data'][cnt]), float(Vzin)*(float(profile['y_data'][cnt]))])
        else:
            for cnt in range(0, len(profile['x_data'])):
                zingen_profile.append([profile['x_data'][cnt], float(profile['y_data'][cnt]), float(Vzin)*(float(profile['y_data'][cnt])/50.0)])
        set_profile(workspace_folder+"\\ZINGEN1", zingen_profile)
        total_duration=profile['x_data'][-1]-0.1
        pass
        
    elif(scenario_params['Test Type']=='V_profile'):
        # write profile to zingen file 
        zingen_profile=[[-0.004,	50.0000,	Vzin],
                     [-0.002,	50.0000,	Vzin],
                     ]
        profile=ProfilesDict[scenario_params['Test profile']]
        if(profile['scaling']=='relative'):
            for cnt in range(0, len(profile['x_data'])):
                zingen_profile.append([profile['x_data'][cnt], 50.0, float(Vzin)*float(profile['y_data'][cnt])])
        else:
            for cnt in range(0, len(profile['x_data'])):
                zingen_profile.append([profile['x_data'][cnt], 50.0, float(Vbase_POC/math.sqrt(3))*float(profile['y_data'][cnt])])
        set_profile(workspace_folder+"\\ZINGEN1", zingen_profile)
        total_duration=profile['x_data'][-1]-0.1
        
    elif(scenario_params['Test Type']=='ANG_profile'):
        profile=ProfilesDict[scenario_params['Test profile']]
        for cnt in range(0, len(profile['x_data'])):            
            event_queue.append({'time':profile['x_data'][cnt], 'type':'angle_change', 'POC':POC, 'angle':profile['y_data'][cnt]})
        #write angle changes to event queue (change of transformer angle)
        #setpoint changes are internal events. The specifics of this can depend on the model used. Best would be to have compatibility with various models built in and just specify model type in scenario spreadsheet
        total_duration=event_queue[-1]['time']+5
    
    #---------------------SETPOINT CHANGE PROFILES (CAN BE MODEL-SPECIFIC)-----
    
    elif(scenario_params['Test Type']== 'V_stp_profile'):
        profile=ProfilesDict[scenario_params['Test profile']]
        scaling_factor=profile['scaling_factor_PSSE']
        if(not is_number(scaling_factor)):
            scaling_factor=1.0
        offset=profile['offset_PSSE']
        if(not is_number(offset)):
            offset=0.0   
        profile=interpolate(profile=profile, TimeStep=scenario_params['TimeStep']*0.001, density=20.0, scaling=scaling_factor, offset=offset )         
        #write changes to event queue --> specify in scenario spreadsheet, which VAR should change
#        L = psspy.mdlind(322813, '1 ', 'EXC', 'VAR')[1]
#        event_queue.append({'time':5.0, 'type':'var_change_abs', 'rel_id': 0,'model':'SMAHYC25', ' bus':322813, 'mac':'1', 'value':1.07})
 #       event_queue.append({'time':5.0, 'type':'var_change_abs', 'rel_id': 0,'abs_id':L+16, 'model':'SMAHYC25', ' bus':322813, 'mac':'1', 'value':1.07}) #adjust PPC voltage setpoint
#        event_queue.append({'time':5.0, 'type':'change_VREF', 'bus':32280, 'mac':'1', 'value':1.07}) #adjust SynCon voltage setpoint.
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
                            event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_rel', 'rel_id': Vset_dict[Vset_inst]['var'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                    elif(profile['scaling']=='absolute'):
                        for cnt in range(0, len(profile['x_data'])):
                            event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_abs', 'rel_id': Vset_dict[Vset_inst]['var'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif('con' in Vset_dict[Vset_inst].keys()): 
                    if(profile['scaling']=='relative'):
                        for cnt in range(0, len(profile['x_data'])):
                            event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_rel', 'rel_id': Vset_dict[Vset_inst]['con'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                    elif(profile['scaling']=='absolute'):
                        for cnt in range(0, len(profile['x_data'])):
                            event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_abs', 'rel_id': Vset_dict[Vset_inst]['con'], 'model_type':Vset_dict[Vset_inst]['type'], 'model':Vset_dict[Vset_inst]['model'], 'bus':Vset_dict[Vset_inst]['bus'], 'mac':str(Vset_dict[Vset_inst]['mac']), 'value':profile['y_data'][cnt]})
                    
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
        
    elif(scenario_params['Test Type']== 'Q_stp_profile'):   
        #scaling_factor=140.0/98.4# setpoint is defined in p.u. on Pltn MW base. However, the PSS/E model takes as input a p.u. setpoint expressed on MVA base.
        #'Total_MVA in the test definition sheet shoudl be udpapted to list total inver rating instead. 
        #write changes to event queue --> specify in scenario spreadsheet, which VAR should change
        #for SMA: change variable in HyCon per entry in spreadsheet. Interpolate profile based on points provided in profiles dict.
        profile=(ProfilesDict[scenario_params['Test profile']])
        scaling_factor=profile['scaling_factor_PSSE']
        if(not is_number(scaling_factor)):
            scaling_factor=1.0
        offset=profile['offset_PSSE']
        if(not is_number(offset)):
            offset=0.0      
        profile=interpolate(profile=profile, TimeStep=scenario_params['TimeStep']*0.001, density=20.0, scaling=scaling_factor, offset=offset )          
        Qset_params=[]
        for key in PSSEmodelDict.keys():
            if('Qset' in key):
                Qset_params.append(key)
        Qset_cnt=1
        Qset_dict={}
        while ( any( 'Qset'+str(Qset_cnt) in key for key in Qset_params)):
            Qset_dict[Qset_cnt]={}
            for param_cnt in range(0, len(Qset_params)):                
                param=Qset_params[param_cnt]
                if('Qset'+str(Qset_cnt) in param):
                    Qset_dict[Qset_cnt][param.replace('Qset'+str(Qset_cnt)+'_', '')]=PSSEmodelDict[param]                     
            Qset_cnt+=1
        for Qset_inst in Qset_dict.keys(): #iterate over all instances of voltage setpoints requiring to be changed (e.g. across different machines/control systems)
            
            if('var' in Qset_dict[Qset_inst].keys()): #It means the setpoint that needs to be changes is a variable
                #L = psspy.mdlind(322813, '1 ', 'EXC', 'VAR')[1]
                #L=psspy.mdlind(Vset_dict[Vset_inst]['bus'], Vset_dict[Vset_inst]['mac'], Vset_dict[Vset_inst]['type'], 'VAR')[1]
                
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_rel', 'rel_id': Qset_dict[Qset_inst]['var'], 'model_type':Qset_dict[Qset_inst]['type'], 'model':Qset_dict[Qset_inst]['model'], 'bus':Qset_dict[Qset_inst]['bus'], 'mac':str(Qset_dict[Qset_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_abs', 'rel_id': Qset_dict[Qset_inst]['var'], 'model_type':Qset_dict[Qset_inst]['type'], 'model':Qset_dict[Qset_inst]['model'], 'bus':Qset_dict[Qset_inst]['bus'], 'mac':str(Qset_dict[Qset_inst]['mac']), 'value':profile['y_data'][cnt]})
                       
            elif('con' in Qset_dict[Qset_inst].keys()):    
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_rel', 'rel_id': Qset_dict[Qset_inst]['con'], 'model_type':Qset_dict[Qset_inst]['type'], 'model':Qset_dict[Qset_inst]['model'], 'bus':Qset_dict[Qset_inst]['bus'], 'mac':str(Qset_dict[Qset_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_abs', 'rel_id': Qset_dict[Qset_inst]['con'], 'model_type':Qset_dict[Qset_inst]['type'], 'model':Qset_dict[Qset_inst]['model'], 'bus':Qset_dict[Qset_inst]['bus'], 'mac':str(Qset_dict[Qset_inst]['mac']), 'value':profile['y_data'][cnt]})
        
        event_queue=order_event_queue(event_queue)
        total_duration=event_queue[-1]['time']+5        
        pass
        
        pass
        
    elif(scenario_params['Test Type']==  'PF_stp_profile'):
        #write changes to event queue --> specify in scenario spreadsheet, which VAR should change
        profile=(ProfilesDict[scenario_params['Test profile']])
        scaling_factor=profile['scaling_factor_PSSE']
        if(not is_number(scaling_factor)):
            scaling_factor=1.0
        offset=profile['offset_PSSE']
        if(not is_number(offset)):
            offset=0.0      
        profile=interpolate(profile=profile, TimeStep=scenario_params['TimeStep']*0.001, density=20.0, scaling=scaling_factor, offset=offset )          
        PFsetset_params=[]
        for key in PSSEmodelDict.keys():
            if('PFset' in key):
                PFsetset_params.append(key)
        PFset_cnt=1
        PFset_dict={}
        while ( any( 'PFset'+str(PFset_cnt) in key for key in PFsetset_params)):
            PFset_dict[PFset_cnt]={}
            for param_cnt in range(0, len(PFsetset_params)):                
                param=PFsetset_params[param_cnt]
                if('PFset'+str(PFset_cnt) in param):
                    PFset_dict[PFset_cnt][param.replace('PFset'+str(PFset_cnt)+'_', '')]=PSSEmodelDict[param]                     
            PFset_cnt+=1
        for PFset_inst in PFset_dict.keys(): #iterate over all instances of voltage setpoints requiring to be changed (e.g. across different machines/control systems)
            
            if('var' in PFset_dict[PFset_inst].keys()): #It means the setpoint that needs to be changes is a variable
                #L = psspy.mdlind(322813, '1 ', 'EXC', 'VAR')[1]
                #L=psspy.mdlind(Vset_dict[Vset_inst]['bus'], Vset_dict[Vset_inst]['mac'], Vset_dict[Vset_inst]['type'], 'VAR')[1]
                
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_rel', 'rel_id': PFset_dict[PFset_inst]['var'], 'model_type':PFset_dict[PFset_inst]['type'], 'model':PFset_dict[PFset_inst]['model'], 'bus':PFset_dict[PFset_inst]['bus'], 'mac':str(PFset_dict[PFset_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_abs', 'rel_id': PFset_dict[PFset_inst]['var'], 'model_type':PFset_dict[PFset_inst]['type'], 'model':PFset_dict[PFset_inst]['model'], 'bus':PFset_dict[PFset_inst]['bus'], 'mac':str(PFset_dict[PFset_inst]['mac']), 'value':profile['y_data'][cnt]})
                       
            elif('con' in PFset_dict[PFset_inst].keys()):    
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_rel', 'rel_id': PFset_dict[PFset_inst]['con'], 'model_type':PFset_dict[PFset_inst]['type'], 'model':PFset_dict[PFset_inst]['model'], 'bus':PFset_dict[PFset_inst]['bus'], 'mac':str(PFset_dict[PFset_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_abs', 'rel_id': PFset_dict[PFset_inst]['con'], 'model_type':PFset_dict[PFset_inst]['type'], 'model':PFset_dict[PFset_inst]['model'], 'bus':PFset_dict[PFset_inst]['bus'], 'mac':str(PFset_dict[PFset_inst]['mac']), 'value':profile['y_data'][cnt]})
        
        event_queue=order_event_queue(event_queue)
        total_duration=event_queue[-1]['time']+5            
        pass
        
    elif(scenario_params['Test Type']== 'P_stp_profile'):
        #scaling_factor=float(ProjectDetailsDict['PlantMW'])/(ProjectDetailsDict['genPerSite1']*ProjectDetailsDict['genMVA1'])# setpoint is defined in p.u. on Pltn MW base. However, the PSS/E model takes as input a p.u. setpoint expressed on MVA base.
        #'Total_MVA in the test definition sheet shoudl be udpapted to list total inver rating instead. 
        #write changes to event queue --> specify in scenario spreadsheet, which VAR should change
        #for SMA: change variable in HyCon per entry in spreadsheet. Interpolate profile based on points provided in profiles dict.
        profile=(ProfilesDict[scenario_params['Test profile']])
        scaling_factor=profile['scaling_factor_PSSE']
        if(not is_number(scaling_factor)):
            scaling_factor=1.0
        offset=profile['offset_PSSE']
        if(not is_number(offset)):
            offset=0.0      
        profile=interpolate(profile=profile, TimeStep=scenario_params['TimeStep']*0.001, density=20.0, scaling=scaling_factor, offset=offset )          
        Pset_params=[]
        for key in PSSEmodelDict.keys():
            if('Pset' in key):
                Pset_params.append(key)
        Pset_cnt=1
        Pset_dict={}
        while ( any( 'Pset'+str(Pset_cnt) in key for key in Pset_params)):
            Pset_dict[Pset_cnt]={}
            for param_cnt in range(0, len(Pset_params)):                
                param=Pset_params[param_cnt]
                if('Pset'+str(Pset_cnt) in param):
                    Pset_dict[Pset_cnt][param.replace('Pset'+str(Pset_cnt)+'_', '')]=PSSEmodelDict[param]                     
            Pset_cnt+=1
        for Pset_inst in Pset_dict.keys(): #iterate over all instances of voltage setpoints requiring to be changed (e.g. across different machines/control systems)
            
            if('var' in Pset_dict[Pset_inst].keys()): #It means the setpoint that needs to be changes is a variable
                #L = psspy.mdlind(322813, '1 ', 'EXC', 'VAR')[1]
                #L=psspy.mdlind(Vset_dict[Vset_inst]['bus'], Vset_dict[Vset_inst]['mac'], Vset_dict[Vset_inst]['type'], 'VAR')[1]
                
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_rel', 'rel_id': Pset_dict[Pset_inst]['var'], 'model_type':Pset_dict[Pset_inst]['type'], 'model':Pset_dict[Pset_inst]['model'], 'bus':Pset_dict[Pset_inst]['bus'], 'mac':str(Pset_dict[Pset_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_abs', 'rel_id': Pset_dict[Pset_inst]['var'], 'model_type':Pset_dict[Pset_inst]['type'], 'model':Pset_dict[Pset_inst]['model'], 'bus':Pset_dict[Pset_inst]['bus'], 'mac':str(Pset_dict[Pset_inst]['mac']), 'value':profile['y_data'][cnt]})
                       
            elif('con' in Pset_dict[Pset_inst].keys()):    
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_rel', 'rel_id': Pset_dict[Pset_inst]['con'], 'model_type':Pset_dict[Pset_inst]['type'], 'model':Pset_dict[Pset_inst]['model'], 'bus':Pset_dict[Pset_inst]['bus'], 'mac':str(Pset_dict[Pset_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_abs', 'rel_id': Pset_dict[Pset_inst]['con'], 'model_type':Pset_dict[Pset_inst]['type'], 'model':Pset_dict[Pset_inst]['model'], 'bus':Pset_dict[Pset_inst]['bus'], 'mac':str(Pset_dict[Pset_inst]['mac']), 'value':profile['y_data'][cnt]})
        
        event_queue=order_event_queue(event_queue)
        total_duration=event_queue[-1]['time']+5        
        pass
    
    #Auxiliary profile used to alter Pprim
    elif(scenario_params['Test Type']== 'Auxiliary_profile'):
        #scaling_factor=0.001
        #write changes to event queue --> specify in scenario spreadsheet, which VAR should change. This is used here to alter the vailabel power for tests per DMAT guideline
        #for SMA: change constant per entry in test spreadsheet. Interpolate profile based on points provided in profiles dict.
        profile=(ProfilesDict[scenario_params['Test profile']])
        scaling_factor=profile['scaling_factor_PSSE']
        if(not is_number(scaling_factor)):
            scaling_factor=1.0
        offset=profile['offset_PSSE']
        if(not is_number(offset)):
            offset=0.0      
        profile=interpolate(profile=profile, TimeStep=scenario_params['TimeStep']*0.001, density=50.0, scaling=scaling_factor, offset=offset )       
        Pprim_params=[]
        for key in PSSEmodelDict.keys():
            if('Pprim' in key):
                Pprim_params.append(key)
        Pprim_cnt=1
        Pprim_dict={}
        while ( any( 'Pprim'+str(Pprim_cnt) in key for key in Pprim_params)):
            Pprim_dict[Pprim_cnt]={}
            for param_cnt in range(0, len(Pprim_params)):                
                param=Pprim_params[param_cnt]
                if('Pprim'+str(Pprim_cnt) in param):
                    Pprim_dict[Pprim_cnt][param.replace('Pprim'+str(Pprim_cnt)+'_', '')]=PSSEmodelDict[param]                     
            Pprim_cnt+=1
        for Pprim_inst in Pprim_dict.keys(): #iterate over all instances of voltage setpoints requiring to be changed (e.g. across different machines/control systems)
            
            if('var' in Pprim_dict[Pprim_inst].keys()): #It means the setpoint that needs to be changes is a variable
                #L = psspy.mdlind(322813, '1 ', 'EXC', 'VAR')[1]
                #L=psspy.mdlind(Vset_dict[Vset_inst]['bus'], Vset_dict[Vset_inst]['mac'], Vset_dict[Vset_inst]['type'], 'VAR')[1]
                
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_rel', 'rel_id': Pprim_dict[Pprim_inst]['var'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Prim_dict[Pprim_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'var_change_abs', 'rel_id': Pprim_dict[Pprim_inst]['var'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':profile['y_data'][cnt]})

            elif('con' in Pprim_dict[Pprim_inst].keys()):    
                if(profile['scaling']=='relative'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_rel', 'rel_id': Pprim_dict[Pprim_inst]['con'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':profile['y_data'][cnt]})
                elif(profile['scaling']=='absolute'):
                    for cnt in range(0, len(profile['x_data'])):
                        event_queue.append({'time':profile['x_data'][cnt], 'type':'con_change_abs', 'rel_id': Pprim_dict[Pprim_inst]['con'], 'model_type':Pprim_dict[Pprim_inst]['type'], 'model':Pprim_dict[Pprim_inst]['model'], 'bus':Pprim_dict[Pprim_inst]['bus'], 'mac':str(Pprim_dict[Pprim_inst]['mac']), 'value':profile['y_data'][cnt]})
        
        event_queue=order_event_queue(event_queue)
        total_duration=event_queue[-1]['time']+5        
        pass
    #--------------------FAULT SCENARIOS/TOV-----------------------------------
        
    elif(scenario_params['Test Type']=='Fault'):
        #add fault and and fault clearing to event queue
        Fault_X_R=scenario_params['Fault X_R']
        if(Fault_X_R==''):
            Fault_X_R=3.0
        if(scenario_params['F_Impedance']=='' and scenario_params['Vresidual']!=''):
            Zfault, Fault_X_R=calc_fault_impedance(float(scenario_params['Vresidual']), Fault_X_R, V_POC, Vbase_POC, ProjectDetailsDict['SCRMVA'], SetpointsDict[scenario_params['setpoint ID']]['SCR'],GRID_MVA, GRID_X_R)        #for more accurate calculation V_POC should be replaced with Vth, but that would also need to be done in PSCAD to align scenarios.    
#            Zfault, Fault_X_R=calc_fault_impedance(float(scenario_params['Vresidual']), Fault_X_R, Vth, ang, Vbase_POC, ProjectDetailsDict['SCRMVA'], SetpointsDict[scenario_params['setpoint ID']]['SCR'],GRID_MVA, GRID_X_R)        #for more accurate calculation V_POC should be replaced with Vth, but that would also need to be done in PSCAD to align scenarios. 
            Rfault=Zfault/math.sqrt(1+math.pow(Fault_X_R,2))
            Xfault=Rfault*Fault_X_R
            scenario_params['F_Impedance']=Zfault
        else:
            Zfault=scenario_params['F_Impedance']
                        
        Rfault=Zfault/math.sqrt(1+math.pow(Fault_X_R,2))
        Xfault=Rfault*Fault_X_R
        
        event_queue.append({'time':scenario_params['Ftime'], 'type':'apply_'+scenario_params['Ftype'], 'reactance':Xfault, 'resistance':Rfault})
        event_queue.append({'time':scenario_params['Ftime']+scenario_params['Fduration'], 'type':'clear_'+scenario_params['Ftype'] })            
        pass
        event_queue=order_event_queue(event_queue)
        total_duration=event_queue[-1]['time']+5
        
    elif(scenario_params['Test Type']=='Multifault'):
        #add all faults and fault clearing to event queue --> check PSCAD implementation onhow to access fault list provided by "readtestinfo" routine
        for fault_cnt in range(0, len(scenario_params['Fduration'])): #pick duration vector as reference to determine how many faults are included in list
            Fault_X_R=scenario_params['Fault X_R'][fault_cnt]
            if(Fault_X_R==''):
                Fault_X_R=3.0
            if(scenario_params['F_Impedance'][fault_cnt]=='' and scenario_params['Vresidual'][fault_cnt]!=''):
                Zfault, Fault_X_R=calc_fault_impedance(float(scenario_params['Vresidual'][fault_cnt]), Fault_X_R, V_POC, Vbase_POC, ProjectDetailsDict['SCRMVA'], SetpointsDict[scenario_params['setpoint ID']]['SCR'],GRID_MVA, GRID_X_R)        #for more accurate calculation V_POC should be replaced with Vth, but that would also need to be done in PSCAD to align scenarios.    
#                Zfault, Fault_X_R=calc_fault_impedance(float(scenario_params['Vresidual'][fault_cnt]), Fault_X_R, Vth, ang, Vbase_POC, ProjectDetailsDict['SCRMVA'], SetpointsDict[scenario_params['setpoint ID']]['SCR'],GRID_MVA, GRID_X_R)        #for more accurate calculation V_POC should be replaced with Vth, but that would also need to be done in PSCAD to align scenarios.
                Rfault=Zfault/math.sqrt(1+math.pow(Fault_X_R,2))
                Xfault=Rfault*Fault_X_R
                scenario_params['F_Impedance'][fault_cnt]=Zfault
            else:
                Zfault=scenario_params['F_Impedance'][fault_cnt]
                            
            Rfault=Zfault/math.sqrt(1+math.pow(Fault_X_R,2))
            Xfault=Rfault*Fault_X_R
            
            event_queue.append({'time':scenario_params['Ftime'][fault_cnt], 'type':'apply_'+scenario_params['Ftype'][fault_cnt], 'reactance':Xfault, 'resistance':Rfault})
            event_queue.append({'time':scenario_params['Ftime'][fault_cnt]+scenario_params['Fduration'][fault_cnt], 'type':'clear_'+scenario_params['Ftype'][fault_cnt] })    
        pass
        total_duration=event_queue[-1]['time']+5
    
    elif(scenario_params['Test Type']=='Multifault_random'):
        #add all faults and fault clearing to event queue --> use routine from PSCAD for random fault events
        random_times=[0.01, 0.01, 0.2, 0.2, 0.5, 0.5, 0.75, 1, 1.5, 2, 2, 3, 5, 7, 10]
        random.shuffle(random_times)
        random_fault_duration=[0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.22, 0.22, 0.22, 0.22, 0.22, 0.22, 0.43]
        random.shuffle(random_fault_duration)
        Zgrid=math.pow((ProjectDetailsDict['VbaseTestSrc']*1000),2)/setpoint_params['GridMVA']/1000000 #Zgrid
        random_impedances=[0, 0, 0.2, 0.2, 0.2, 1, 1, 1, 1, 1, 2, 2, 2, 3.5, 3.5]
        random.shuffle(random_impedances)
        random_fault_types=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        random.shuffle(random_fault_types)
        
        faultTimeOffset=scenario_params['Ftime']
        
        scenario_params['F_Impedance']=[]
        scenario_params['Fault X_R']=[setpoint_params['X_R']]
        scenario_params['Ftime']=[]
        scenario_params['Fduration']=[]
        scenario_params['Ftype']=['3PHG'] #only 3PHG fault in PSS/E random multifault per DMAT guideline
        for faultID in range(0, 15):     
            Zfault=random_impedances[faultID]*Zgrid            
            Fault_X_R=setpoint_params['X_R']            
            Rfault=Zfault/math.sqrt(1+math.pow(Fault_X_R,2))
            Xfault=Rfault*Fault_X_R             
            event_queue.append({'time':faultTimeOffset, 'type':'apply_3PHG', 'reactance':Xfault, 'resistance':Rfault})
            event_queue.append({'time':faultTimeOffset+random_fault_duration[faultID], 'type':'clear_3PHG' })
            faultTimeOffset=faultTimeOffset+random_fault_duration[faultID]+random_times[faultID]
            #write fault details back to scenario_params, so that it is correctly captured in the metadata
            scenario_params['F_Impedance'].append(Zfault)
            scenario_params['Ftime'].append(faultTimeOffset)
            scenario_params['Fduration'].append(random_fault_duration[faultID])
        total_duration=event_queue[-1]['time']+5
    
    elif(scenario_params['Test Type']=='TOV'):
        #add element to event queue to either alter setting of capacitive element at POC or add capacitive element at POC
        #                    
        if(is_number(scenario_params['Capacity(F)']) ):
            capacity=float(scenario_params['Capacity(F)'])
        else:
            Qinj, capacity = TOV_calc.calc_capacity(setpoint_params['GridMVA'], setpoint_params['X_R'], setpoint_params['P'], setpoint_params['Q'], setpoint_params['V_POC'], ProjectDetailsDict['VbaseTestSrc']*1000, scenario_params['Vresidual']) #Vbase must be provided in volts
            Qcap=Qinj/(scenario_params['Vresidual']*scenario_params['Vresidual'])
            scenario_params['Capacity(F)']=round(capacity, 4)
        psspy.shunt_data(FaultBus,r"""1""",0,[0.0,Qcap])
        standalone_script+="psspy.shunt_data("+str(FaultBus)+",'2',0,[0.0,"+str(Qinj)+'])'
        event_queue.append({'time':scenario_params['time'], 'type':'shunt_on', 'bus': FaultBus, 'id': '1'})#add event to switch shunt on
        event_queue.append({'time':scenario_params['time']+scenario_params['Fduration'], 'type':'shunt_off', 'bus': FaultBus, 'id': '1'})
        #add event to switch shunt off       
        total_duration=event_queue[-1]['time']+5
    
    dyr_to_use=PSSEmodelDict['dyrFileName']
    if('dyr' in SetpointsDict[scenario_params['setpoint ID']].keys()):
        alternative_dyr=SetpointsDict[scenario_params['setpoint ID']]['dyr']
        if(alternative_dyr!=''):
            dyr_to_use=alternative_dyr        
            
    initialise_dynamics(workspace_folder+"\\"+dyr_to_use, workspace_folder, PSSEmodelDict)
    #deactivate network frequency dependence
    psspy.set_netfrq(0)
#    psspy.set_netfrq(1) # Activate the frequency dependence
    
    # Tunning Parameters
    tune_parameters(PSSEmodelDict)
    
    #psspy.dynamics_solution_param_2([999,_i,_i,_i,_i,_i,_i,_i],[ 0.2,_f, 0.001,_f,_f,_f, 0.2,_f])
    if( (scenario_params['AccFactor']!=None) and (scenario_params['AccFactor']!='') ):
        AccFactor=scenario_params['AccFactor']
    if( (scenario_params['TimeStep']!=None) and (scenario_params['TimeStep']!='') ):
        TimeStep=scenario_params['TimeStep']
    # psspy.dynamics_solution_param_2([_i,_i,_i,_i,_i,_i,_i,_i],[AccFactor,_f, TimeStep/1000.0,_f,_f,_f,_f,_f])
    # standalone_script+="psspy.dynamics_solution_param_2([_i,_i,_i,_i,_i,_i,_i,_i],["+str(AccFactor)+",_f, "+str(TimeStep/1000.0)+",_f,_f,_f,_f,_f])\n"
    psspy.dynamics_solution_param_2([200,_i,_i,_i,_i,_i,_i,_i],[AccFactor,_f, TimeStep/1000.0,_f,_f,_f,_f,_f])
    standalone_script+="psspy.dynamics_solution_param_2([200,_i,_i,_i,_i,_i,_i,_i],["+str(AccFactor)+",_f, "+str(TimeStep/1000.0)+",_f,_f,_f,_f,_f])\n"
    psspy.set_relang(1,InfiniteBus,r"""1""")
    standalone_script+="psspy.set_relang(1,"+str(InfiniteBus)+",'1')\n"
    
    set_channels(PSSEmodelDict)  
    
    psspy.set_chnfil_type(0)
    standalone_script+="psspy.set_chnfil_type(0)"
    
    standalone_script+="#Initialise Dynamics\n"
    OUTPUT_name = PSSEmodelDict['savFileName'][0:-4]+'_'+scenario+'.out'
    
    psspy.lines_per_page_one_device(1,60)
    standalone_script+="psspy.lines_per_page_one_device(1,60)\n"
    #psspy.progress_output(2,r"""SAMPLE_SMA.LOG""",[0,0])
    psspy.progress_output(2,os.path.join(workspace_folder, OUTPUT_name)[0:-4]+".LOG",[0,0])
    standalone_script+="psspy.progress_output(2,'"+str(OUTPUT_name[0:-4]+".LOG")+"',[0,0])\n"
    
    psspy.lines_per_page_one_device(1,60)
    standalone_script+="psspy.lines_per_page_one_device(1,60)\n"
    
    psspy.strt(0, os.path.join(workspace_folder, OUTPUT_name))
    standalone_script+="psspy.strt(0, '"+str(OUTPUT_name)+"')\n"
        

    #psspy.strt(0,r"""SAMPLE_SMA.OUT""")
    #ADDITIONAL STUFF REQUIRED FOR CORRECT MODEL STARTUP
#    psspy.change_plmod_var(322813, '1', 'SMAHYC26', 21, 0.85366*0.05)
#    standalone_script+="psspy.change_plmod_var(322813, '1', 'SMAHYC26', 21, 0.85366*0.05)\n"
#    psspy.change_plmod_con(322813, '1', 'SMASC172', 1, 1.1)
    #standalone_script+="psspy.change_plmod_con(322813, '1', 'SMASC172', 1, 1.1)\n"

#    ############################################################    
#    # 31/8/2022: Initialise droop characteristic - Lancaster only: Update voltage setpoint base on the actual voltage and Q at POC
#    droop_value = 3.3 #%
#    droop_base = 31.6
#    vol_deadband = 0  
#    implement_droop(droop_value, droop_base, vol_deadband, V_POC, Q_POC, PSSEmodelDict, ProfilesDict, scenario_params)
#    
    
#    Vspnt = 1.02
##    L=psspy.windmind(367176, '1', 'WAUX', 'VAR')[1]
##    psspy.change_wnmod_var(367176, '1', 'EMSPCI2_1', L+6, Vspnt) #Update votlage setpoint at initialisation -> make sure extreme Q cases will stay.
#    psspy.change_wnmod_var(367176, '1', 'EMSPCI2_1', 7, Vspnt) #Update votlage setpoint at initialisation -> make sure extreme Q cases will stay.
#    ############################################################
    
#    ############################################################    
#    # 07/09/2022: Implement Pavai to initialise model at correct irradiance level
#    implement_Pavai(Pavai, PSSEmodelDict, ProfilesDict, scenario_params, event_queue)
#    ############################################################    
    
    standalone_script+="#Run simulation\n"
    while (len(event_queue)>0):
        ierr, event_queue = process_next_event(event_queue, start_offset, PSSEmodelDict)# added start_offset to delay the start of dynamic simulation 20/1/2022
    
    var_init_dict.clear() #clear var_init_dict before the next test.
    #run for X more seconds after last event     
    psspy.run(0,total_duration + start_offset, 500, 1, 0) #run until event occurs # added start_offset to delay the start of dynamic simulation 20/1/2022
    standalone_script+="psspy.run(0,"+str(total_duration + start_offset)+", 500, 1, 0)\n" # added start_offset to delay the start of dynamic simulation 20/1/2022  
#    psspy.run(0,total_duration + start_offset, 500, 7, 0) #run until event occurs # added start_offset to delay the start of dynamic simulation 20/1/2022
#    standalone_script+="psspy.run(0,"+str(total_duration + start_offset)+", 500, 7, 0)\n" # added start_offset to delay the start of dynamic simulation 20/1/2022                                                                                                                                                    
    csv_file=OutputDir+"\\"+testRun_+"\\"+scenario+"\\"+scenario+'_results'
    try:
            os.mkdir(OutputDir+"\\"+testRun_)
    except:
        print("testRun result folder already exists")
    else:
        print("testRun directory created")
        
    try:
            os.mkdir(OutputDir+"\\"+testRun_+"\\"+scenario)
    except:
        print("scenario folder already exists")
    else:
        print("scenario results folder created")
        
    #save the standalone script in model folder and name it with the scenario name
    standalone_script+="#Check the file '"+str(OUTPUT_name)+"' in the model folder for simulation results.\n"
    text_file = open(workspace_folder+'\\'+str(scenario)+".py", "w")
    text_file.write(standalone_script)
    text_file.close()
    
        
    conv.save_results(workspace_folder+"\\"+OUTPUT_name, csv_file)
    save_test_description(OutputDir+"\\"+testRun_+"\\"+scenario+"\\testInfo", scenario, scenario_params, SetpointsDict[scenario_params['setpoint ID']], ProfilesDict)# write test metadata to human-readable txt or csv file in same folder as results.

    
def main(OutputDir, scenario, scenario_params, workspace_folder, testRun_, ProjectDetailsDict, PSSEmodelDict, SetpointsDict, ProfilesDict):
    run(OutputDir, scenario, scenario_params, workspace_folder, testRun_, ProjectDetailsDict, PSSEmodelDict, SetpointsDict, ProfilesDict)
    
if __name__ == '__main__':
    main(r'C:\\Users\\Mervin Kall\\OneDrive - ESCO Pacific\\AutomatedTesting\\20201019_SMIB_testing_revised\\PSSE_sim',
         'large148', 
         {'Fduration': 0.43, 'run in PSS/E?': u'yes', 'run in PSCAD?': u'yes', 'Active Power': 1, 'Vpoc': 1.04, 'F_Impedance': u'', 'X_R_post': 3.76, 'Fault X_R': 3.0, 'Ftype': u'3PHG', 'Ftime': 4.0, 'Test Type': u'Fault', 'SCR_post': 2.4357142857142855, 'Reactive Power': 0, 'X_R': 3.76, 'Vresidual': 0.9, 'AccFactor': 1, 'TimeStep': 1, 'setpoint ID': 1, 'SCR': 2.4357142857142855},
         r'C:\\Users\\Mervin Kall\\OneDrive - ESCO Pacific\\AutomatedTesting\\20201019_SMIB_testing_revised\\PSSE_sim\\model_copies\\script_check\\large148\\HORSF_v0_2',
         'script_check',
         {u'Fbase': 50, u'Sub': u'HOTS 220 kV', u'Pmax': 1, u'genPerSite1': 41, u'VPOCkv': 220, u'genPerSite3': 0, u'genPerSite2': 0, u'PmaxQmin': 1, u'TotalMVA': 161.84, u'VbaseTestSrc': 220, u'State': u'', u'Type': u'', u'genMVA2': 0, u'genMVA1': 4, u'PmaxQmax': 1, u'SCRLow': 3, u'genMVA3': 0, u'PminQmax': 0, u'Abbr': u'', u'PlantMW': 138, u'VTERkv': 0.6, u'SCRHigh': 10, u'Town': u'', u'Name': u'Horsham', u'Gens': 1, u'Dev': u'ESCO', u'NSP': u'AEMO', u'PminQmin': 0, u'Pmin': 0, u'VPOCpu': 1, u'SCRMVA': 161.84, u'GenMW': 140},
         {u'MVmeasBus2': 322821, u'MVmeasBus1': 322811, u'POCmeasBus1': 32280, u'savFileName': u'HOR_v0_2.sav', u'LVtoBus1': 322812, u'HVmeasBus1': 322800, u'HVmeasBus2': 322800, u'dll1': u'PRISMIC_AC3100OEL_rev0.5.dll', u'FaultBus': 2, u'InfiniteBus': 1, u'dll2': u'PRISMIC_AC3100UEL_rev0.4.dll', u'dll3': u'SMPROT_R0.dll', u'dll4': u'SMASC_E170_SMAHYC_021206A05_F25_342_IVF111.dll', u'POCtoBus1': 32280, u'POC': 32280, u'HVtoBus2': 322800, u'HVtoBus1': 322800, u'LVmeasBus1': 322813, u'Generator2': 322821, u'Generator1': 322813, u'POCfromBus1': 322800, u'LVfromBus1': 322813, u'MVfromBus1': 322811, u'MVfromBus2': 322821, u'HVfromBus2': 322820, u'HVfromBus1': 322810, u'dyrFileName': u'HOR_v0_2.dyr', u'MVtoBus1': 322810, u'MVtoBus2': 322820},
         {},
         {},
            )