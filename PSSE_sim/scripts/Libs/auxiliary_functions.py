# -*- coding: utf-8 -*-
"""
Created on Thu Apr 29 10:39:15 2021

@author: PSCAD
28/6/2023: include rating2 RATE2 in branch monitor
"""

###############################################################################
# IMPORT PSSE
###############################################################################
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
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()


import dyntools

import numpy as np
#import other stuff
import pandas as pd
import datetime as dt

import multiprocessing as mp
from multiprocessing import Pool
import operator
###############################################################################
#Class definitions
###############################################################################
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

###############################################################################
#FUNCTION DEFINITIONS
###############################################################################

        
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
#                    print('bus number is'+str(bus_list))
                    bus_info[bus_list][param_list]=param_val[0][0]
    return bus_info

#return information on branches between the two provided buses. Only returns data in in-service branches
def get_branch_info(frombus, tobus, ckt='', include_offline=0):
    return_value=[]
    off_ass=include_offline+3
    if (type(tobus)==int):
        psspy.bsys(sid=0, numbus=2, buses=[frombus, tobus])
        ierr, branch_int_out=psspy.abrnint(0,1,1,off_ass,1,['FROMNUMBER', 'TONUMBER', 'STATUS', 'METERNUMBER'])
        ierr, branch_char_out=psspy.abrnchar(0,1,1,off_ass,1,['ID', 'FROMNAME', 'TONAME'])
        ierr, branch_data=psspy.abrnreal(0,1,1,off_ass,2, ['PCTMVARATE','PCTMVARATE1','PCTMVARATE2','PCTMVARATE3','RATE1','RATE2','P','Q','MVA'])
        #return_value = [ID, P, Q, MVA, PCTMVARATE, Rate1]
        number_of_lines=len(branch_data[0])/2
        for i in range (0, len(branch_data[0])/2):
            if(branch_int_out[3][i]==frombus):#if metered end is frombus
                index=i
            else: #metered end is tobus, look at the second half of the result array.
                index=i+number_of_lines
            return_value.append({'FROMBUS':branch_int_out[0][i], 'TOBUS':branch_int_out[1][i], 'ID':branch_char_out[0][i], 'P':branch_data[6][index], 'Q':branch_data[7][index], 'MVA':branch_data[8][index], 'PCTMVARATE':branch_data[0][index], 'RATING1':branch_data[4][index], 'RATING2':branch_data[5][index], 'STATUS':branch_int_out[2][i]})
    
    elif (type(tobus)==list):
        for tb_id in range(0, len(tobus)):
            psspy.bsys(sid=0, numbus=2, buses=[frombus, tobus[tb_id]])
            ierr, branch_int_out=psspy.abrnint(0,1,1,off_ass,1,['FROMNUMBER', 'TONUMBER', 'STATUS', 'METERNUMBER'])
            ierr, branch_char_out=psspy.abrnchar(0,1,1,off_ass,1,['ID', 'FROMNAME', 'TONAME'])
            #ierr, branch_data= psspy.abrnreal(0,1,1,3,1, ['PCTMVARATE','PCTMVARATE1','PCTMVARATE2','PCTMVARATE3','RATE1','P','Q','MVA'])
            ierr, branch_data=psspy.abrnreal(0,1,1,off_ass,2, ['PCTMVARATE','PCTMVARATE1','PCTMVARATE2','PCTMVARATE3','RATE1','RATE2','P','Q','MVA']) #the data from non-metered end is puyt at the end of the array. (first alll results from metered end and then all results from non-metered end)
            #return_value = [ID, P, Q, MVA, PCTMVARATE, Rate1]
            number_of_lines=len(branch_data[0])/2
            for i in range (0, len(branch_data[0])/2):
                if(branch_int_out[3][i]==frombus):#if metered end is frombus
                    index=i
                else: #metered end is tobus, loos at the second half of the result array.
                    index=i+number_of_lines
                return_value.append({'FROMBUS':branch_int_out[0][i], 'TOBUS':branch_int_out[1][i], 'ID':branch_char_out[0][i], 'P':branch_data[6][index], 'Q':branch_data[7][index], 'MVA':branch_data[8][index], 'PCTMVARATE':branch_data[0][index], 'RATING1':branch_data[4][index], 'RATING2':branch_data[5][index], 'STATUS':branch_int_out[2][i]})
                
    if(ckt==''):
        return return_value
    else:
        return_select=[]
        for i in range(0,len(return_value)):
            if int(return_value[i]['ID'])==int(ckt):
                return_select.append(return_value[i])
        return return_select

def create_eqv_load(frombus, tobus):
    P=0
    Q=0
    branch_data = get_branch_info(frombus, tobus)
    for brn_cnt in range (0, len(branch_data)):
        P+=branch_data[brn_cnt]['P']
        Q+=branch_data[brn_cnt]['Q']
    if(P!=0 or Q!=0):      
        psspy.load_data_5(frombus,r"""10""",[1,2,98,1,1,0,0],[P,Q,0.0,0.0,0.0,0.0,0.0,0.0])
    
def dscn_buses(buses):
    if(type(buses) == list):
        for i in range (0, len(buses)):
            psspy.dscn(buses[i])
    elif(type(buses)==int):
        psspy(dscn(buses))
        
def identify_slack_buses():
    slack_buses={}
    ierr, buses=psspy.abusint(-1,1,'Number')
    for i in range(0, len(buses[0])):
        ierr, bus_type = psspy.busint(buses[0][i],'TYPE')
        ierr, bus_zone = psspy.busint(buses[0][i],'AREA')
        if(bus_type==3):
            slack_buses[buses[0][i]]={}
            slack_buses[buses[0][i]]['Zone']=bus_zone
            
            Pmax=None
            mac_id=0
            while (Pmax==None) and (mac_id<10):
                mac_id+=1
                ierr, P=psspy.macdat(buses[0][i], str(mac_id), 'P') #assume slack bus machine id is always 1 later add function to detect connected machines
                ierr, Q=psspy.macdat(buses[0][i], str(mac_id), 'Q')
                ierr, Pmax=psspy.macdat(buses[0][i], str(mac_id), 'PMAX')
                ierr, Qmax=psspy.macdat(buses[0][i], str(mac_id), 'QMAX')
                ierr, Qmin=psspy.macdat(buses[0][i],str(mac_id), 'QMIN')
                
                slack_buses[buses[0][i]]['P']=P
                slack_buses[buses[0][i]]['Q']=Q
                slack_buses[buses[0][i]]['Pmax']=Pmax
                slack_buses[buses[0][i]]['Qmax']=Qmax
                slack_buses[buses[0][i]]['Qmin']=Qmin
               
    return slack_buses

#check if bus number exists in case
def bus_nr_exists(number):
    bus_nr_exists=False
    psspy.bsys(sid=0, numarea=6, areas=[2,3,4,5,6,7])
    bus_nrs = psspy.abusint(-1,2,'NUMBER') #should also be possible to create subsystem containing all buses by using -1 as subsystem identifier instead, and not defining a subsystem at all.
    for bus_nr in bus_nrs[1][0]:
        if(bus_nr == number):
            bus_nr_exists=True
    return bus_nr_exists

def branch_status(frombus, tobus, branch_id):
    ierr, status=psspy.brnint(frombus, tobus, branch_id, 'STATUS')
    if(ierr>0): return -1
    else: return status        

#add plant to snapshot via rdch and initialise the angles to get the case to converge        
def add_plant_rdch(raw_file, POC, buses, gens):
      #read voltage and angle at point where additional elements to be connected
    poc_info_pre=get_bus_info(POC,'ANGLED')  #angle from POC bus is taken as reference 
      #bus_info[32090]={'v_init':voltages[0][0])
    #ierr, angles= psspy.abusreal(1,1,'ANGLED')
    #bus_info[32090]={'ang_init':angles[0][0]}
    #angle = angles[0][0]
    plant_included = 0
    for bus in buses: # if any of the buses of the model to be included is missing in original file, include the whole model
        if (bus_nr_exists(bus)):
            plant_included = 0
        else:
            plant_included = 0
               
    if (plant_included==0):
    #add elements --> this changes PCO bus angle as well
        psspy.rdch(0,raw_file)
        #read new POC angle and determine delta. 
        poc_info_post=get_bus_info(POC,'ANGLED')
        angle_delta = poc_info_post[POC]['ANGLED']-poc_info_pre[POC]['ANGLED']
        psspy.bus_chng_4(POC,0,[_i,_i,_i,_i],[_f,_f,poc_info_pre[POC]['ANGLED'],_f,_f,_f,_f],_s)
        for bus in buses:
            bus_info=get_bus_info(bus, 'ANGLED')
            new_angle=bus_info[bus]['ANGLED']-angle_delta
            psspy.bus_chng_4(bus,0,[_i,_i,_i,_i],[_f,_f,new_angle,_f,_f,_f,_f],_s)
        #check how much power is being generated from the plant
        #as a first step, set output from new plant to 0 and check in Kieran's script how he organises the automated redispatch and possibly adopt that approach
        for gen_id in range (0, len(gens)):
            #psspy.machine_chng_2(19557,r"""1""",[_i,_i,_i,_i,_i,_i],[ Pout,Qout, Qmax,Qmin,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            psspy.machine_chng_2(gens[gen_id][0],str(gens[gen_id][1]),[_i,_i,_i,_i,_i,_i],[ 0,0, 0,0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])


def test_convergence(tree = 0, method='fnsl', taps='step'):
    if(tree ==1):
        psspy.tree(1,0)
        psspy.tree(2,1)
        psspy.tree(2,1) 
        psspy.tree(2,1)
        psspy.tree(2,1)
        psspy.tree(2,1)
        psspy.tree(2,1)
        
    #optimised params from PSS/E "robust solution" 
#    psspy.solution_parameters_4([200,200,20,99,10],[ 0.4, 0.4,  0.4,   10.0,    0.4,   10.0,  1.0,  0.1E-04,  5.0,   0.7,   0.0001,   0.005,  1.0,   0.05,    0.99,  0.99,  0.1,   0.1E-04,  100.0])
    #default params
#    psspy.solution_parameters_4([100,20,20,99,10],[ 1.60, 1.60, 1.000, 0.00010, 1.000, 0.100, 1.00, 0.10E-04, 5.000, 0.700, 0.000100, 0.0050, 1.000, 0.05000, 0.990, 0.990, 0.100, 0.10E-04, 100.0])
    #adjusted params - best results
    psspy.solution_parameters_4([100,200,20,99,10],[ 0.4, 0.4, 1.000, 0.00010, 1.000, 0.100, 1.00, 0.10E-04, 5.000, 0.700, 0.000100, 0.0050, 1.000, 0.05000, 0.990, 0.990, 0.100, 0.10E-04, 100.0])

    #psspy.fdns([1,0,0,1,1,0,99,0])
    #psspy.fdns([1,0,0,1,1,0,99,0]) 
    #psspy.fdns([1,0,0,1,1,0,99,0]) 
    if(method=='fnsl'):
        psspy.fnsl([1,0,0,1,1,0,0,0])
        psspy.fnsl([1,0,0,1,1,0,0,0])
        psspy.fnsl([1,0,0,1,1,0,0,0])
    elif(method=='fdns'):
        psspy.fdns([0,0,0,1,1,0,0,0])
        psspy.fdns([0,0,0,1,1,0,0,0])
        psspy.fdns([0,0,0,1,1,0,0,0])
    elif(method=='nsol'):
        psspy.nsol([0,0,0,1,1,0,0])
        psspy.nsol([0,0,0,1,1,0,0])
        psspy.nsol([0,0,0,1,1,0,0])
    elif(method=='solv'):
        psspy.solv([0,0,0,1,1,0])
        psspy.solv([0,0,0,1,1,0])
        psspy.solv([0,0,0,1,1,0])
    elif(method=='mslv'):
        psspy.mslv([0,0,0,1,1,0])
        psspy.mslv([0,0,0,1,1,0])
        psspy.mslv([0,0,0,1,1,0])
    
#    psspy.fdns([0,0,0,0,0,0,99,0])
#    psspy.fdns([0,0,0,0,0,0,99,0])
#    psspy.fdns([0,0,0,0,0,0,99,0])   
#    psspy.fdns([0,0,0,1,1,0,99,0])
#    psspy.fdns([0,0,0,1,1,0,99,0])
#    psspy.fdns([0,0,0,1,1,0,99,0])
    mismatch=psspy.sysmsm()
    print("the total mismatch is "+str(mismatch))
    if(abs(mismatch)>1):
        print("The system did not converge.")
    else:
        print("The system converged.")       
    return mismatch

#will add list of generators to the load flow case. If a generator already exists in a location, it will always add the new generator via a 0-imepdance line to ensure that there is no other machine at the same bus.
#the parameter "total_power" allows to add the generators as propotionally reduced in size. This can be used for  "proposed" projects, where you want to include a weitghted representaiton of proposed generators per are, 
#whilst limiting the total to a lower number than the sum of all propsoed generators, to account for the fact that not all projects will be built.
def add_generators(year, new_gens, public_announced_out, pc_out=0): #pc_out is initial output of newly added generators. Set to 0 to not disturb loadflow case
    total_power=0
    new_gen_stats=calc_new_gen_stats(new_gens)
    
    #change rating of proposed gens depending on scenario
        #Proposed gens: assumed to connect into nearest HV bus. Highest available voltage is assumed
    for gen_cnt in range(0, len(new_gens)):#gen in new_gens: #add generators to case, adapt rating of the generators with status "publicly announced"
        gen=new_gens[gen_cnt]
        print(gen)
        if(True):
#            if('Mulwala' in gen['name']):
#                print('stop here')
        #with silence():
            if('line' in gen.keys()):#add POC
                ierr, pu = psspy.busdat(gen['line']['bus1'],'PU') 
                ierr, ang = psspy.busdat(gen['line']['bus1'],'ANGLED')
                ierr, base = psspy.busdat(gen['line']['bus1'],'BASE')
                if(ierr!=0):
                    dummy_stop=1
                else:
                    ierr=psspy.ltap(gen['line']['bus1'],gen['line']['bus2'],gen['line']['id'], gen['line']['ratio'],gen['bus'],gen['name'][0:8]+"_"+gen['type'][0:3], base)
                    psspy.bus_chng_4(gen['bus'], 0, [2,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s)
            else:
                ierr, pu = psspy.busdat(gen['POC'],'PU') 
                ierr, ang = psspy.busdat(gen['POC'],'ANGLED')
                ierr, base = psspy.busdat(gen['POC'],'BASE')
                if(ierr!=0):
                    dummy_stop=1
                else:
                    ierr=psspy.bus_data_4(gen['bus'],0,[2,int(str(gen['bus'])[0]),1,1],[base, pu, ang, 1.1, 0.9, 1.1, 0.9],gen['name'][0:8]+"_"+gen['type'][0:3])

                psspy.branch_data_3(gen['bus'],gen['POC'],r"""1""",[1,gen['POC'],1,0,0,0],[0.0, 0.1,0.0,0.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0],[9999.0,9999.0,9999.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")#zero impedance branch
#            psspy.plant_data_3(gen['bus'],0,0,[ 1.0, 100.0])   
            if(ierr!=0): #if buses are missing, do not add generator
                gen['sim_status']='failed'
            elif( (not type(gen['status'])==int) or ( (type(gen['status'])==int) and (gen['status']<=year) ) ):
                if(not 'Pmax' in gen.keys()):
                    MWbase=0.0
                #if generator has a Pmax (is an actual generator or storage) and is of status publicly announced, change rating according to amounts to be included
                elif(gen['status']=='publicly announced'):
                    if (str(gen['bus'])[0]=='2' ):
                        pub_announced_scal=public_announced_out['NSW']/new_gen_stats['NSW_pub']
                    elif(str(gen['bus'])[0]=='3' ):
                        pub_announced_scal=public_announced_out['VIC']/new_gen_stats['VIC_pub']
                    MWbase=pub_announced_scal*gen['Pmax']
                    gen['Pmax']=gen['Pmax']*pub_announced_scal
                    if('Qmin') in gen.keys():
                        gen['Qmin']=gen['Qmin']*pub_announced_scal
                    if('Qmax') in gen.keys():
                        gen['Qmin']=gen['Qmin']*pub_announced_scal
                    if('Cap') in gen.keys():
                        gen['Cap']=gen['Cap']*pub_announced_scal                    
                else:
                    MWbase=base=gen['Pmax']
                if(pu<1.0):
                    v_target=1.0
                elif(pu<1.1):
                    v_target=pu
                else:
                    v_target=1.1
                psspy.plant_data_3(gen['bus'],0,0,[v_target,_f])
                if(gen['type']=='storage'):
                    psspy.machine_data_2(gen['bus'],"1",[1,1,0,0,0,0],[MWbase*pc_out/100,0.0, MWbase*0.4,-MWbase*0.4, MWbase,-MWbase, MWbase,0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
                elif((gen['type']=='solar')or(gen['type']=='wind')or(gen['type']=='gas')or(gen['type']=='waste')):
                    ierr=psspy.machine_data_2(gen['bus'],"1",[1,1,0,0,0,0],[MWbase*pc_out/100,0.0, MWbase*0.4,-MWbase*0.4, MWbase,0.0, MWbase,0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
                elif (gen['type']=='SVC'):
                    psspy.machine_data_2(gen['bus'],"1",[1,1,0,0,0,0],[0.0,0.0, gen['Qmax'],gen['Qmin'], 0.0,0.0, max(abs(gen['Qmax']), abs(gen['Qmin'])),0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
                elif (gen['type']=='syncon'):
                    psspy.machine_data_2(gen['bus'],"1",[1,1,0,0,0,0],[0.0,0.0, gen['Qmax'],-gen['Qmax']*0.6, 0.0,0.0, gen['Qmax'],0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
                
                total_power+=MWbase
                gen['sim_status']='success'
            mismatch=test_convergence(tree=1)
            if(mismatch>1):
                print('shit')
    #    if(total_power>0):
    #        total_P=0
    #        for gen in new_gens:
    #            total_P=total_P+gen[2]
    #        gen_scal=total_power/total_P
    #        for gen in new_gens:
#            gen[2]=gen[2]*gen_scal
    test_convergence(tree=1)
    return total_power

def calc_new_gen_stats(new_gens):
    NSW_MW=0
    for gen in new_gens: 
        if (str(gen['bus'])[0]=='2' ):
            if(gen['status']=='publicly announced'): 
                NSW_MW+=gen['Pmax']
    VIC_MW=0
    for gen in new_gens: 
        if (str(gen['bus'])[0]=='3' ):
            if(gen['status']=='publicly announced'): 
                VIC_MW+=gen['Pmax']
                
    return{'VIC_pub':VIC_MW, 'NSW_pub':NSW_MW}

#Takes a loadflow case, list of wind and solar generators, merit-order list of synchronous generators to be displaced, and reference nodes for the wind/solar projects
#All wind and Solar projects will be dispatched per the provided reference node (propotional to their rating). If one reference node is offline and a backup node is provided, that node will be used.
#Set upper P limit of renewable assets to the same as output. (which may be below rating) --> ensure that ACCC algorithm does not increase output of renewables. 
def redispatch2(reference_gen_out, gens, current_case, ): #variables have been updated
    #set synchronous generators for redispatch
    merit_order=[[30491,'11'],[30492,'12'],[30361,'1'],[30362,'2'],[30363,'3'],[30364,'4'],[30365,'1'],[30366,'2'],[30367,'3'],[30421,'1'],[30422,'2'],[30451,'1'],[30452,'2'],[30453,'3'],[30454,'4'],[30455,'5'],[30456,'6'],[30841,'1'],[30842,'2'],[30843,'3'],[30844,'4'],[30524,'1'],[30525,'2'],[30540,'1'],
                 [30441,'1'],[30442,'2'],[30443,'3'],[30444,'4'],[30445,'1'],[30446,'2'],[30941,'1'],[30942,'2'],[30943,'3'],[30944,'4']] #list of most expensive generators to switch off if generation too high, most expensive generator comes first.
    
    total_sync=0
    for sync_gen in merit_order:
        ierr, p_out = psspy.macdat(sync_gen[0],sync_gen[1],'P')
        ierr, p_max =  psspy.macdat(sync_gen[0],sync_gen[1],'PMAX')
        sync_gen.append(p_out)
        sync_gen.append(p_max)
        total_sync+=p_out
    
    #determine wind setpoint in pu --> previously set to McArthur, but changed to Waubra because of higher capacity factor
    ierr, p_out=psspy.macdat(30810, '1', 'P')
    if(ierr!=0):
         p_out=0
    waubra_pu=p_out/187.5 #divided by waubra rating
#    mcArthur_pu=psspy.brnflo(32790,35791,'1')
#    if(mcArthur_pu[1]==None): #if the branch is switched off (mcArthur disconnected)
#        mcArthur_pu=0
#    else:
#        mcArthur_pu=mcArthur_pu[1].real/(64.575+43.05+107.625+107.625+67.575+107.625) #nameplate rating is only 420 MW. although the machine rating indicates Pmax of almost 500 MW


    #wind_pu=mcArthur_pu
    wind_pu=waubra_pu
    solar_pu=hor_output/100000000.0
    
    solar_total=0
    wind_total=0
    renewable_gens=[]
    for gen in wind_gens:
        disp=gen[2]*wind_pu
        gen=gen.append(disp) #adds entry with dispatch
        wind_total+=disp

    for gen in solar_gens:
        disp=gen[2]*solar_pu
        gen=gen.append(disp) #adds entry with dispatch. will be added in 4th position
        solar_total+=disp

    hor_init=solar_gens[-1][4]
        
    renewables_total=solar_total+wind_total
    renewable_gens=wind_gens+solar_gens 
    renewables_total_reduced=renewables_total
    #reduce renewables output preoportionally if required
    if(renewables_total>total_sync):
        renewables_total_reduced=0
        for gen in renewable_gens:
            gen[4]=gen[4]*(total_sync/renewables_total) #set total amount of additional renewables to roughly displace all synchronouns generators
            renewables_total_reduced+=gen[4]
     
    prev_slack_bus_data=identify_slack_buses()  
    slack_bus_nr=prev_slack_bus_data.keys()[0]# assuming only one slack bus is in the system, which should be the case per previous checks
    #add renewables 1GW at a time
    gen_cnt=0
    while(gen_cnt<len(renewable_gens)): #make sure all renewables are taken into account
        temp_sum=0
        while (temp_sum<1000) and (gen_cnt<len(renewable_gens)) : #always displace ~1GW at a time --> redispatch occurs in multiple batches
            print(renewable_gens[gen_cnt])
            psspy.machine_data_2(renewable_gens[gen_cnt][0],renewable_gens[gen_cnt][1],[_i,_i,_i,_i,_i,_i],[(renewable_gens[gen_cnt][4]),0.0, renewable_gens[gen_cnt][2]*0.4,-renewable_gens[gen_cnt][2]*0.4, (renewable_gens[gen_cnt][4]),0.0, renewable_gens[gen_cnt][2],0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
            print ('unit '+str(renewable_gens[gen_cnt][1])+' at bus '+str(renewable_gens[gen_cnt][0])+' is redispatched to '+str(renewable_gens[gen_cnt][4])+' MW')
            test_convergence()
            temp_sum+=renewable_gens[gen_cnt][4] #add dispatch point of renewable_gen to temporary sum
            gen_cnt+=1
        slack_bus_data=identify_slack_buses()
        #lower output of synchronous generators to accomodate new renewable generation
        i=0
        while(slack_bus_data[slack_bus_nr]['P']<prev_slack_bus_data[slack_bus_nr]['P']-1):
            i+=1
            if(i==len(merit_order)):
                break;
            p_max=merit_order[i][3]
            p_out=merit_order[i][2]
            if( (p_out is not None) and (p_max is not None) ):
                if(p_out>0):
                    #if more than the cyns machine output is displaced, swithc it off
                    if((prev_slack_bus_data[slack_bus_nr]['P']-slack_bus_data[slack_bus_nr]['P'])>p_out): 
                        #switch machine off
                        psspy.machine_chng_2(merit_order[i][0],merit_order[i][1],[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        merit_order[i][2]=0
                        test_convergence()
                        slack_bus_data=identify_slack_buses()
                    #if less than machine output is displaced, reduce setpoint (as long as it remains above 0.4*p_max)
                    elif(p_out>= 0.4*p_max): #do not lower machine dispatch below 40% max output. If that is required, then switch off instead
                        new_p=max((p_max*0.4), p_out+(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P']) ) #lower by the amount that slack bus produces less than prior to the renewables being added in
                        psspy.machine_data_2(merit_order[i][0],merit_order[i][1],[_i,_i,_i,_i,_i,_i],[new_p,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        merit_order[i][2]=new_p
                        print (str(merit_order[i][0])+' is redispatched')
                        test_convergence()
                        slack_bus_data=identify_slack_buses()
    #after all renewables have been dispatched, slightly increase output of any remaining synchronous generators in case they were reduced by too much (may not be necessary to implement this)
    #while(slack_bus_data[slack_bus_nr]['P']>prev_slack_bus_data[slack_bus_nr]['P']+1):
    
    #disconnect line from 66 kV HOR to 66 kV Buangor --> only do that here, to avoid messing up the load flow during integration of the additional plants.
    
    psspy.branch_chng_3(36030,36050,r"""1""",[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
    test_convergence(tree=1)
    test_convergence()

    return wind_pu, solar_pu, renewable_gens[-1][4], hor_init

#takes a list of newly added generators (with dispatch target), and a list of redispatchable generators (with lower limits indicated)
#then tries to reduce output of redispatchable (coal and gas ) an a way that allows for the renewables to be dispatched at their dispatch target. 
#Function attempts to achieve balance per state (to keep interconnector flows steady). 
#Batteries are being dispatched as absorbing if both wind and solar dispatch targets in the (respective state) are above average
#Batteries are being dispatched as providing if both wind and solar dispatch targets in the respective state are below average.
    #Maybe check that no lines leading away from the new asset are being overloaded as a result of the redispatch. (otpional)
    #Maybe exclude proposed projects > x GW.
def redispatch3(reference_gen_out, reference_gens, new_gens, flow_shift): #shift balance in quntity in MW of flow shift towards VIC for tomes of high solar availability in NSW
    bat_flags={'VIC':0, 'NSW':0}
    #declare available coal and gas generation in VIC that can be switched off. Determine total combined capacity that can be freed
    merit_order_VIC=[{'bus':352011,'id':'11','name':'Mortlake1','type':'gas'},      {'bus':352012,'id':'12','name':'Mortlake2','type':'gas'},
                     {'bus':337001,'id':'1','name':'Jeeralang1','type':'gas'},      {'bus':337002,'id':'2','name':'Jeeralang2', 'type':'gas'},      {'bus':337003, 'id':'3','name':'Jeeralang3', 'type':'gas'},     {'bus':337004, 'id':'4','name':'Jeeralang4','type':'gas'}, {'bus':338001, 'id':'1', 'name':'Jeeralang5', 'type':'gas'},{'bus':338002,'id':'2','name':'Jeeralang6', 'type':'gas'},{'bus':338003,'id':'3','name':'Jeeralang7', 'type':'gas'},
                     {'bus':344001,'id':'1','name':'Laverton1', 'type':'gas'},      {'bus':344002,'id':'2','name':'Laverton2', 'type':'gas'},
                     {'bus':382001,'id':'1', 'name':'ValleyPower1', 'type':'gas'},  {'bus':382002,'id':'2', 'name':'ValleyPower2', 'type':'gas'},   {'bus':382003, 'id':'3', 'name':'ValleyPower3', 'type':'gas'},  {'bus':382004, 'id':'4', 'name':'ValleyPower4', 'type':'gas'},{'bus':382005, 'id':'5', 'name':'ValleyPower5', 'type':'gas'},{'bus':382006, 'id':'6', 'name':'ValleyPower6', 'type':'gas'},
                     {'bus':370001,'id':'1', 'name':'Somerton1', 'type':'gas'},     {'bus':370002,'id':'2', 'name':'Somerton2', 'type':'gas'},      {'bus':370003, 'id':'3', 'name':'Somerton3', 'type':'gas'},     {'bus':370004, 'id':'4', 'name':'Somerton4', 'type':'gas'},
                     {'bus':308001,'id':'1','name':'Bairnsdale1','type': 'gas'},    {'bus':308002,'id': '2','name':'Bairnsdale2', 'type':'gas'},
                     {'bus':360001,'id':'1','name':'Newport', 'type':'gas'},
                     {'bus':346001,'id':'1', 'name':'LoyYang1', 'type':'coal'},     {'bus':346002, 'id':'2', 'name':'LoyYang2', 'type':'coal'},     {'bus':346003, 'id':'3', 'name':'LoyYang3', 'type':'coal'},     {'bus':346004, 'id':'4', 'name':'LoyYang4', 'type':'coal'}, {'bus':347001, 'id':'1', 'name':'LoyYang5', 'type':'coal'},{'bus':347002,'id':'2', 'name':'LoyYang6', 'type':'coal'},
                     {'bus':390001,'id':'1', 'name':'Yallourn1', 'type':'coal'},    {'bus':390002, 'id':'2', 'name':'Yallour21', 'type':'coal'},    {'bus':390003, 'id':'3', 'name':'Yallourn3', 'type':'coal'},    {'bus':390004, 'id':'4', 'name':'Yallourn4', 'type':'coal'}] #list of most expensive generators to switch off if generation too high, most expensive generator comes first.
    #determine and declare availabel coal and gas generation in NSW that can be switched off. Determine total combined capacity that can be freed
    merit_order_NSW=[                    #[218801, '1', 'Blowering', 'hydro'],
                    #[221205, '3', 'Burrinjuck1', 'hydro'],[221205, '4', 'Burrinjuck2', 'hydro'],[221205, '5', 'Burrinjuck3', 'hydro'],
                    #[239201, '1' 'Guthega1', 'hydro'], [239202, '2' 'Guthega2', 'hydro'],
                    #[240801, '1' 'Hume1, 'hydro'], [240802, '2', 'Hume2', 'hydro'],
                    #[242801, '1', 'Jounama', 'hydro'],
                    {'bus':227601, 'id':'1', 'name':'Colongra1', 'type':'gas'},     {'bus':227602, 'id':'2', 'name':'Colongra1', 'type':'gas'}, {'bus':227603, 'id':'3', 'name':'Colongra1', 'type':'gas'}, {'bus':227604, 'id':'4', 'name':'Colongra1', 'type':'gas'},
                    {'bus':219601, 'id':'1', 'name':'BrokenHill1', 'type':'gas'},   {'bus':219602, 'id':'2', 'name':'BrokenHill2', 'type':'gas'},
                    {'bus':234801, 'id':'1', 'name':'Gadara', 'type':'gas'}, #assuming gas. No info online
                    {'bus':241201, 'id':'1', 'name':'HunterValley1', 'type':'gas'}, {'bus':241202, 'id':'2', 'name':'HunterValley2', 'type':'gas'},
                    {'bus':270801, 'id':'1', 'name':'Smithsfield1', 'type':'gas'},  {'bus':270802, 'id':'2', 'name':'Smithsfield2', 'type':'gas'}, {'bus':270803, 'id':'3', 'name':'Smithsfield3', 'type':'gas'}, {'bus':270804, 'id':'4', 'name':'Smithsfield4', 'type':'gas'},
                    {'bus':274801, 'id':'1', 'name':'Tallawarra', 'type':'gas'}, #Combined Cycle
                    {'bus':279611, 'id':'1', 'name':'Uranquinty1', 'type':'gas'},   {'bus':279612, 'id':'1', 'name':'Uranquinty2', 'type':'gas'}, {'bus':279613, 'id':'1', 'name':'Uranquinty3', 'type':'gas'}, {'bus':279614, 'id':'1', 'name':'Uranquinty4', 'type':'gas'}, 
                    {'bus':233201, 'id':'1', 'name':'Eraring1', 'type':'coal'},     {'bus':233202, 'id':'2', 'name':'Eraring2', 'type':'coal'}, {'bus':233203, 'id':'3', 'name':'Eraring3', 'type':'coal'}, {'bus':233211, 'id':'11', 'name':'Eraring1', 'type':'coal'},
                    {'bus':215201, 'id':'1', 'name':'Baywswater1', 'type':'coal'},  {'bus':215202, 'id':'2', 'name':'Baywswater2', 'type':'coal'}, {'bus':215203, 'id':'3', 'name':'Baywswater3', 'type':'coal'},{'bus':215204, 'id':'4', 'name':'Baywswater4', 'type':'coal'},
                    {'bus':249201, 'id':'1', 'name':'Liddell1', 'type':'coal'},     {'bus':249202, 'id':'2', 'name':'Liddell2', 'type':'coal'}, {'bus':249203, 'id':'3', 'name':'Liddell3', 'type':'coal'}, {'bus':249204, 'id':'4', 'name':'Liddell4', 'type':'coal'},
                    {'bus':257601, 'id':'1', 'name':'MountPiper1', 'type':'coal'},  {'bus':257602, 'id':'2', 'name':'MountPiper2', 'type':'coal'},
                    {'bus':280005, 'id':'5', 'name':'ValesPoint', 'type':'coal'},   {'bus':280006, 'id':'6', 'name':'ValesPoint', 'type':'coal'},
                    {'bus':216801, 'id':'1', 'name':'Bendeela1', 'type':'stor','Pmin':-40}, {'bus':216801, 'id':'2', 'name':'Bendeela2', 'type':'stor','Pmin':-40}, 
                    {'bus':243203, 'id':'3', 'name':'KangarooValley', 'type':'stor', 'Pmin':-80}, {'bus':243204, 'id':'4', 'name':'KangarooValley', 'type':'stor', 'Pmin':-80},
                    {'bus':251201, 'id':'1', 'name':'LowerTumut1', 'type':'stor', 'Pmin':-10}, {'bus':251202, 'id':'2', 'name':'LowerTumut2', 'type':'stor', 'Pmin':-10},{'bus':251203, 'id':'3', 'name':'LowerTumut3', 'type':'stor', 'Pmin':-10}, {'bus':251204, 'id':'4', 'name':'LowerTumut4', 'type':'stor', 'Pmin':-195}, {'bus':251205, 'id':'5', 'name':'LowerTumut5', 'type':'stor', 'Pmin':-195}, {'bus':251206, 'id':'6', 'name':'LowerTumut6', 'type':'stor', 'Pmin':-195},
                    {'bus':279201, 'id':'1', 'name':'UpperTumut1', 'type':'stor', 'Pmin':-10}, {'bus':279202, 'id':'2', 'name':'UpperTumut2', 'type':'stor', 'Pmin':-10}, {'bus':279203, 'id':'3', 'name':'UpperTumut3', 'type':'stor', 'Pmin':-10}, {'bus':279204, 'id':'4', 'name':'UpperTumut4', 'type':'stor', 'Pmin':-10},#[279205, '5', 'UpperTumut5', 'stor', -0], [279206, '6', 'UpperTumut6', 'stor', -0], [279207, '7', 'UpperTumut7', 'stor', -0], [279208, '8', 'UpperTumut8', 'stor', -0],

                    ]
    #total displacable generation
    total_sync_VIC=0
    for sync_gen in merit_order_VIC:
        ierr, p_out = psspy.macdat(sync_gen['bus'],sync_gen['id'],'P')
        ierr, p_max =  psspy.macdat(sync_gen['bus'],sync_gen['id'],'PMAX')
        ierr, p_min =  psspy.macdat(sync_gen['bus'],sync_gen['id'],'PMIN')
        sync_gen['Pout']=p_out
        sync_gen['Pmax']=p_max
        sync_gen['Pmin']=p_min
        total_sync_VIC+=p_out
    
    #total displacable generation
    total_sync_NSW=0
    for sync_gen in merit_order_NSW:
        ierr, p_out = psspy.macdat(sync_gen['bus'],sync_gen['id'],'P')
        ierr, p_max =  psspy.macdat(sync_gen['bus'],sync_gen['id'],'PMAX')
        ierr, p_min =  psspy.macdat(sync_gen['bus'],sync_gen['id'],'PMIN')
        sync_gen['Pout']=p_out
        sync_gen['Pmax']=p_max
        sync_gen['Pmin']=p_min
        total_sync_NSW+=p_out
        

    
    #Calculate amount of additional renewables in NSW for given dispatch target
    additional_gen_NSW=0
    for gen in new_gens:
        if ((str(gen['bus'])[0]=='2') and ('Pmax' in gen.keys()) and (gen['sim_status']=='success')): #if generator in NSW and it is a generator and not SVC or SynCon
            if(gen['type']=='solar'):
                additional_gen_NSW+=gen['Pmax']*reference_gen_out['NSW']['solar'][3]  
            elif(gen['type']=='wind'):
                additional_gen_NSW+=gen['Pmax']*reference_gen_out['NSW']['wind'][3]  
            #assume that batteries are discharging when both sind and solar are belwo 30% capacity, and are charging when both wind and solar are above 70% capacity            
            elif(gen['type']=='storage'):
                if( (reference_gen_out['NSW']['solar'][3] >0.7) and (reference_gen_out['NSW']['wind'][3] >0.7) ): #if lots of wind and solar is available, prices will be low and storage is assumed to be recharging
                    additional_gen_NSW-=gen['Pmax']
                elif( (reference_gen_out['NSW']['solar'][3] <0.2) and (reference_gen_out['NSW']['wind'][3] <0.2) ): #if few wind and solar is available prices will be higher and storage is assumed to be discharging
                    additional_gen_NSW+=gen['Pmax']
    #adjust dispatch target if additional generation exceeds displacable generation
    if(additional_gen_NSW>total_sync_NSW):
        NSW_scaling=total_sync_NSW/additional_gen_NSW
    else:
        NSW_scaling=1.0
    reference_gen_out['NSW']['wind'].append(NSW_scaling*reference_gen_out['NSW']['wind'][3])
    reference_gen_out['NSW']['solar'].append(NSW_scaling*reference_gen_out['NSW']['solar'][3])
    #Calculate amount of additioanl renewables in VIC for given dispatch target
    additional_gen_VIC=0
    for gen in new_gens:
        if ((str(gen['bus'])[0]=='3') and ('Pmax' in gen.keys()) and(gen['sim_status']=='success')): #if generator in NSW and it is a generator and not SVC or SynCon
            if(gen['type']=='solar'):
                additional_gen_VIC+=max(gen['Pmax']*reference_gen_out['VIC']['solar'][3], 0.0) 
            elif(gen['type']=='wind'):
                additional_gen_VIC+=max(gen['Pmax']*reference_gen_out['VIC']['wind'][3], 0.0)
            #assume that batteries are discharging when both wind and solar are belwo 20% capacity, and are charging when both wind and solar are above 70% capacity            
            elif(gen['type']=='storage'):
                if( (reference_gen_out['VIC']['solar'][3] >0.7) and (reference_gen_out['VIC']['wind'][3] >0.7) ): #if lots of wind and solar is available, prices will be low and storage is assumed to be recharging
                    additional_gen_VIC-=gen['Pmax']
                elif( (reference_gen_out['VIC']['solar'][3] <0.2) and (reference_gen_out['VIC']['wind'][3] <0.2) ): #if few wind and solar is available prices will be higher and storage is assumed to be discharging
                    additional_gen_VIC+=gen['Pmax']
    #adjust dispatch target if additional generation exceeds displacable generation
    if(additional_gen_VIC>total_sync_VIC):
        VIC_scaling=total_sync_VIC/additional_gen_VIC
    else:
        VIC_scaling=1.0
    reference_gen_out['VIC']['wind'].append(VIC_scaling*reference_gen_out['VIC']['wind'][3])
    reference_gen_out['VIC']['solar'].append(VIC_scaling*reference_gen_out['VIC']['solar'][3])
    #    print('SYNC DISPATCH BEFORE REDISAPTCH\n')
    print("total sync VIC: "+str(total_sync_VIC))
    print("total renewables VIC: "+str(additional_gen_VIC)+'\n')
#    for gen in merit_order_VIC: print(gen['name']+' bus: '+str(gen['bus'])+' Pout: '+str(gen['Pout']))
#    print('')
    print("total sync NSW: "+str(total_sync_NSW))
    print("total renewables NSW: "+str(additional_gen_NSW)+'\n')
#    for gen in merit_order_NSW: print(gen['name']+' bus: '+str(gen['bus'])+' Pout: '+str(gen['Pout']))
    
    #Dispatch renewables per dispatch target, whilst conitnuously checking convergence.
    prev_slack_bus_data=identify_slack_buses()  
    slack_bus_nr=prev_slack_bus_data.keys()[0]# assuming only one slack bus is in the system, which should be the case per previous checks
    prev_slack_bus_data[slack_bus_nr]['P']=prev_slack_bus_data[slack_bus_nr]['P']+flow_shift*reference_gen_out['NSW']['solar'][3] #Offsetting "prev_slack_bus_nr" to shift energy flow towards NSW
    #add VIC renewables 1GW at a time
    gen_cnt=0
    while(gen_cnt<len(new_gens)): #make sure all renewables are taken into account
        temp_sum_vic=0
        while (temp_sum_vic<1000) and (gen_cnt<len(new_gens)) : #always displace ~1GW at a time --> redispatch occurs in multiple batches
            gen=new_gens[gen_cnt]
            if('Stockyard' in gen['name']):
                print('debug_breakpoint')
            if(str(gen['bus'])[0]=='3' and (gen['sim_status']=='success')): #VIC and generator successfully added
                #print(new_gens[gen_cnt])
                if(gen['type']=='solar'):
                    Ptemp=max(gen['Pmax']*reference_gen_out['VIC']['solar'][3]*VIC_scaling, 0.0)
                    with silence():
                        ierr=psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, Ptemp,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    gen['Pout']=Ptemp
                    if(ierr!=0):
                        dummy_stop=1
                    temp_sum_vic+=Ptemp  
                elif(gen['type']=='wind'):
                    Ptemp=max(gen['Pmax']*reference_gen_out['VIC']['wind'][3]*VIC_scaling, 0.0) 
                    with silence():
                        ierr=psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, Ptemp,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    gen['Pout']=Ptemp
                    if(ierr!=0):
                        dummy_stop=1
                    temp_sum_vic+=Ptemp
                #assume that batteries are discharging when both sind and solar are belwo 30% capacity, and are charging when both wind and solar are above 70% capacity            
                elif(gen['type']=='storage'):
                    if( (reference_gen_out['VIC']['solar'][3] >0.7) and (reference_gen_out['VIC']['wind'][3] >0.7) ): #if lots of wind and solar is available, prices will be low and storage is assumed to be recharging
                        bat_flags['VIC']=1
                        Ptemp=-gen['Pmax']*VIC_scaling
                        with silence():
                            psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) # do not restrict grid on the 
                        gen['Pout']=Ptemp
                        temp_sum_vic+=Ptemp
                       
                    elif( (reference_gen_out['VIC']['solar'][3] <0.2) and (reference_gen_out['VIC']['wind'][3] <0.2) ): #if few wind and solar is available prices will be higher and storage is assumed to be discharging
                        bat_flags['VIC']=-1
                        Ptemp=gen['Pmax']*VIC_scaling
                        with silence():
                            psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) # do not restrict grid on the 
                        gen['Pout']=Ptemp
                        temp_sum_vic+=Ptemp
            gen_cnt+=1
            mismatch=test_convergence()
            if(mismatch>=0.5):
                print('weird')
                
            

        with silence():
            test_convergence()
        #save snapshot for debugging
        #set_slack_bus_nsw()
        slack_bus_data=identify_slack_buses()
        #lower output of synchronous generators to accomodate new renewable generation
        i=0
        
        while(slack_bus_data[slack_bus_nr]['P']<prev_slack_bus_data[slack_bus_nr]['P']-1):
            i+=1
            if(i==len(merit_order_VIC)):
                break;
            p_min=merit_order_VIC[i]['Pmin']
            p_max=merit_order_VIC[i]['Pmax']
            p_out=merit_order_VIC[i]['Pout']
            if( (p_out is not None) and (p_max is not None) ):
                if(p_out>0):
                    if(merit_order_VIC[i]['type']!='stor'):
                        #if more than the cyns machine output is displaced, swithc it off
                        if((prev_slack_bus_data[slack_bus_nr]['P']-slack_bus_data[slack_bus_nr]['P'])>p_out): 
                            #switch machine off (do not switch off but set output power to 0 MW. THis way it is easier to reactivate if required to resolve line overloadings)
                            with silence():
                                psspy.machine_chng_2(merit_order_VIC[i]['bus'],merit_order_VIC[i]['id'],[_i,_i,_i,_i,_i,_i],[0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_VIC[i]['bus'])+' type ' +merit_order_VIC[i]['type']+' is redispatched from ' +str(p_out)+' to 0')
                            merit_order_VIC[i]['Pout']=0
                            #merit_order_VIC[i][2]=0
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()
                        #if less than machine output is displaced, reduce setpoint (as long as it remains above 0.4*p_max)
                        elif(p_out>= 0.4*p_max): #do not lower machine dispatch below 40% max output. If that is required, then switch off instead
                            new_p=max((p_max*0.4), p_out+(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P']) ) #lower by the amount that slack bus produces less than prior to the renewables being added in
                            with silence():
                                psspy.machine_data_2(merit_order_VIC[i]['bus'],merit_order_VIC[i]['id'],[_i,_i,_i,_i,_i,_i],[new_p,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_VIC[i]['bus'])+' type ' +merit_order_VIC[i]['type']+' is redispatched from ' +str(p_out)+' to '+str(round( new_p,2)))
                            merit_order_VIC[i]['Pout']=new_p
                            #merit_order_VIC[i][2]=new_p
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()
                    else:#storage generator
                        if((prev_slack_bus_data[slack_bus_nr]['P']-slack_bus_data[slack_bus_nr]['P'])>(p_out-p_min)):
                            with silence():
                                psspy.machine_data_2(merit_order_VIC[i]['bus'],merit_order_VIC[i]['id'],[_i,_i,_i,_i,_i,_i],[p_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_VIC[i]['bus'])+' type ' +merit_order_VIC[i]['type']+' is redispatched from ' +str(p_out)+' to '+str(round( new_p,2)))
                            merit_order_VIC[i]['Pout']=p_min
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()
                        else:
                            new_p=p_out+(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P']) #lower by the amount that slack bus produces less than prior to the renewables being added in
                            with silence():
                                psspy.machine_data_2(merit_order_VIC[i]['bus'],merit_order_VIC[i]['id'],[_i,_i,_i,_i,_i,_i],[new_p,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_VIC[i]['bus'])+' type ' +merit_order_VIC[i]['type']+' is redispatched from ' +str(p_out)+' to '+str(round(new_p,2)))
                            merit_order_VIC[i]['Pout']=new_p
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()
    #ADD NSW RENEWABLES 1GW at a time
    #Change Slack bus to NSW
    if(abs(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P'])<=2):
        print("Successful redispacth of VIC")
    else:
        print("dispatch gap is "+str(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P'])+" after addition of VIC generators")
    set_slack_bus_nsw()
    #shift Energy balance to reduce flow to NSW slightly    
    psspy.machine_chng_2(348099,r"""1""",[_i,_i,_i,_i,_i,_i],[slack_bus_data[slack_bus_nr]['P']-flow_shift*reference_gen_out['NSW']['solar'][3],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    test_convergence()
    
    prev_slack_bus_data=identify_slack_buses()  
    slack_bus_nr=prev_slack_bus_data.keys()[0]# assuming only one slack bus is in the system, which should be the case per previous checks
    gen_cnt=0
    while(gen_cnt<len(new_gens)): #make sure all renewables are taken into account
        temp_sum_nsw=0
        while (temp_sum_nsw<1000) and (gen_cnt<len(new_gens)) : #always displace ~1GW at a time --> redispatch occurs in multiple batches
            gen=new_gens[gen_cnt]
            if('Bomen' in gen['name']):
                print('break')
                identify_overloaded_lines()
            if(str(gen['bus'])[0]=='2' and (gen['sim_status']=='success')): #NSW and gen successfully added
#                print(new_gens[gen_cnt])
                if(gen['type']=='solar'):
                    Ptemp=max(gen['Pmax']*reference_gen_out['NSW']['solar'][3]*NSW_scaling,0.0)
                    with silence():
                        psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, Ptemp,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    gen['Pout']=Ptemp
                    temp_sum_nsw+=Ptemp  
                elif(gen['type']=='wind'):
                    Ptemp=max(gen['Pmax']*reference_gen_out['NSW']['wind'][3]*NSW_scaling,0.0) 
                    with silence():
                        psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, Ptemp,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    gen['Pout']=Ptemp
                    temp_sum_nsw+=Ptemp
                #assume that batteries are discharging when both sind and solar are belwo 30% capacity, and are charging when both wind and solar are above 70% capacity            
                elif(gen['type']=='storage'):
                    if( (reference_gen_out['NSW']['solar'][3] >0.7) and (reference_gen_out['NSW']['wind'][3] >0.7) ): #if lots of wind and solar is available, prices will be low and storage is assumed to be recharging
                        bat_flags['NSW']=1
                        Ptemp=-gen['Pmax']*NSW_scaling
                        with silence():
                            psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) # do not restrict grid on the 
                        gen['Pout']=Ptemp
                        temp_sum_nsw+=Ptemp
                       
                    elif( (reference_gen_out['NSW']['solar'][3] <0.2) and (reference_gen_out['NSW']['wind'][3] <0.2) ): #if few wind and solar is available prices will be higher and storage is assumed to be discharging
                        bat_flags['NSW']=-1
                        Ptemp=gen['Pmax']*NSW_scaling
                        with silence():
                            psspy.machine_chng_2(new_gens[gen_cnt]['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ Ptemp,_f,_f,_f, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) # do not restrict grid on the 
                        gen['Pout']=Ptemp
                        temp_sum_nsw+=Ptemp
            gen_cnt+=1
            mismatch=test_convergence()
            if(mismatch>=0.5):
                print('weird')
                
        with silence():
            test_convergence()
        slack_bus_data=identify_slack_buses()
        #lower output of synchronous generators to accomodate new renewable generation
        i=0
        while(slack_bus_data[slack_bus_nr]['P']<prev_slack_bus_data[slack_bus_nr]['P']-1): #Test 
            i+=1
            if(i==len(merit_order_NSW)):
                break;
            p_min=merit_order_NSW[i]['Pmin']
            p_max=merit_order_NSW[i]['Pmax']
            p_out=merit_order_NSW[i]['Pout']
            if( (p_out is not None) and (p_max is not None) ):
                if(p_out>p_min):
                    if(merit_order_NSW[i]['type']!='stor'):
                        #if more than the cyns machine output is displaced, swithc it off
                        if((prev_slack_bus_data[slack_bus_nr]['P']-slack_bus_data[slack_bus_nr]['P'])>p_out): 
                            #switch machine off (set to 0 MW)
                            with silence():
                                psspy.machine_chng_2(merit_order_NSW[i]['bus'],merit_order_NSW[i]['id'],[_i,_i,_i,_i,_i,_i],[0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_NSW[i]['bus'])+' type ' +merit_order_NSW[i]['type']+' is redispatched from ' +str(p_out)+' to 0')
                            merit_order_NSW[i]['Pout']=0
                            #merit_order_NSW[i][2]=0
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()
                        #if less than machine output is displaced, reduce setpoint (as long as it remains above 0.4*p_max)
                        elif(p_out>= 0.4*p_max): #do not lower machine dispatch below 40% max output. If that is required, then switch off instead
                            new_p=max((p_max*0.4), p_out+(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P']) ) #lower by the amount that slack bus produces less than prior to the renewables being added in
                            with silence():
                                psspy.machine_data_2(merit_order_NSW[i]['bus'],merit_order_NSW[i]['id'],[_i,_i,_i,_i,_i,_i],[new_p,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_NSW[i]['bus'])+' type ' +merit_order_NSW[i]['type']+' is redispatched from ' +str(p_out)+' to '+str(round( new_p,2)))
                            merit_order_NSW[i]['Pout']=new_p
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()
                    else:#storage generator
                        if((prev_slack_bus_data[slack_bus_nr]['P']-slack_bus_data[slack_bus_nr]['P'])>(p_out-p_min)):
                            with silence():
                                psspy.machine_data_2(merit_order_NSW[i]['bus'],merit_order_NSW[i]['id'],[_i,_i,_i,_i,_i,_i],[p_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_NSW[i]['bus'])+' type ' +merit_order_NSW[i]['type']+' is redispatched from ' +str(p_out)+' to '+str(round( new_p,2)))
                            merit_order_NSW[i]['Pout']=p_min
                            #merit_order_NSW[i][2]=0
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()
                        else:
                            new_p=p_out+(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P']) #lower by the amount that slack bus produces less than prior to the renewables being added in
                            with silence():
                                psspy.machine_data_2(merit_order_NSW[i]['bus'],merit_order_NSW[i]['id'],[_i,_i,_i,_i,_i,_i],[new_p,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            print (str(merit_order_NSW[i]['bus'])+' type ' +merit_order_NSW[i]['type']+' is redispatched from ' +str(p_out)+' to '+str(round(new_p,2)))
                            merit_order_NSW[i]['Pout']=new_p
                            with silence():
                                test_convergence()
                            slack_bus_data=identify_slack_buses()

        with silence():
            test_convergence()
        slack_bus_data=identify_slack_buses()
    if(abs(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P'])<=2):
        print("Successful redispacth of NSW")
    else:
        print("dispatch gap is "+str(slack_bus_data[slack_bus_nr]['P']-prev_slack_bus_data[slack_bus_nr]['P'])+" after addition of NSW generators")
    set_slack_bus_vic() #set slack bus back to Victoria
    
#    print('SYNC DISPATCH AFTER REDISAPTCH')
#    for gen in merit_order_VIC: print(gen['name']+' bus: '+str(gen['bus'])+' Pout: '+str(gen['Pout']))
#    print('')
#    for gen in merit_order_NSW: print(gen['name']+' bus: '+str(gen['bus'])+' Pout: '+str(gen['Pout']))
    #after all renewables have been dispatched, slightly increase output of any remaining synchronous generators in case they were reduced by too much (may not be necessary to implement this)
    #while(slack_bus_data[slack_bus_nr]['P']>prev_slack_bus_data[slack_bus_nr]['P']+1):
    
    #disconnect line from 66 kV HOR to 66 kV Buangor --> only do that here, to avoid messing up the load flow during integration of the additional plants.
    
    #psspy.branch_chng_3(36030,36050,r"""1""",[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
    test_convergence(tree=1)
    test_convergence()
    #return wind_pu, solar_pu, renewable_gens[-1][4], hor_init
    return bat_flags
    
    

#return list of buses with voltage outside 0.9 to 1.1 p.u.
def check_voltages(lower=0.9, upper=1.1):
    abnormal_buses=[]
    ierr, buses = psspy.abusint(-1, 1, 'NUMBER')
    bus_info=get_bus_info(buses[0], ['PU', 'NAME', 'BASE'])
    for bus in bus_info.keys():
        if((bus_info[bus]['PU']>upper) or (bus_info[bus]['PU']<lower) ):
            abnormal_buses.append([bus, bus_info[bus]['PU'], bus_info[bus]['NAME'], bus_info[bus]['BASE']])    
    abnormal_buses=sorted(abnormal_buses, key=lambda x: x[1] )
    return abnormal_buses

def print_volts(lower=0.9, upper=1.1):
    voltages=check_voltages(lower, upper)
    for bus_entry in voltages: print(bus_entry)
    return 0

#return list of lines operating beyond specification
def identify_overloaded_lines():
    overloaded_lines=[]
    #psspy.bsys(sid=0, numbus=2, buses=[frombus, tobus])
    ierr, branch_int_out=psspy.abrnint(-1,1,1,3,1,['FROMNUMBER', 'TONUMBER', 'STATUS', 'METERNUMBER'])
    ierr, branch_char_out=psspy.abrnchar(-1,1,1,3,1,['ID', 'FROMNAME', 'TONAME'])
    ierr, branch_data=psspy.abrnreal(-1,1,1,3,1, ['PCTMVARATE','PCTMVARATE1','PCTMVARATE2','PCTMVARATE3','RATE1','P','Q','MVA'])   
    for i in range(0, len(branch_int_out[0])):
#        if(branch_int_out[0][i]==230890 ):
#            print('debug_stop')
        if(branch_data[0][i]>100.0):
            overloaded_lines.append({'FROMBUS':branch_int_out[0][i], 'TOBUS':branch_int_out[1][i], 'ID':branch_char_out[0][i], 'P':branch_data[5][i], 'Q':branch_data[6][i], 'MVA':branch_data[7][i], 'PCTMVARATE':branch_data[0][i], 'RATING':branch_data[4][i]})
    return overloaded_lines   

#routine will go through a list of snapshots, check the output of generators included in gen_list as a function of it's rating 
# and then calculate the capacity factor for the entire period. gen_list must have formate[{'name':'Horsham', 'bus':32280, 'macID':1, 'rating':140},{...}]
def check_capacity_factor(gen_list, case_dir):
    case_list=[]
    for dirpath, dirnames, filenames in os.walk(case_dir):
        for filename in filenames:
            filepath=os.path.join(dirpath,filename)
            if filepath.endswith('.raw'):
                case_list.append([filename,filepath, gen_list])
    for gen in gen_list:
        gen['Pmax']=np.array(())
        gen['P']=np.array(())
        gen['Q']=np.array(())
#        gen['Pmax_vec']=np.array(())
#        gen['P_vec']=np.array(())
#        gen['Q_vec']=np.array(())
        P_tot=0
        Q_tot=0
        Pmax_tot=0
        n_cases=len(case_list)
        
        p=Pool(30)
        p.map(check_capacity_factor_subpr, case_list)
        
        for i in range(0, len(case_list)):
            print(i)
            with silence():
                psspy.read(0,case_list[i][1])
            p_out=0
            q_out=0
            p_max=0
            for machine in gen['machines']:
                with silence():
                    ierr, p_out_part = psspy.macdat(machine[0],str(machine[1]),'P')
                    ierr, q_out_part = psspy.macdat(machine[0],str(machine[1]),'Q')
                    ierr, p_max_part = psspy.macdat(machine[0],str(machine[1]),'PMAX')
                p_out=p_out+p_out_part
                q_out=q_out+q_out_part
                p_max=p_max+p_max_part
            gen['P']=np.append(gen['P'], p_out)
            gen['Q']=np.append(gen['Q'], q_out)
            gen['Pmax']=np.append(gen['Pmax'], p_max)
            P_tot+=p_out
            Q_tot+=q_out
            Pmax_tot+=p_max
        P_avg=P_tot/n_cases
        Q_avg=Q_tot/n_cases
        Pmax_avg=Pmax_tot/n_cases
        gen['Pmax_avg']=Pmax_avg
        gen['P_avg']=P_avg
        gen['Q_avg']=Q_avg
        gen['cap_fact']=P_avg/Pmax_avg
        
def check_capacity_factors_par(gen_list, case_dir):
    case_list=[]
    for dirpath, dirnames, filenames in os.walk(case_dir):
        for filename in filenames:
            filepath=os.path.join(dirpath,filename)
            if filepath.endswith('.raw'):
                case_list.append([filename,filepath, gen_list])
    p=Pool(30)
    print('starting subprocesses')
    results=p.map(check_capacity_factor_subpr, case_list)
#    results=[]
#    for case in case_list:
#        results.append(check_capacity_factor_subpr(case))
    for gen in gen_list:
        gen['P']=np.array(()) 
        gen['Q']=np.array(())    
        gen['Pmax']=np.array(()) 
        gen['caselist']=np.array(())
        for result in results:
            gen['P']=np.append(gen['P'], result[gen['name']][0])
            gen['Q']=np.append(gen['Q'], result[gen['name']][1])
            gen['Pmax']=np.append(gen['Pmax'], result[gen['name']][2])
            gen['caselist']=np.append(gen['caselist'], result[gen['name']][3])
        gen['P_avg']=np.average(gen['P'])
        gen['Q_avg']=np.average(gen['Q'])
        gen['Pmax_avg']=np.average(gen['Pmax'])
        gen['cap_fact']=gen['P_avg']/gen['Pmax_avg']
    p.terminate()
    print("done")    
            
def check_capacity_factor_subpr(args):
    print(args[1][-34:])
    case=args[1]
    gen_list=args[2]
    result={}
    with silence():
        psspy.read(0,case)
    for gen in gen_list:
        p_out=0
        q_out=0
        p_max=0
        for machine in gen['machines']:
            with silence():
                ierr, p_out_part = psspy.macdat(machine[0],str(machine[1]),'P')
                ierr, q_out_part = psspy.macdat(machine[0],str(machine[1]),'Q')
                ierr, p_max_part = psspy.macdat(machine[0],str(machine[1]),'PMAX')
            if(p_out_part is not None): p_out=p_out+p_out_part
            if(q_out_part is not None): q_out=q_out+q_out_part
            if(p_max_part is not None): p_max=p_max+p_max_part
        result[gen['name']]=[p_out, q_out, p_max, args[0]] #return filename as well to be able to map it to generator output
    return result

def check_line_loading_par(line_list, case_dir):
    case_list=[]
    for dirpath, dirnames, filenames in os.walk(case_dir):
        for filename in filenames:
            filepath=os.path.join(dirpath,filename)
            if filepath.endswith('.raw'):
                case_list.append([filename,filepath, line_list])
    p=Pool(30)
    print('starting subprocesses')
    results=p.map(check_line_loading_subpr, case_list)
#    results=[]
#    for case in case_list:
#        results.append(check_line_loading_subpr(case))
        
    for line in line_list: 
        line['P']=np.array(()) 
        line['Q']=np.array(())    
        line['MVA']=np.array(()) 
        line['RATING']=np.array(())
        line['PCTMVARATE']=np.array(())
        line['caselist']=np.array(())
        for result in results:
            line['P']=np.append(line['P'], result[line['name']][0])
            line['Q']=np.append(line['Q'], result[line['name']][1])
            line['MVA']=np.append(line['MVA'], result[line['name']][2])
            line['RATING']=np.append(line['RATING'], result[line['name']][3])
            line['PCTMVARATE']=np.append(line['PCTMVARATE'], result[line['name']][4])
            line['caselist']=np.append(line['caselist'], result[line['name']][5])
        line['P_avg']=np.average(line['P'])
        line['Q_avg']=np.average(line['Q'])
        line['MVA_avg']=np.average(line['MVA'])
        line['RATING_avg']=np.average(line['RATING'])
        line['PCTMVARATE_avg']=np.average(line['PCTMVARATE'])
    p.terminate()
    print("done")   


def check_line_loading_subpr(args):
    print(args[1][-34:])
    case=args[1]
    line_list=args[2]
    result={}
    with silence():
        psspy.read(0,case)
    for line in line_list:
        line_info={}
        for frombus in line['frombus']:
            for tobus in line['tobus']:
                line_info_temp=get_branch_info(frombus, tobus, line['id'])
                if(line_info_temp!=[]):
                    line_info=line_info_temp[0]
        if line_info=={}:
            result[line['name']]=[0,0,0,0,0, args[0]]
        else:        
            result[line['name']]=[line_info['P'], line_info['Q'], line_info['MVA'], line_info['RATING'], line_info['PCTMVARATE'], args[0]] #return filename as well to be able to map it to generator output
    return result    


#is provided a list of buses and lines, and a path to a folder with multiple snapshots. 
#The function will then iterate over all the snapshots and determine which of the buses and lines are offline or non-existent in any of the snapshots. 
#A list of dictionaries is returned, containing the information which of the buses and lines are offline or non-existent for each snapshot 
#The analysis is done using parallel processing, making it suitable for large datasets.
def check_bus_and_line_status(asset_list, case_dir):
    case_list=[]
    for dirpath, dirnames, filenames in os.walk(case_dir):
        for filename in filenames:
            filepath=os.path.join(dirpath,filename)
            if filepath.endswith('.raw'):
                case_list.append([filename,filepath, asset_list])
    p=Pool(30)
    print('starting subprocesses')
    results=p.map(check_asset_subpr, case_list)
    p.terminate()
#    results=[]
#    for case_cnt in range(0, len(case_list)): 
#        results.append(check_asset_subpr(case_list[case_cnt]))

    print("done")
    return results
    
def check_asset_subpr(args):
    filename=args[0]
    case=args[1]
    asset_list=args[2]
    print(args[1][-34:])
    with silence():
        psspy.read(0,case)
    result={'name': filename, 'missing_buses':[], 'missing_lines':[], 'offline_buses':[], 'offline_lines':[]}
    buses=[]
    lines=[]
    for item in asset_list:
        if type(item)==int:
            buses.append(item)
        elif(type(item)==list):
            lines.append(item)
    bus_info=get_bus_info(buses, 'TYPE')
    for bus in bus_info.keys():
        if not'TYPE' in bus_info[bus].keys():
            result['missing_buses'].append(bus)
        elif(bus_info[bus]['TYPE']==4):
            result['offline_buses'].append(bus)
    for line in lines:
        line_info=get_branch_info(line[0], line[1], str(line[2]),1)
        if(line_info==[]):
            result['missing_lines'].append(line)
        elif(line_info[0]['STATUS']!=1):    
            result['offline_lines'].append(line)
    return result
        
    
    
    

def create_slack_bus_vic():
    #add slack bus at interconnector to Tasmania (Yallourn is common choice for slack bus in hourly snapshots).
    #in order to avoid checking status of existing units, a new machine is added at 35447 (Main Yallourn busbar) --> check include check routine in final simulation script to check if that bus is disconencted in any of the snapshots
    LOY_bus_data=get_bus_info(348090,['PU','ANGLED'])
    psspy.bus_data_4(348099,0,[3,3,1,1],[500.0, LOY_bus_data[348090]['PU'],LOY_bus_data[348090]['ANGLED'], 1.1, 0.9, 1.1, 0.9],"") #make sure to add bus as victorian bus to ensure convergence
    psspy.branch_data_3(348090,348099,r"""1""",[1,348090,1,0,0,0],[0.0, 0.001,0.0,0.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0],[4000.0,4000.0,4000.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
    psspy.plant_data_4(348099,0,[0,0],[ LOY_bus_data[348090]['PU'], 100.0])
    psspy.machine_data_2(348099,r"""1""",[1,1,0,0,0,0],[0.0,0.0, 999.0,-999.0, 999.0,-999.0, 100.0,0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])


def create_slack_bus_nsw():
    #add slack bus at interconnector to Tasmania (Yallourn is common choice for slack bus in hourly snapshots).
    #in order to avoid checking status of existing units, a new machine is added at 35447 (Main Yallourn busbar) --> check include check routine in final simulation script to check if that bus is disconencted in any of the snapshots
    MTP_bus_data=get_bus_info(257691,['PU','ANGLED'])
    psspy.bus_data_4(257699,0,[3,2,1,1],[500.0, MTP_bus_data[257691]['PU'],MTP_bus_data[257691]['ANGLED'], 1.1, 0.9, 1.1, 0.9],"") #make sure to add bus as victorian bus to ensure convergence
    psspy.branch_data_3(257691,257699,r"""1""",[1,257691,1,0,0,0],[0.0, 0.001,0.0,0.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0],[4000.0,4000.0,4000.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
    psspy.plant_data_4(257699,0,[0,0],[ MTP_bus_data[257691]['PU'], 100.0])
    psspy.machine_data_2(257699,r"""1""",[1,1,0,0,0,0],[0.0,0.0, 999.0,-999.0, 999.0,-999.0, 100.0,0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
    
def set_slack_bus_nsw():
    slack_buses=identify_slack_buses()    
    for bus in slack_buses.keys():
        if( (slack_buses[bus]['Zone']!=2) or (bus!=257699) ):
            psspy.bus_chng_4(bus, 0, [2,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s) #    
    psspy.bus_chng_4(257699, 0, [3,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s) #
    
def set_slack_bus_vic():
    slack_buses=identify_slack_buses()    
    for bus in slack_buses.keys():
        if( (slack_buses[bus]['Zone']!=3) or (bus!=348099) ):
            psspy.bus_chng_4(bus, 0, [2,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s) #    
    psspy.bus_chng_4(348099, 0, [3,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s) #

#check output of list of generators as a percentage of their maximum output capability. This is used to determine the output of pre-determined reference projects.    
def check_ref_output(ref_gens, summary=False):
    result={}
    for state in ref_gens.keys():
        result[state]={}
        if(summary): print(state+':')
        for gen_type in ref_gens[state].keys():            
            gen=ref_gens[state][gen_type]        
            p_out=0
            q_out=0
            p_max=0
            for machine in gen['machines']:
                with silence():
                    ierr, p_out_part = psspy.macdat(machine[0],str(machine[1]),'P')
                    ierr, q_out_part = psspy.macdat(machine[0],str(machine[1]),'Q')
                    ierr, p_max_part = psspy.macdat(machine[0],str(machine[1]),'PMAX')
                if(ierr==2 and machine[1]==1):# in case machine number does not exist at the bus, try some other numbers at the same bus in case it has changed.
                    mac_id=2
                    while(mac_id<5 and ierr==2):
                        ierr, p_out_part = psspy.macdat(machine[0],str(mac_id),'P')
                        ierr, q_out_part = psspy.macdat(machine[0],str(mac_id),'Q')
                        ierr, p_max_part = psspy.macdat(machine[0],str(mac_id),'PMAX')
                        mac_id+=1                        
                if(p_out_part is not None): p_out=p_out+p_out_part
                if(q_out_part is not None): q_out=q_out+q_out_part
                if(p_max_part is not None): p_max=p_max+p_max_part
            result[state][gen_type]=[p_out, q_out, p_max, p_out/p_max]
            if(summary): print('    '+gen_type+': '+str(round(100*p_out/p_max, 3))+' %')    
    return result
    

    
#define contingencies here for ACCC routine. Potentially write this file manually instead
def create_con_file(raw_file, output_folder):
    pass
    #runback schemes are hard-coded in here
    #raw file might not be needed, network extensions will be automatically included (or not included) by virtue of their zone number
    #include contingencies on the 66 kV network as well.

#define the bus subsystems for the ACCC routine, defining where faults should be applied, which lines should be monitored, and which generators can be redispatched.
#it may be necessary to create this file during runtime, to reflect the network of conditions of the respective case, which can contain differnet generators and network upgrades.
def create_subs_file(output_folder, year, optimism, timestamp):
    if not (os.path.isfile(output_folder+"\\subsystems"+str(year)+optimism+timestamp+".sub")):
        fault_exclusion_list=[258543, 234041] #buses excluded because the lines connecting to these buses are already considered as part of an explicitly defined runback scheme in the contingency file.
        subsys_file = open(output_folder+"\\subsystems"+str(year)+optimism+timestamp+".sub", "w") 
        psspy.bsys(sid=0, numarea=2, areas=[2,3] )
        [ierr, [all_buses, types]]=psspy.abusint(0,1,['NUMBER', 'TYPE']) #consider only in-service buses
        [ierr, [base_kV]]=psspy.abusreal(0,1,'BASE')
        gen_buses=[]
        fault_buses=[]
        subsys_file.write("SUBSYSTEM 'ALL_BUSES'\n")
        for bus in all_buses:
            subsys_file.write('    BUS '+str(bus)+'\n')
        subsys_file.write('END\n')
        subsys_file.write("\nSUBSYSTEM 'GEN_BUSES'\n")
        for bus_cnt in range(0, len(all_buses)):
            if(types[bus_cnt]==2):
                gen_buses.append(all_buses[bus_cnt])
                subsys_file.write('    BUS '+str(all_buses[bus_cnt])+'\n')
        subsys_file.write('END\n')
        subsys_file.write("\nSUBSYSTEM 'FAULT_BUSES'\n")
        for bus_cnt in range(0, len(all_buses)):
            if(base_kV[bus_cnt]>33):
                if not (bus in fault_exclusion_list):
                    fault_buses.append(all_buses[bus_cnt])
                    subsys_file.write('    BUS '+str(all_buses[bus_cnt])+'\n')
        subsys_file.write('END\n')
        subsys_file.write('END')
        #file.write("Your text goes here") 
        
        #include categories:
            #all_buses
                #this is the set of buses for which overloading will be monitored. It should include all buses with zone identifier 2 or 3 (VIC or NSW)
                #include all buses that are added as part of network upgrades and/or newly added generators. 
            #generator_buses
                #all buses that I want to allow to be redispatchedw as part of mitigating network overloadings. 
                #This is like a whitelist. It should not include the buses that are representing interconnectors, and also not include buses that are 
            #fault_buses --> contingencies are only applied on the lines connecting the buses of this subset. 
                #It is to contain mainly HV buses, but some MV buses relevant to the generators that are being analysed may be included.
        subsys_file.close() 
        return 0
    else: return 1 #file already exists

def NEMDE_imitation():
    #apply contingencies defined in .con file one by one. This includes runback schemes defined in that file.
    #for each contingency 
        #check line overloadings
        #while line overloadings exist:
            #for worst overloading:
                #check influence of every generator on the overloading, and determine contribution factors. 
                #Using all generators with contribution factor > 0.07: determine redispatch resolving the overloading for the most overloaded line. (reducin output of some generators and increasing output of generators with the highest negative contribution factor (or lowest positive)if any)
        #record lowest dispatch of every generator in that process.
        #record "constraint on" for every generator in that process
    #save output in easily
    
    #long-term: include voltage stability checks in this process:
    pass    

def resolve_overloading(gens_lower, watchlist, mismatch_pre, backup_path): #add list of generators that can increase their outout as avariable in the medium term.
    #save backup
    psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0,backup_path)
    #The function in its current form is very conservative because it assumes that only the newly added generators woudl be constrained.
    #declare available coal and gas generation in VIC that can be switched off. Determine total combined capacity that can be freed
    merit_order={'VIC':[{'bus':352011,'id':'11','name':'Mortlake1','type':'gas'},      {'bus':352012,'id':'12','name':'Mortlake2','type':'gas'},
                     {'bus':337001,'id':'1','name':'Jeeralang1','type':'gas'},      {'bus':337002,'id':'2','name':'Jeeralang2', 'type':'gas'},      {'bus':337003, 'id':'3','name':'Jeeralang3', 'type':'gas'},     {'bus':337004, 'id':'4','name':'Jeeralang4','type':'gas'}, {'bus':338001, 'id':'1', 'name':'Jeeralang5', 'type':'gas'},{'bus':338002,'id':'2','name':'Jeeralang6', 'type':'gas'},{'bus':338003,'id':'3','name':'Jeeralang7', 'type':'gas'},
                     {'bus':344001,'id':'1','name':'Laverton1', 'type':'gas'},      {'bus':344002,'id':'2','name':'Laverton2', 'type':'gas'},
                     {'bus':382001,'id':'1', 'name':'ValleyPower1', 'type':'gas'},  {'bus':382002,'id':'2', 'name':'ValleyPower2', 'type':'gas'},   {'bus':382003, 'id':'3', 'name':'ValleyPower3', 'type':'gas'},  {'bus':382004, 'id':'4', 'name':'ValleyPower4', 'type':'gas'},{'bus':382005, 'id':'5', 'name':'ValleyPower5', 'type':'gas'},{'bus':382006, 'id':'6', 'name':'ValleyPower6', 'type':'gas'},
                     {'bus':370001,'id':'1', 'name':'Somerton1', 'type':'gas'},     {'bus':370002,'id':'2', 'name':'Somerton2', 'type':'gas'},      {'bus':370003, 'id':'3', 'name':'Somerton3', 'type':'gas'},     {'bus':370004, 'id':'4', 'name':'Somerton4', 'type':'gas'},
                     {'bus':308001,'id':'1','name':'Bairnsdale1','type': 'gas'},    {'bus':308002,'id': '2','name':'Bairnsdale2', 'type':'gas'},
                     {'bus':360001,'id':'1','name':'Newport', 'type':'gas'},
                     {'bus':346001,'id':'1', 'name':'LoyYang1', 'type':'coal'},     {'bus':346002, 'id':'2', 'name':'LoyYang2', 'type':'coal'},     {'bus':346003, 'id':'3', 'name':'LoyYang3', 'type':'coal'},     {'bus':346004, 'id':'4', 'name':'LoyYang4', 'type':'coal'}, {'bus':347001, 'id':'1', 'name':'LoyYang5', 'type':'coal'},{'bus':347002,'id':'2', 'name':'LoyYang6', 'type':'coal'},
                     {'bus':390001,'id':'1', 'name':'Yallourn1', 'type':'coal'},    {'bus':390002, 'id':'2', 'name':'Yallour21', 'type':'coal'},    {'bus':390003, 'id':'3', 'name':'Yallourn3', 'type':'coal'},    {'bus':390004, 'id':'4', 'name':'Yallourn4', 'type':'coal'} #list of most expensive generators to switch off if generation too high, most expensive generator comes first.
                        ],
    #determine and declare availabel coal and gas generation in NSW that can be switched off. Determine total combined capacity that can be freed
    'NSW':[                    #[218801, '1', 'Blowering', 'hydro'],
                    #[221205, '3', 'Burrinjuck1', 'hydro'],[221205, '4', 'Burrinjuck2', 'hydro'],[221205, '5', 'Burrinjuck3', 'hydro'],
                    #[239201, '1' 'Guthega1', 'hydro'], [239202, '2' 'Guthega2', 'hydro'],
                    #[240801, '1' 'Hume1, 'hydro'], [240802, '2', 'Hume2', 'hydro'],
                    #[242801, '1', 'Jounama', 'hydro'],
                    {'bus':227601, 'id':'1', 'name':'Colongra1', 'type':'gas'},     {'bus':227602, 'id':'2', 'name':'Colongra1', 'type':'gas'}, {'bus':227603, 'id':'3', 'name':'Colongra1', 'type':'gas'}, {'bus':227604, 'id':'4', 'name':'Colongra1', 'type':'gas'},
                    {'bus':219601, 'id':'1', 'name':'BrokenHill1', 'type':'gas'},   {'bus':219602, 'id':'2', 'name':'BrokenHill2', 'type':'gas'},
                    {'bus':234801, 'id':'1', 'name':'Gadara', 'type':'gas'}, #assuming gas. No info online
                    {'bus':241201, 'id':'1', 'name':'HunterValley1', 'type':'gas'}, {'bus':241202, 'id':'2', 'name':'HunterValley2', 'type':'gas'},
                    {'bus':270801, 'id':'1', 'name':'Smithsfield1', 'type':'gas'},  {'bus':270802, 'id':'2', 'name':'Smithsfield2', 'type':'gas'}, {'bus':270803, 'id':'3', 'name':'Smithsfield3', 'type':'gas'}, {'bus':270804, 'id':'4', 'name':'Smithsfield4', 'type':'gas'},
                    {'bus':274801, 'id':'1', 'name':'Tallawarra', 'type':'gas'}, #Combined Cycle
                    {'bus':279611, 'id':'1', 'name':'Uranquinty1', 'type':'gas'},   {'bus':279612, 'id':'1', 'name':'Uranquinty2', 'type':'gas'}, {'bus':279613, 'id':'1', 'name':'Uranquinty3', 'type':'gas'}, {'bus':279614, 'id':'1', 'name':'Uranquinty4', 'type':'gas'}, 
                    {'bus':233201, 'id':'1', 'name':'Eraring1', 'type':'coal'},     {'bus':233202, 'id':'2', 'name':'Eraring2', 'type':'coal'}, {'bus':233203, 'id':'3', 'name':'Eraring3', 'type':'coal'}, {'bus':233211, 'id':'11', 'name':'Eraring1', 'type':'coal'},
                    {'bus':215201, 'id':'1', 'name':'Baywswater1', 'type':'coal'},  {'bus':215202, 'id':'2', 'name':'Baywswater2', 'type':'coal'}, {'bus':215203, 'id':'3', 'name':'Baywswater3', 'type':'coal'},{'bus':215204, 'id':'4', 'name':'Baywswater4', 'type':'coal'},
                    {'bus':249201, 'id':'1', 'name':'Liddell1', 'type':'coal'},     {'bus':249202, 'id':'2', 'name':'Liddell2', 'type':'coal'}, {'bus':249203, 'id':'3', 'name':'Liddell3', 'type':'coal'}, {'bus':249204, 'id':'4', 'name':'Liddell4', 'type':'coal'},
                    {'bus':257601, 'id':'1', 'name':'MountPiper1', 'type':'coal'},  {'bus':257602, 'id':'2', 'name':'MountPiper2', 'type':'coal'},
                    {'bus':280005, 'id':'5', 'name':'ValesPoint', 'type':'coal'},   {'bus':280006, 'id':'6', 'name':'ValesPoint', 'type':'coal'},
                    {'bus':216801, 'id':'1', 'name':'Bendeela1', 'type':'stor','Pmin':-40}, {'bus':216801, 'id':'2', 'name':'Bendeela2', 'type':'stor','Pmin':-40}, 
                    {'bus':243203, 'id':'3', 'name':'KangarooValley', 'type':'stor', 'Pmin':-80}, {'bus':243204, 'id':'4', 'name':'KangarooValley', 'type':'stor', 'Pmin':-80},
                    {'bus':251201, 'id':'1', 'name':'LowerTumut1', 'type':'stor', 'Pmin':-10}, {'bus':251202, 'id':'2', 'name':'LowerTumut2', 'type':'stor', 'Pmin':-10},{'bus':251203, 'id':'3', 'name':'LowerTumut3', 'type':'stor', 'Pmin':-10}, {'bus':251204, 'id':'4', 'name':'LowerTumut4', 'type':'stor', 'Pmin':-195}, {'bus':251205, 'id':'5', 'name':'LowerTumut5', 'type':'stor', 'Pmin':-195}, {'bus':251206, 'id':'6', 'name':'LowerTumut6', 'type':'stor', 'Pmin':-195},
                    {'bus':279201, 'id':'1', 'name':'UpperTumut1', 'type':'stor', 'Pmin':-10}, {'bus':279202, 'id':'2', 'name':'UpperTumut2', 'type':'stor', 'Pmin':-10}, {'bus':279203, 'id':'3', 'name':'UpperTumut3', 'type':'stor', 'Pmin':-10}, {'bus':279204, 'id':'4', 'name':'UpperTumut4', 'type':'stor', 'Pmin':-10},#[279205, '5', 'UpperTumut5', 'stor', -0], [279206, '6', 'UpperTumut6', 'stor', -0], [279207, '7', 'UpperTumut7', 'stor', -0], [279208, '8', 'UpperTumut8', 'stor', -0],
                    ]
    }
    #save copy of snapshot
    #initialise watchlist_redisp --> summary of overloaded lines leading to redispatch of gens on watchlist.
    watchlist_redisp={}
    for i in range (0, len(watchlist)): watchlist_redisp[watchlist[i]]={'total':0,'actions':[]}
    #add machines at "regional reference nodes"
    vic_rrn=379081
    nsw_rrn=274490
    bus_info=get_bus_info([nsw_rrn,vic_rrn], ['PU','ANGLED']) #Reference nodes in Sydney West and Thomastown
    if(bus_info[379081]=={}): vic_rrn=379080 #select alternative bus number for Thomastown if
    bus_info=get_bus_info([nsw_rrn,vic_rrn], ['PU','ANGLED'])
    for bus in bus_info.keys():
        if(bus_info[bus]=={}): # in case the regional reference node is not available, stop script execution straight away.

            os.remove(backup_path)#delete backup
            return 2, watchlist_redisp
    psspy.bus_data_4(274499,0,[2,2,1,1],[330.0, bus_info[nsw_rrn]['PU'],bus_info[nsw_rrn]['ANGLED'], 1.1, 0.9, 1.1, 0.9],r"""NSW_RRN""")
    psspy.bus_data_4(379089,0,[2,3,1,1],[220.0, bus_info[vic_rrn]['PU'],bus_info[vic_rrn]['ANGLED'], 1.1, 0.9, 1.1, 0.9],r"""VIC_RRN""")
    psspy.branch_data_3(274499,nsw_rrn,r"""1""",[1,274499,1,0,0,0],[0.0,0.001,0.0,0.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0],[9999.0,9999.0,9999.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
    psspy.branch_data_3(379089,vic_rrn,r"""1""",[1,379089,1,0,0,0],[0.0,0.001,0.0,0.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0],[9999.0,9999.0,9999.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
    psspy.plant_data_3(274499,0,0,[bus_info[nsw_rrn]['PU'],_f])
    psspy.plant_data_3(379089,0,0,[bus_info[vic_rrn]['PU'],_f])
    psspy.machine_data_2(274499,"1",[1,1,0,0,0,0],[0.0,0.0, 4.0,-4.0, 10.0,-10.0, 10.0,0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
    psspy.machine_data_2(379089,"1",[1,1,0,0,0,0],[0.0,0.0, 4.0,-4.0, 10.0,-10.0, 10.0,0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
    
    psspy.fnsl([1,0,0,1,1,0,0,0])
    #read initial slack bus output
    slack_bus_data=identify_slack_buses()  
    slack_bus_nr=slack_bus_data.keys()[0]# assuming only one slack bus is in the system, which should be the case per previous checks
    swing_bus_P_init=slack_bus_data[slack_bus_nr]['P']
    
                                                            #determine overloaded lines (using rating A)
    overloaded_lines=identify_overloaded_lines()

    
    if overloaded_lines==[]:
        os.remove(backup_path)#delete backup
        return -1, watchlist_redisp #no overloaded line
    else:
        init_worst_ovrld=overloaded_lines[0]
        for line in overloaded_lines:
            if (line['PCTMVARATE']>init_worst_ovrld['PCTMVARATE']):
                init_worst_ovrld=line        
        iterations_cnt=0
        summary='ITERATION '+str(iterations_cnt)+'\n'
        #for every overloaded line:
        while overloaded_lines!=[] and iterations_cnt<10: #Maybe re-work this criterion to include permissible tolerance and accept slightly overloaded lines             
            summary='ITERATION '+str(iterations_cnt)+'\n'
            with silence():
                worst_ovrld=overloaded_lines[0]
                for line in overloaded_lines:
                    if (line['PCTMVARATE']>worst_ovrld['PCTMVARATE']):
                        worst_ovrld=line
                if(worst_ovrld['PCTMVARATE']<110):#Initially set to 5%, now tryin 10%
                    os.remove(backup_path)#delete backup
                    return 0, watchlist_redisp #if worst remaining overloading is 110%, consider this a success and abort
                summary+='Worst overload: '+str(worst_ovrld)+'\n'
                #for every generator in "gens_lower"):
                loading_init=psspy.brnflo(worst_ovrld['FROMBUS'], worst_ovrld['TOBUS'], worst_ovrld['ID'])
                loading_init=abs(loading_init[1])
                #establish contribution factor of Regional reference node
                if(str(worst_ovrld['FROMBUS'])[0]=='2'): 
                    rrn_bus=nsw_rrn
                    state='NSW'
                elif(str(worst_ovrld['FROMBUS'])[0]=='3'): 
                    rrn_bus=vic_rrn
                    state='VIC'
                psspy.machine_chng_2(rrn_bus,r"""1""",[_i,_i,_i,_i,_i,_i],[1.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #increase reference node output by 1 MW
                psspy.fnsl([1,0,0,1,1,0,0,0])
                loading=psspy.brnflo(worst_ovrld['FROMBUS'], worst_ovrld['TOBUS'], worst_ovrld['ID'])
                loading=abs(loading[1])
                rrn_factor=loading-loading_init
                summary+='RRN_factor: '+str(rrn_factor)+'\n'
                psspy.machine_chng_2(rrn_bus,r"""1""",[_i,_i,_i,_i,_i,_i],[0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #increase reference node output by 1 MW                
                #calculate contribution factors of other generators
                psspy.fnsl([1,0,0,1,1,0,0,0])
                max_factor=0.001               
                for gen_id in range(0,len(gens_lower)):
                    gen=gens_lower[gen_id]
                    if('Pout' in gen.keys()):
                        if(gen['Pout']>1.0 and gen['sim_status']=='success'):
                           #redispatch generator at 1 MW less and check contribution factor
                           loading_pre=psspy.brnflo(worst_ovrld['FROMBUS'], worst_ovrld['TOBUS'], worst_ovrld['ID'])
                           loading_pre=abs(loading_pre[1])
                           psspy.machine_chng_2(gen['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[gen['Pout']-1.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #decrease machine output by 1 MW                
                           psspy.fnsl([1,0,0,1,1,0,0,0])
                           loading=psspy.brnflo(worst_ovrld['FROMBUS'], worst_ovrld['TOBUS'], worst_ovrld['ID'])
                           loading=abs(loading[1])
                           gen_factor=loading_pre-loading-rrn_factor #calculate contribution factor relative to RRN (regional reference node)
                           gen['cont_factor']=gen_factor #store contribution factor temporarily in generator information
                           if(gen_factor>max_factor): max_factor=gen_factor
                           psspy.machine_chng_2(gen['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[gen['Pout'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #reset machine output               
                           psspy.fnsl([1,0,0,1,1,0,0,0])
                        else: gen['cont_factor']=0
                    else: gen['cont_factor']=0
                summary+='max_factor: '+str(max_factor)+'\n'            
                ovrld_MVA=loading_init-worst_ovrld['RATING']
                summary+='line overloaded by '+str(ovrld_MVA)+'MVA\n'
                number_LHS_gens=0
                for gen in gens_lower: 
                    if('cont_factor' in gen.keys()):
                        if max_factor==0: max_factor=0.0001
                        gen['cont_factor']=gen['cont_factor']/max_factor #normalise contribution factors to 1
                        if(gen['cont_factor']>0.07):
                            number_LHS_gens+=1
                gens_lower = sorted(gens_lower, key=lambda k: k['cont_factor'])
                #go from most contributing generator to least contributing generator and lower output until overloading shoudl be resolved (based on extrapolation from contribution factor)
                total_lower_amount=0
                gen_id=len(gens_lower)-1
                summary+='gens redispatched: \n'
                while ovrld_MVA>0.1 and gen_id>0: 
                    gen=gens_lower[gen_id]
                    if(gen['cont_factor']>=0.07):
                        temp_divisor=(max_factor*gen['cont_factor']+rrn_factor)
                        if temp_divisor==0: temp_divisor=0.0001
                        new_P=gen['Pout']-ovrld_MVA/temp_divisor
                        if new_P<0: new_P=0
                        psspy.machine_chng_2(gen['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[new_P,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #reset machine output               
                        summary+='   '+gen['name']+' from '+str(gen['Pout'])+' to '+str(new_P)+' MW\n'
                        P_diff=gen['Pout']-new_P
                        for watch_gen in watchlist:
                            if watch_gen in gen['name']:
                                watchlist_redisp[watch_gen]['total']-=P_diff
                                watchlist_redisp[watch_gen]['actions'].append({'ovrld_ln':worst_ovrld, 'cont_factor':gen['cont_factor']})
                        total_lower_amount+=P_diff
                        ovrld_MVA-=(P_diff)*(max_factor*gen['cont_factor']+rrn_factor) #determine how much overloading is expected to be reduced by redispatch of generator.
                        gen['Pout']=new_P #double-check that this actually updates the value in the new_gens array!!
                    gen_id-=1
                    
                if(gen_id<0): print("all generators from 'gens_lower' list with contribution factor >0.07 have been redispatched to 0 MW but overloading could not be resolved.")
                psspy.fnsl([1,0,0,1,1,0,0,0])
    
                #increase dispatch of coal_fired power stationsto match total_lower_amount
                swing_bus_P=slack_bus_data=identify_slack_buses()[slack_bus_nr]['P']
                gen_id=len(merit_order[state])-1
                while (swing_bus_P>(swing_bus_P_init+1.0) and gen_id>0 ):
                    gen=merit_order[state][gen_id]
                    #for machines in service: determine contribution factors, scaled by max-factor
                    ierr, P=psspy.macdat(gen['bus'], gen['id'], 'P')                
                    ierr, Pmax=psspy.macdat(gen['bus'], gen['id'], 'PMAX')
                    if(ierr==0) and P<Pmax-1.0:
                        gen['Pout']=P
                        loading_pre=psspy.brnflo(worst_ovrld['FROMBUS'], worst_ovrld['TOBUS'], worst_ovrld['ID'])
                        loading_pre=abs(loading_pre[1])
                        psspy.machine_chng_2(gen['bus'],gen['id'],[_i,_i,_i,_i,_i,_i],[P+1.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #reset machine output               
                        psspy.fnsl([1,0,0,1,1,0,0,0])
                        loading=psspy.brnflo(worst_ovrld['FROMBUS'], worst_ovrld['TOBUS'], worst_ovrld['ID'])
                        loading=abs(loading[1])
                        gen_factor=loading-loading_pre #calculate contribution factor relative to RRN (regional reference node)
                        if max_factor==0: max_factor=0.0001
                        gen['cont_factor']=gen_factor/max_factor #store contribution factor temporarily in generator information
                        psspy.machine_chng_2(gen['bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[gen['Pout'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #reset machine output               
                        psspy.fnsl([1,0,0,1,1,0,0,0])
                        if(gen['cont_factor']<0.04): #if cont_factor ok, increase generator output by
                            new_P=min(Pmax, P+(swing_bus_P-swing_bus_P_init))
                            psspy.machine_chng_2(gen['bus'],gen['id'],[_i,_i,_i,_i,_i,_i],[new_P,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #reset machine output               
                            summary+='   '+gen['name']+' from '+str(gen['Pout'])+' to '+str(new_P)+' MW\n'
                        psspy.fnsl([1,0,0,1,1,0,0,0])
                        swing_bus_P=slack_bus_data=identify_slack_buses()[slack_bus_nr]['P']
                    gen_id-=1
                
                summary+='initial swing bus out: '+str(swing_bus_P_init)+'\n'
                summary+='final swing bus out: '+str(swing_bus_P)+'\n'            
                iterations_cnt+=1
                overloaded_lines=identify_overloaded_lines()
            print(summary)
        #If overloadings completely resolved
        if overloaded_lines==[]:
            mismatch=test_convergence()
            if(mismatch>mismatch_pre+1 or mismatch>1.0):
                watchlist_redisp={}
                for i in range (0, len(watchlist)): watchlist_redisp[watchlist[i]]={'total':0,'actions':[]}
                psspy.read(0,backup_path) #load previous case if function has increased mismatch
                os.remove(backup_path)#delete backup
                return 1, watchlist_redisp
            else:
                os.remove(backup_path)#delete backup                
                return 0, watchlist_redisp
        #If overloadings could not be resolved
        else:
            mismatch=test_convergence()
            worst_ovrld=overloaded_lines[0]
            for line in overloaded_lines:
                if (line['PCTMVARATE']>worst_ovrld['PCTMVARATE']):
                    worst_ovrld=line
            if(worst_ovrld['PCTMVARATE']>init_worst_ovrld['PCTMVARATE']) or (mismatch>mismatch_pre+1.0) or (mismatch>1.0):
                watchlist_redisp={}
                for i in range (0, len(watchlist)): watchlist_redisp[watchlist[i]]={'total':0,'actions':[]}
                psspy.read(0,backup_path) #load previous case if function has increased mismatch or has worsened overloading
            os.remove(backup_path)#delete backup
            return 1, watchlist_redisp
#takes list of generators (containing machine buses belonging to one generator and point of connetion) 
#detects if there is any capacitors or reactors connected. 
#aggregates the branches to create one equivalent model
    #Goal: end up with a simplified network model in which some generators are replaced by simple power sources, which will hopefully help with convergence.
def simplify_gens():
    pass

#replace all pwoer plant modelw with simplified ones, connecting directly into HV. 
#start with P, Q=0 at every plant and then slowly apply power step changes to match the intended power flows from every plant. 

def simplify_network():
    pass

def aprint(print_array):
    for element in range(0, len(print_array)):
        print(print_array[element])
                
def dprint(print_dict):
    for key in dict.keys():
        print(print_dict[key])
        

def getFaultLvl(psseLog):
    line_counter=len(psseLog)-1
    while line_counter>0 and not ('X------------ BUS ------------X'  in psseLog[line_counter]):
        line_counter-=1
    if(line_counter>0):
        return float(psseLog[line_counter+1][38:47])
    else:
        return -1
        

#takes a set of lines and generators as input. For each of those lines, the function will check whether it is overloaded, and if so, 
#the function will calculate the contribution factors of the generators in the list and return then in descending order for each overloaded line. 
#lines=[{'name': 'example_line', 'frombus': 123456, 'tobus':654321, 'id':'1'}, ...]    , 'gens'=[{'name':'exampleGen', 'bus', 'id':'1'}, ...]  
def calc_contribution_factors(lines, gens, delta=1.0):
    #save ref_snapshot
    with silence():
        results={}
        psspy.save(r'intermed.sav')
        for line in lines: 
            line_info=get_branch_info(line['frombus'], line['tobus'])
            if(line_info!=[]):
                if(line_info[0]['PCTMVARATE']>100):
                    line_shrt=line['name']+"_"+str(line['frombus'])+"-"+str(line['tobus'])
                    results[line_shrt]=[]
                    line_loading_pre=line_info[0]['MVA']
                    for gen in gens:
                        ierr, P=psspy.macdat(gen['bus'], gen['id'], 'P')
                        if((P>delta) and (ierr==0)):
                            psspy.machine_chng_2(gen['bus'],gen['id'],[_i,_i,_i,_i,_i,_i],[P-delta,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            ierr = psspy.fnsl([0,0,0,1,1,0,99,0])#loadflow locked taps
                            ierr = psspy.fnsl([0,0,0,1,1,0,99,0])
                            line_loading_post=get_branch_info(line['frombus'], line['tobus'])[0]['MVA']
                            contribution_factor=(line_loading_pre-line_loading_post)/delta
                            results[line_shrt].append([gen['name'], contribution_factor])
                            psspy.case(r'intermed.sav')
    # print results
    print("contribution Factors have been calculated using active power steps of delta = "+str(delta)+" MW")
    for line in results.keys():
        print(line+":")
        contribution_factors=results[line]
        contribution_factors=sorted(contribution_factors, key = lambda x:x[1], reverse=True )
        for gen in contribution_factors:
            print("    "+gen[0]+": "+str(round(gen[1],5)))
    return results

        
    

    
    
if __name__=='__main__':
    mp.freeze_support()
    
        
        
            
    
    