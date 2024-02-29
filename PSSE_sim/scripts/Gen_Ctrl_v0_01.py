# -*- coding: utf-8 -*-
"""
Created on Thu Apr 14 10:19:35 2022

@author: Mani Aulakh

FUNCTIONALITY:
    The script set up the plants into the desired control system which are anot availble in the PSSE module. Following functionality is provided by the script:
        1. Power factor control
        2. Voltage droop control
        3. Fix P or Q at the POC
COMMENTS:
    1. Inputs are required from the user within this script itself.
    2. This script is called by the Steady State Analysis script.

"""
from datetime import datetime
start_time = datetime.now()
#Do work here!


#Define PSS/E  local path
import sys
import os

sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
sys.path.append(sys_path_PSSE)
os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE


#module intialisation
import psspy
import redirect
redirect.psse2py()
psspy.psseinit(10000)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()
redirect.psse2py()


# Locating required existing folder paths
main_folder_path=os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
sys.path.append(main_folder_path+"\\PSSE_sim\\scripts\\Libs")

# Import required libraries
import auxiliary_functions as af
import numpy as np


# USER INPUTS REQUIRED

# Input data for the plants that requires Power Factor Control
#Note: POC bus is the poc bus, wheras ibus is the bus before poc when doing walkthrough from generator to POC, hence P is negative.
gens_with_pf = {'gens':[{'gen_bus':30714,'gen_id':'1','poc_bus':36711,'ibus':36712,'poc_pf':-0.9600001,'gen_q_max':21.1857,'gen_p_gen':-66.00},
                         {'gen_bus':30715,'gen_id':'1','poc_bus':36711,'ibus':36713,'poc_pf':-0.960001,'gen_q_max':17.6143,'gen_p_gen':-34.00},
                         {'gen_bus':30709,'gen_id':'1','poc_bus':36716,'ibus':36717,'poc_pf':-0.990001,'gen_q_max':19.25625,'gen_p_gen':-48.75},
                         {'gen_bus':30710,'gen_id':'1','poc_bus':36716,'ibus':36718,'poc_pf':-0.990001,'gen_q_max':10.36875,'gen_p_gen':-26.25},
                         {'gen_bus':302,'gen_id':'1','poc_bus':1,'ibus':102,'poc_pf':-0.995,'gen_q_max':30.0276,'gen_p_gen':-76},
                         #{'gen_bus':367169,'gen_id':'1','poc_bus':36721,'ibus':367167,'poc_pf':-0.96,'gen_q_max':15.8,'gen_p_gen':-40},
                         #{'gen_bus':367166,'gen_id':'1','poc_bus':36721,'ibus':367164,'poc_pf':-0.96,'gen_q_max':15.8,'gen_p_gen':-40}
                         ]}
    

# Input data for the plants that requires Voltage droop control    
gens_with_vdc = {'gens':[{'gen_bus':9942,'gen_id':'1','poc_bus':9910,'ibus':9920,'branch_id':'1','poc_trgt_volt':1.02,'gen_q_max':35.55,'gen_droop':0.0502,},
                         #{'gen_bus':334093,'gen_id':'1','poc_bus':334081,'ibus':334090,'branch_id':'1','poc_trgt_volt':1.02,'gen_q_max':46.926,'gen_droop':0.039,},
                         #{'gen_bus':334095,'gen_id':'1','poc_bus':334081,'ibus':334097,'branch_id':'1','poc_trgt_volt':0.995,'gen_q_max':23.463,'gen_droop':0.039,'gen_p_gen':0.0},
                         ]}

# Input data for the plants that requires hybrid control of Power factor and voltage control. At this stage used     
gens_with_pf_vc = {'gens':[{'gen_bus':30714,'gen_id':'1','poc_bus':36711,'ibus':36712,'poc_pf':-0.96,'gen_q_max':33.5,'gen_p_gen':-66.00},
                         {'gen_bus':30715,'gen_id':'1','poc_bus':36711,'ibus':36713,'poc_pf':-0.96,'gen_q_max':17,'gen_p_gen':-34.00},
                         {'gen_bus':302,'gen_id':'1','poc_bus':1,'ibus':102,'poc_pf':-0.995,'gen_q_max':47,'gen_p_gen':-76},
                         {'gen_bus':100001,'gen_id':'1','poc_bus':37703,'ibus':10001,'poc_pf':-0.99,'gen_q_max':3.5,'gen_p_gen':0.0},
                         #{'gen_bus':100001,'gen_id':'1','poc_bus':37703,'ibus':100001,'poc_pf':0.0,'gen_q_max':3.5,'gen_p_gen':0.0}
                         #{'gen_bus':367169,'gen_id':'1','poc_bus':36721,'ibus':367167,'poc_pf':-0.985,'gen_q_max':15.8,'gen_p_gen':-40},
                         #{'gen_bus':367166,'gen_id':'1','poc_bus':36721,'ibus':367164,'poc_pf':-0.985,'gen_q_max':15.8,'gen_p_gen':-40}
                         ]}
    
# Input data for the plants that requires Q to be fixed at the POC when the plant is in voltage control within PSSE module    
gens = {'gens':[#{'poc':367171,'mc_bus':367176,'mc_id':'1','mc_string':'VSCHED','ibus':36717,'Qtrgt':15.8},
                    #{'poc':367172,'mc_bus':367179,'mc_id':'1','mc_string':'VSCHED','ibus':36717,'Qtrgt':15.8},
                    {'poc':36713,'mc_bus':30715,'mc_id':'1','mc_string':'VSCHED','ibus':36711,'Qtrgt':21.1857},
                    {'poc':36712,'mc_bus':30714,'mc_id':'1','mc_string':'VSCHED','ibus':36711,'Qtrgt':17.6143},
                    #{'poc':2103,'mc_bus':2100,'mc_id':'1','mc_string':'VSCHED','ibus':1000,'Qtrgt':7.2190},
                    #{'poc':1103,'mc_bus':1100,'mc_id':'1','mc_string':'VSCHED','ibus':1000,'Qtrgt':20.7640},
                    {'poc':102,'mc_bus':302,'mc_id':'1','mc_string':'VSCHED','ibus':1,'Qtrgt':30.0276},
                    {'poc':37703,'mc_bus':100001,'mc_id':'1','mc_string':'VSCHED','ibus':100001,'Qtrgt':3.5}
                    ]}

    
    
    
    
# This Function intialise the plants at fixed power factor control     
def init_gens_pf(gens_with_pf):
    if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
    for i in range(0,5): 
        for gen in gens_with_pf.keys():
         for gen in gens_with_pf['gens']:
             ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
             
             if ival == 0:
                 print('GEN is OFF')
             else:
    
                 ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                 q_poc = ibranch_inf.imag # poc q mvar
                 p_poc = ibranch_inf.real # poc p mvar
                 
                 delta_p = p_poc - gen['gen_p_gen']
                 q_poc_req = np.sqrt((abs(gen['gen_p_gen'])/abs(gen['poc_pf']))**2-(abs(gen['gen_p_gen']))**2)
                 delta_q = q_poc - q_poc_req
                 
                 ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q mvar
                 ierr,p_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'P') # gen p mvar
                 
                 if abs(q_poc) > gen['gen_q_max']:
                     for i in range(0,3):
                         ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                         q_poc = ibranch_inf.imag # poc q mvar
                         if abs(q_poc) > gen['gen_q_max']:
                             q_diff = gen['gen_q_max'] - abs(q_poc)
                             ierr, mc_q_max=psspy.macdat(gen['gen_bus'],gen['gen_id'],'QMAX')
                             #q_max =  mc_q_max + q_diff
                             q_min =  -mc_q_max - q_diff
                             psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[p_gen+delta_p,_f, q_min,q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                             if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                    
                     ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                     q_poc = ibranch_inf.imag # poc q mvar
                     p_poc = ibranch_inf.real # poc p mvar
                     
                     delta_p = p_poc - gen['gen_p_gen']
                     q_poc_req = np.sqrt((abs(gen['gen_p_gen'])/abs(gen['poc_pf']))**2-(abs(gen['gen_p_gen']))**2)
                     delta_q = q_poc - q_poc_req
                     
                     ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q mvar
                     ierr,p_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'P') # gen p mvar
                     psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ p_gen+delta_p,_f, (q_gen+delta_q),(q_gen+delta_q) ,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                     if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                 else:
                    psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[ p_gen+delta_p,_f, (q_gen+delta_q),(q_gen+delta_q) ,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
    
    


        
# This function intialise plant at the required voltage droop control
def init_gens_vdc(gens_with_vdc):
    if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
    #if(af.test_convergence(method='fdns',taps='step')>1.0):raise
    for i in range(0,25): 
        for gen in gens_with_vdc.keys():
         for gen in gens_with_vdc['gens']:
             ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
             
             if ival == 0:
                 print('GEN is OFF')
             else:
                 #for i in range(0,5):
                 poc_volt_lvl = af.get_bus_info(gen['poc_bus'],'PU')
                 poc_volt_lvl = poc_volt_lvl[gen['poc_bus']]['PU'] #poc volt level
                 
                 ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                 q_poc = ibranch_inf.imag # poc q mvar
                 
                 ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q mvar
                 
                 if gen['poc_trgt_volt']>poc_volt_lvl: # Gen has to inject reactive power
                     delta_v =  poc_volt_lvl - gen['poc_trgt_volt']
                 #q_poc_req = 1#
                     q_poc_req = ((gen['gen_q_max']/gen['gen_droop']) * delta_v)  # 2 gens are modelled
                     if abs(q_poc_req) > gen['gen_q_max']:
                         q_diff = gen['gen_q_max'] - abs(q_poc)
                         q_max = abs(q_gen) + q_diff
                         psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_max, q_max,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                         if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                         for i in range(0,3):
                             ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                             q_poc = ibranch_inf.imag # poc q mvar
                             if abs(q_poc) > gen['gen_q_max']:
                                 ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q')
                                 q_diff = abs(q_poc) - gen['gen_q_max']
                                 q_max = abs(q_gen)  - q_diff
                                 psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_max, q_max,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                                 if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                     else:
                         delta_q = ((q_poc - q_poc_req)/10)
                         if abs(q_poc - q_poc_req)>0.1:
                             psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, (q_gen+delta_q), (q_gen+delta_q),_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                             if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                         else:
                            pass
                         #if(af.test_convergence(method='fdns',taps='step')>1.0):raise
                     
                 elif gen['poc_trgt_volt']<poc_volt_lvl: # Gen has to absorb reactive power
                     delta_v =  poc_volt_lvl - gen['poc_trgt_volt']
                 #q_poc_req = 1#
                     q_poc_req = ((gen['gen_q_max']/gen['gen_droop']) * delta_v)  # 2 gens are modelled
                     if abs(q_poc_req) > gen['gen_q_max']:
                         q_diff = gen['gen_q_max'] - abs(q_poc)
                         q_min = -abs(q_gen) - q_diff
                         psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_min, q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                         if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                         for i in range(0,4):
                             ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                             q_poc = ibranch_inf.imag # poc q mvar
                             if abs(q_poc) > gen['gen_q_max']:
                                 ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q')
                                 q_diff = abs(q_poc) - gen['gen_q_max']
                                 q_min = -(abs(q_gen)  - q_diff)
                                 psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_min, q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                                 if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                        
                     else:
                         delta_q = ((q_poc - q_poc_req)/10)
                         if abs(q_poc - q_poc_req)>0.1:
                             psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, (q_gen+delta_q), (q_gen+delta_q),_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                             if(af.test_convergence(method='fnsl',taps='step')>1.0):raise
                             #if(af.test_convergence(method='fdns',taps='step')>1.0):raise
                         else:
                            pass
                        
# This function provides the voltage droop control functionality with locked taps    
def lckd_gens_vdc(gens_with_vdc):
    
    if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
    for i in range(0,3): 
        for gen in gens_with_vdc.keys():
         for gen in gens_with_vdc['gens']:
             ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
             
             if ival == 0:
                 print('GEN is OFF')
             else:
                 #for i in range(0,5):
                 poc_volt_lvl = af.get_bus_info(gen['poc_bus'],'PU')
                 poc_volt_lvl = poc_volt_lvl[gen['poc_bus']]['PU'] #poc volt level
                 
                 ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                 q_poc = ibranch_inf.imag # poc q mvar
                 
                 ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q mvar
                 
                 if gen['poc_trgt_volt']>poc_volt_lvl: # Gen has to inject reactive power
                     delta_v =  poc_volt_lvl - gen['poc_trgt_volt']
                 #q_poc_req = 1#
                     q_poc_req = ((gen['gen_q_max']/gen['gen_droop']) * delta_v)  # 2 gens are modelled
                     if abs(q_poc_req) > gen['gen_q_max']:
                         q_diff = gen['gen_q_max'] - abs(q_poc)
                         q_max = abs(q_gen) + q_diff
                         psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_max, q_max,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                         if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                         for i in range(0,3):
                             ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                             q_poc = ibranch_inf.imag # poc q mvar
                             if abs(q_poc) > gen['gen_q_max']:
                                 ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q')
                                 q_diff = abs(q_poc) - gen['gen_q_max']
                                 q_max = abs(q_gen)  - q_diff
                                 psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_max, q_max,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                                 if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                     else:
                         delta_q = ((q_poc - q_poc_req)/2) # half of the delta q is implemented and further need to distributed to 2 gens
                         psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, (q_gen+delta_q), (q_gen+delta_q),_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                         if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                     
                 elif gen['poc_trgt_volt']<poc_volt_lvl: # Gen has to absorb reactive power
                     delta_v =  poc_volt_lvl - gen['poc_trgt_volt']
                 #q_poc_req = 1#
                     q_poc_req = ((gen['gen_q_max']/gen['gen_droop']) * delta_v)  # 2 gens are modelled
                     if abs(q_poc_req) > gen['gen_q_max']:
                         q_diff = gen['gen_q_max'] - abs(q_poc)
                         q_min = -abs(q_gen) - q_diff
                         psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_min, q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                         if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                         for i in range(0,3):
                             ibranch_inf=psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                             q_poc = ibranch_inf.imag # poc q mvar
                             if abs(q_poc) > gen['gen_q_max']:
                                 ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q')
                                 q_diff = abs(q_poc) - gen['gen_q_max']
                                 q_min = -(abs(q_gen)  - q_diff)
                                 psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_min, q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                                 if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                     else:
                         delta_q = ((q_poc - q_poc_req)/2) # half of the delta q is implemented and further need to distributed to 2 gens
                         psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, (q_gen+delta_q), (q_gen+delta_q),_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                         if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
        
        
# This function changes the fixed power factor plants to the inbulid psse voltage control scheme, hence setting up the plants with hybrid control scheme                     
def init_gens_hc(gens_with_pf_vc):
    
    for gen in gens_with_pf_vc.keys():
     for gen in gens_with_pf_vc['gens']:
         ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
         
         if ival == 0:
             print('GEN is OFF')
         else:
             
             poc_volt_lvl = af.get_bus_info(gen['poc_bus'],'PU')
             poc_volt_lvl = poc_volt_lvl[gen['poc_bus']]['PU'] #poc volt level
             ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q mvar
             psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, gen['gen_q_max'],-gen['gen_q_max'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
             psspy.plant_data_4(gen['gen_bus'],0,[_i,_i],[ poc_volt_lvl,gen['poc_bus']])
    
               
     if(af.test_convergence(method='fnsl',taps='step')>1.0):raise

     
# This function fixes the Q at POC for the plants that are in direct voltage control scheme                
def fix_q_poc():
    
    for gen in gens.keys():
        for gen in gens['gens']:
            poc_volt = af.get_bus_info(gen['poc'],'PU')
            poc_volt = poc_volt[gen['poc']]['PU']
            ierr,vsched = psspy.macdat(gen['mc_bus'],gen['mc_id'],gen['mc_string'])
            ierr,ival = psspy.macint(gen['mc_bus'],gen['mc_id'],'STATUS')
            
            if ival == 0:
                print('M/C is OFF')
                pass
            else:
                for i in range(0,3):
                    ibranch_inf=psspy.brnflo(gen['poc'],gen['ibus'],'1')[1]
                    qvar_poc = ibranch_inf.imag
                    poc_volt = af.get_bus_info(gen['poc'],'PU')
                    poc_volt = poc_volt[gen['poc']]['PU']
                    if abs(poc_volt-vsched)<0.001 and (abs(qvar_poc)>gen['Qtrgt']):
                        q_diff = gen['Qtrgt'] - abs(qvar_poc)
                        ierr, mc_q_max=psspy.macdat(gen['mc_bus'],gen['mc_id'],'QMAX')
                        q_max =  mc_q_max + q_diff
                        q_min =  -mc_q_max - q_diff
                        psspy.machine_chng_2(gen['mc_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_max,q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
                      
                    else:
                        q_diff = gen['Qtrgt'] - abs(qvar_poc)
                        ierr, mc_q_max=psspy.macdat(gen['mc_bus'],gen['mc_id'],'QMAX')
                        q_max =  mc_q_max + q_diff
                        q_min =  -mc_q_max - q_diff
                        psspy.machine_chng_2(gen['mc_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_max,q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise

# Run functions
def run():
    path_dir = r'C:\Users\Mani Aulakh\Desktop\Desktop_01\04 HOR\03 Grid\1. Power System Studies\1. Main Test Environment\20220928_HORSF\PSSE_sim\base_model\HighLoad\New folder' # path for the sav files that requires this routine
    ext = ('.sav')
    for sav_file in os.listdir(path_dir):
        if sav_file.endswith(ext):
            print(sav_file)
            psspy.case(path_dir + "\\" + sav_file)
            #init_gens_pf(gens_with_pf)
            init_gens_vdc(gens_with_vdc)
            #lckd_gens_vdc(gens_with_vdc)
            #init_gens_hc(gens_with_pf_vc)
            psspy.save(path_dir + "\\" + sav_file[0:-4] + '_new.sav')
        else:
            continue
#run()


#Do work above
end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))                        