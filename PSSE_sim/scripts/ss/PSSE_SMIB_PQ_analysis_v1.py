"""
Created on Tue Jan 17 13:08:16 2023

@author: ESCO

FUNCTIONALYTY:
The script will run PQ curve using steady state analysis
    + Load the case, set the POC voltage as required (e.g. 0.9, 1.0, 1.1pu), then solve with tap_change on
    + Change the active power as required, solved with tap locked
    + Change the Q level to achive the capability, solve with tap locked

COMMENTS:
    + Check model and path before running the script (check bus number)
    + Check and adapt the mode: PV/BESS or 35degC/50degC
    + Check Pmax, Pmin for PV/BESS mode accordingly
    + Check the buses to be deactivated (only applicable when one of the two branches need to be deactivated.)

@NOTE: 


"""
from __future__ import with_statement
from contextlib import contextmanager

import os, sys
import math
import time
import pandas as pd
import scipy.interpolate



timestr = time.strftime("%Y%m%d-%H%M%S") 
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

redirect.psse2py()
with silence():
    psspy.psseinit(80000)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()

base_dir = '..\SMIB_Model_S5251'
output_loc = os.getcwd()
case_name = '\SUMSF_SMIB_V0.sav'
source_type = "BESS_PV" #"BESS" # "PV" #"BESS_PV"
temperature = "50degC" #"35degC" # "50degC" #"40degC"


temp_name = temperature + "_" + source_type #35degC_BESS"


# Generator (OEM) inputs
PV_INV_num = 24.0 # PV INV quantity aggregated
PV_INV_rate = 4.2 # MVA INV individual rating
BESS_INV_num = 18.0 # BESS INV quantity aggregated
BESS_INV_rate = 3.6 # MVA INV individual rating
    
# SMIB model inputs
tapstep = 1 
bus_inf = 9900 # INF bus to control the voltage for over and under excitation mode
bus_POC_fr = 9920 # POC bus from: connect the branch for measuring power flow
bus_POC_to = 9910 # POC bus to: connect the branch for measuring power flow

# Branch 1 - PV leg
busgen1 = 9942 # Generator bus
busTx1 = 9941
# Branch 2 - BESS leg
busgen2 = 9944
busTx2 = 9943


if source_type == "BESS":
    bus_dis_list = [busgen1, busTx1] # all the buses in the PV leg to be disconnected
    bus_gen_list = [busgen2] # generator bus to be mornitored voltage and power flow
    Pmax_POC = 50   
    Pmin_POC = -50

elif source_type == "PV":
    bus_dis_list = [busgen2, busTx2] # all the buses in the BESS leg to be disconnected
    bus_gen_list = [busgen1]
    Pmax_POC = 90 
    Pmin_POC = 0
    
else: # "BESS_PV": PV and BESS
    bus_dis_list = []
    bus_gen_list = [busgen1, busgen2]  # generator buses to be mornitored voltage and power flow
    Pmax_POC = 90 
    Pmin_POC = -50


Sbase_analysis_PV = PV_INV_num * PV_INV_rate
Sbase_analysis_BESS = BESS_INV_num * BESS_INV_rate

busgen_V_PV = busgen1
busgen_V_BESS = busgen2
    
Pmax_PV = Pmax_POC + 0.7
Pmin_BESS = Pmin_POC + 0.3
Pmax_loop = Pmax_POC + 1.5 # account for the losses
Pmin_loop = Pmin_POC + 0.7 # account for the losses

#Pmax_loop = Sbase_analysis_PV

import readtestinfo as readtestinfo

TestDefinitionSheet = r'20230828_SUM_TESTINFO_V1.xlsx'
#testDefinitionDir= os.path.abspath(os.path.join(main_folder, os.pardir))+"\\test_scenario_definitions"
testDefinitionDir = output_loc

# return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'SimulationSettings', 'ModelDetailsPSSE', 'SetpointsDict', 'ScenariosSMIB', 'Profiles'])
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['PowerCapability'])
#ProjectDetailsDict = return_dict['ProjectDetails']
## SimulationSettingsDict = return_dict['SimulationSettings']
#PSSEmodelDict = return_dict['ModelDetailsPSSE']
#SetpointsDict = return_dict['Setpoints']
#ScenariosDict = return_dict['ScenariosSMIB']
DeratingDict = return_dict['PowerCapability']

#psspy.case(os.path.join(base_dir, case_name))
psspy.case(base_dir + case_name)
#if source_type in ["BESS", "PV"]:
for bus_dis_i in bus_dis_list:
    psspy.dscn(bus_dis_i)
#        psspy.dscn(bus_dis2)

psspy.fnsl([1,0,0,1,1,0,99,0])
psspy.fnsl([1,0,0,1,1,0,99,0])
psspy.fnsl([1,0,0,1,1,0,99,0])

    
#
Vinfinite = [0.90, 1.0, 1.10]


psspy.lines_per_page_one_device(1,60)
psspy.report_output(2,'test.log',[0,0])
psspy.lines_per_page_one_device(1,60)
psspy.progress_output(2,'test.log',[0,0])
psspy.lines_per_page_one_device(1,60)
psspy.alert_output(2,'test.log',[0,0])
psspy.lines_per_page_one_device(1,60)
psspy.prompt_output(2,'test.log',[0,0])    


def monitor_SF(fbus, tbus, bus_POC, bus_gen_list):

    ierr, S_inv_gen1 = psspy.gendat(bus_gen_list[0])
    P_inv_gen1 = S_inv_gen1.real
    Q_inv_gen1 = S_inv_gen1.imag
    ierr, V_inv_gen1 = psspy.busdat(bus_gen_list[0], 'PU')
    if len(bus_gen_list) > 1: #both PV and BESS are considered
        ierr, S_inv_gen2 = psspy.gendat(bus_gen_list[1])
        P_inv_gen2 = S_inv_gen2.real
        Q_inv_gen2 = S_inv_gen2.imag
        ierr, V_inv_gen2 = psspy.busdat(bus_gen_list[1], 'PU')  
    else:
        P_inv_gen2 = 0
        Q_inv_gen2 = 0
        ierr, V_inv_gen2 = 0 
    ierr, cmpval = psspy.brnflo(fbus, tbus, r"""1""")
    P_poc = cmpval.real
    Q_poc = cmpval.imag
    ierr, V_poc = psspy.busdat(bus_POC, 'PU')

    return P_poc, Q_poc, V_poc, P_inv_gen1, Q_inv_gen1, V_inv_gen1, P_inv_gen2, Q_inv_gen2, V_inv_gen2


def iterpolate_derating(derating_profile, P_individual):

    y_interp = scipy.interpolate.interp1d(derating_profile['x_data'], derating_profile['y_data'])
    try: Qlim_out = y_interp(P_individual)
    except: 
        pass
    if P_individual >=0:
        if P_individual <= derating_profile['x_data'][1]: Qlim_out = derating_profile['y_data'][1]
        elif P_individual >= derating_profile['x_data'][-1]: Qlim_out = derating_profile['y_data'][-1]
    else:
        if P_individual >= derating_profile['x_data'][1]: Qlim_out = derating_profile['y_data'][1]
        elif P_individual <= derating_profile['x_data'][-1]: Qlim_out = derating_profile['y_data'][-1]
    Smax_out = abs(derating_profile['x_data'][-1])
    return Qlim_out, Smax_out


def update_Slim(y_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS):
    # Update Sbase with new derating factor based on the terminal voltage level
    try:
        Slim_BESS = y_interp_BESS(Vinv_BESS)
    except: # if the voltage is outside the range [0.9-1.1pu]
        if Vinv_BESS >= 1.1: Slim_BESS = Slim_110_BESS
        elif Vinv_BESS <= 0.9: Slim_BESS = Slim_90_BESS
    return Slim_BESS

P_POC_plot_final = []
Q_POC_plot_final = []

Sbase_PV = Sbase_analysis_PV
Sbase_BESS = Sbase_analysis_BESS

# Define the range for looping P
if Pmin_POC >= 0: #PV
#    Ploop_range = range(0, int(Pmax_POC) + 1, 2) + [Pmax_POC] + range(int(Pmax_POC) + 1, int(Pmax_loop) + 1, 2) + [Pmax_loop]
    Ploop_range = range(0, int(Pmax_POC) + 1, 2) + [Pmax_loop]
else: #BESS
#    Ploop_range =  [-Sbase] + range(int(-Sbase) - 1, int(Pmin_POC) - 1, 2) + [Pmin_POC] + range(int(Pmin_POC) - 1, int(Pmax_POC) + 1, 2) + [Pmax_POC] + range(int(Pmax_POC) + 1, int(Sbase) + 1, 2) + [Sbase]
    Ploop_range =  [Pmin_loop] + range(int(Pmin_loop), int(Pmax_loop), 2) + [Pmax_loop]

writer = pd.ExcelWriter("S5251_" +str(temp_name)+".xlsx", engine = "xlsxwriter") # Preparing for exporting the result
#writer = pd.ExcelWriter("S5251_" +str(temp_name)+str(round(PV_INV_num,0))+str(round(BESS_INV_num,0))+".xlsx", engine = "xlsxwriter") # Preparing for exporting the result

data_out_all = []

for index, Vcp in enumerate(Vinfinite):
    vol_case_name = str(Vcp)
    
    psspy.plant_data(bus_inf,0,[ Vcp,_f])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])    
#    psspy.fnsl([tapstep,0,0,1,1,0,0,0])

    # Force the voltage at POC by changing grid impedance to zero, solve with tap enable 
#        psspy.branch_chng_3(bus_POC_fr,bus_POC_to,r"""1""",[1,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s)
#            psspy.two_winding_chng_5(bus_POC_to,bus_POC_fr,r"""1""",[0,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[ 80.0, 80.0, 80.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s,_s)
    psspy.branch_chng_3(bus_inf,bus_POC_to,r"""1""",[_i,_i,_i,_i,_i,_i],[0.0,0.00001,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
    psspy.fdns([1,0,0,1,1,1,99,0])
    psspy.fdns([1,0,0,1,1,1,99,0])
#    psspy.fnsl([1,0,0,1,1,0,0,0])
#    psspy.fnsl([1,0,0,1,1,0,0,0])
#    psspy.fnsl([1,0,0,1,1,0,0,0])
        
    P_POC_plot = []
    Q_POC_plot = []
    
    P_inv_gen1_all = []
    Q_inv_gen1_all = []
    V_inv_gen1_all = []
    P_inv_gen2_all = []
    Q_inv_gen2_all = []
    V_inv_gen2_all = []
    P_poc_all = []
    Q_poc_all = []
    V_poc_all = []
    Sinv_gen1_all = []
    Iinv_gen1_all = []
    Sinv_gen2_all = []
    Iinv_gen2_all = []
    
    P_inv_gen1_all_pos = []
    Q_inv_gen1_all_pos = []
    V_inv_gen1_all_pos = []
    P_inv_gen2_all_pos = []
    Q_inv_gen2_all_pos = []
    V_inv_gen2_all_pos = []
    P_poc_all_pos = []
    Q_poc_all_pos = []
    V_poc_all_pos = []
    Sinv_gen1_all_pos = []
    Iinv_gen1_all_pos = []
    Sinv_gen2_all_pos = []
    Iinv_gen2_all_pos = []
    
    P_inv_gen1_all_neg = []
    Q_inv_gen1_all_neg = []
    V_inv_gen1_all_neg = []
    P_inv_gen2_all_neg = []
    Q_inv_gen2_all_neg = []
    V_inv_gen2_all_neg = []
    P_poc_all_neg = []
    Q_poc_all_neg = []
    V_poc_all_neg = []
    Sinv_gen1_all_neg = []
    Iinv_gen1_all_neg = []
    Sinv_gen2_all_neg = []
    Iinv_gen2_all_neg = []
    
    for P in Ploop_range:
        if source_type == "BESS":
            P_PV = 0
            P_BESS = P            
            P_individual_PV = P_PV/PV_INV_num #MW
            P_individual_BESS = P_BESS/BESS_INV_num  #MW
        elif source_type == "PV":
            P_PV = P
            P_BESS = 0            
            P_individual_PV = P_PV/PV_INV_num  #MW
            P_individual_BESS = P_BESS/BESS_INV_num  #MW
            
        else: #source_type == "BESS_PV"
    #        if P < Sbase_PV:
            if P <= Pmax_PV and P >= 0: # when PV has capability to provide P
                P_PV = P
                P_BESS = 0
                P_individual_PV = P_PV/PV_INV_num  #MW
                P_individual_BESS = 0 #kW
            elif P > Pmax_PV: # When P is greater than PV capability, limit the PV to max capacity
                P_PV = Pmax_PV
                P_BESS = P-Pmax_PV            
                P_individual_PV = P_PV/PV_INV_num  #MW
                P_individual_BESS = P_BESS/BESS_INV_num  #MW
            elif P < 0: # When P negative - charging side -> P from BESS and Q from PV
                P_PV = 0
                P_BESS = P            
#                if P_BESS > Sbase_analysis_BESS: P_BESS = Sbase_analysis_BESS
                if P_BESS < -Sbase_analysis_BESS: P_BESS = -Sbase_analysis_BESS
                P_individual_PV = 0 #kW
                P_individual_BESS = P_BESS/BESS_INV_num  #MW
                
           
        for Vset_inv in [1.2, 0.7]: # change voltage setpoint of generator to get a Q range. For each case, solve the case with the tap locked
            # Apply the voltage target to all gens
            for gen_i in bus_gen_list:
                psspy.plant_data(gen_i,0,[ Vset_inv,_f])
            soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
            psspy.fnsl([tapstep,0,0,1,1,0,0,0])
            psspy.fnsl([tapstep,0,0,1,1,0,0,0])

            delta_Vinv = 1
            tolerance = 0.001
            iters_out = 0
            P_step = 0.5
            Q_step = 0.2 
                
            # Alocate the derating profile (individual PQ curve) at differnt terminal voltage levels (0.9, 1.0 and 1.1pu)
            ###################################################################
            ###################################################################
            if source_type == "PV":   
                if Vset_inv >1: # when Qgen is positive (over excited)
                    if temperature == "35degC":
                        profile_90_PV = DeratingDict['PV_OE_35_0.9']
                        profile_100_PV = DeratingDict['PV_OE_35_1.0']
                        profile_110_PV = DeratingDict['PV_OE_35_1.1']

                    elif temperature == "40degC":
                        profile_90_PV = DeratingDict['PV_OE_40_0.9']
                        profile_100_PV = DeratingDict['PV_OE_40_1.0']
                        profile_110_PV = DeratingDict['PV_OE_40_1.1']

                    else: #"50degC"
                        profile_90_PV = DeratingDict['PV_OE_50_0.9']
                        profile_100_PV = DeratingDict['PV_OE_50_1.0']
                        profile_110_PV = DeratingDict['PV_OE_50_1.1']
                            
                else: # when Qgen is negative (under excited)
                    if temperature == "35degC":
                        profile_90_PV = DeratingDict['PV_UE_35_0.9']
                        profile_100_PV = DeratingDict['PV_UE_35_1.0']
                        profile_110_PV = DeratingDict['PV_UE_35_1.1']

                    elif temperature == "40degC":
                        profile_90_PV = DeratingDict['PV_UE_40_0.9']
                        profile_100_PV = DeratingDict['PV_UE_40_1.0']
                        profile_110_PV = DeratingDict['PV_UE_40_1.1']
                        
                    else: #"50degC"
                        profile_90_PV = DeratingDict['PV_UE_50_0.9']
                        profile_100_PV = DeratingDict['PV_UE_50_1.0']
                        profile_110_PV = DeratingDict['PV_UE_50_1.1']
                        
                # Estimate the Qlim (individual Qlim) at differnt terminal voltage levels (0.9, 1.0 and 1.1pu) at a given P power level
                Qlim_90_PV, Smax_90_PV = iterpolate_derating(profile_90_PV, P_individual_PV)
                Qlim_100_PV, Smax_100_PV = iterpolate_derating(profile_100_PV, P_individual_PV)
                Qlim_110_PV, Smax_110_PV = iterpolate_derating(profile_110_PV, P_individual_PV)                        
                            
                # Updating the Slim base on the P and Qlim. S suppose to derate linearly with voltage
                Slim_90_PV = math.sqrt(P_individual_PV**2 + Qlim_90_PV**2)
                Slim_100_PV = math.sqrt(P_individual_PV**2 + Qlim_100_PV**2)
                Slim_110_PV = math.sqrt(P_individual_PV**2 + Qlim_110_PV**2)

                # Limit the S if it is over the maximum
                if Slim_90_PV > Smax_90_PV: Slim_90_PV = Smax_90_PV
                if Slim_100_PV > Smax_100_PV: Slim_100_PV = Smax_100_PV
                if Slim_110_PV > Smax_110_PV: Slim_110_PV = Smax_110_PV
                
                # Interporating function of the S on INV terminal voltage level
                S_interp_PV = scipy.interpolate.interp1d([0.9, 1, 1.1], [Slim_90_PV, Slim_100_PV, Slim_110_PV])            
                
                # fetch the terminal voltage for estimating the derating
                ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')

   
                while (delta_Vinv > tolerance) and iters_out < 100: # if the voltage level at inveter varies more than the accepted band
                
                    # Update Sbase with new derating factor based on the terminal voltage level
                    Slim_PV_ind = update_Slim(S_interp_PV,Vinv_PV,Slim_110_PV,Slim_90_PV) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                    if Sbase_PV > abs(P_PV): Qlim_PV = math.sqrt(Sbase_PV**2 - P_PV**2) 
                    else: Qlim_PV = 0
    
                    # Apply the new base
                    psspy.machine_chng_2(busgen_V_PV, '1', [_i,_i,_i,_i,_i,_i], [ P_PV,_f, Qlim_PV, -Qlim_PV,_f,_f, Sbase_PV,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0])    
                
                    # Iteratively converge on maximum operating point by setting Qlim, solving for inverter terminal voltage and re-calculating Qlim                
                    ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                    ierr, Vinv_PV_original = psspy.busdat(busgen_V_PV, 'PU')
                    Slim_PV_ind = update_Slim(S_interp_PV,Vinv_PV,Slim_110_PV,Slim_90_PV) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                    if Sbase_PV > abs(P_PV): Qlim_PV = math.sqrt(Sbase_PV**2 - P_PV**2) 
                    else: Qlim_PV = 0
    
                    ierr, cmpval = psspy.gendat(busgen_V_PV)
                    Pinv_PV = cmpval.real
                    Qinv_PV = cmpval.imag
                    Sinv_PV = math.sqrt(P_PV**2 + Qinv_PV**2)                    
                    Iinv_PV = Sinv_PV / Vinv_PV

                    # If we run into the apparent power limit, start reducing Q, then reducing P.
                    iters = 0
                    print 'SINV_PV = %f' % Sinv_PV
                    while (Sinv_PV > Sbase_PV) and iters < 500:
                        Qlim_PV -= Q_step
                        if Qlim_PV < 0.0: 
                            Qlim_PV = 0.0
                            if P_PV > 0: P_PV -= P_step
#                            elif P_PV < 0: P_PV += P_step
                            if P_PV < 0: P_PV = 0 # PV do not absorb P
                        print '*********************************************************'
                        print 'Apparent power limit reached %1.4f, QlimPV reduced to %1.4f' % (Sinv_PV, Qlim_PV)
                        print '*********************************************************'
                        #
                        psspy.machine_chng_2(busgen_V_PV, '1', [_i,_i,_i,_i,_i,_i], [ P_PV,_f, Qlim_PV, -Qlim_PV,_f,_f, Sbase_PV,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
                        #if soln_ierr:
                        #    break
                        ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                        ierr, cmpval = psspy.gendat(busgen_V_PV)
                        Pinv_PV = cmpval.real
                        Qinv_PV = cmpval.imag
                        Sinv_PV = math.sqrt(Pinv_PV**2 + Qinv_PV**2)
                        Iinv_PV = Sinv_PV / Vinv_PV
                        Slim_PV_ind = update_Slim(S_interp_PV,Vinv_PV,Slim_110_PV,Slim_90_PV) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                        Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                        iters += 1
                
                    time.sleep(0)
                    ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                    delta_Vinv_PV = abs(Vinv_PV_original - Vinv_PV)                
                    delta_Vinv = delta_Vinv_PV
                    iters_out += 1



            ###################################################################
            ###################################################################
            elif source_type == "BESS":   
                if Vset_inv >1: # when Qgen is positive (over excited)
                    if temperature == "35degC":
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-OE_35_0.9']
                            profile_100_BESS = DeratingDict['Dis-OE_35_1.0']
                            profile_110_BESS = DeratingDict['Dis-OE_35_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-OE_35_0.9']
                            profile_100_BESS = DeratingDict['Cha-OE_35_1.0']
                            profile_110_BESS = DeratingDict['Cha-OE_35_1.1']

                    elif temperature == "40degC":
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-OE_40_0.9']
                            profile_100_BESS = DeratingDict['Dis-OE_40_1.0']
                            profile_110_BESS = DeratingDict['Dis-OE_40_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-OE_40_0.9']
                            profile_100_BESS = DeratingDict['Cha-OE_40_1.0']
                            profile_110_BESS = DeratingDict['Cha-OE_40_1.1']

                    else: #"50degC"
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-OE_50_0.9']
                            profile_100_BESS = DeratingDict['Dis-OE_50_1.0']
                            profile_110_BESS = DeratingDict['Dis-OE_50_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-OE_50_0.9']
                            profile_100_BESS = DeratingDict['Cha-OE_50_1.0']
                            profile_110_BESS = DeratingDict['Cha-OE_50_1.1']  
                            
                else: # when Qgen is negative (under excited)
                    if temperature == "35degC":
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-UE_35_0.9']
                            profile_100_BESS = DeratingDict['Dis-UE_35_1.0']
                            profile_110_BESS = DeratingDict['Dis-UE_35_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-UE_35_0.9']
                            profile_100_BESS = DeratingDict['Cha-UE_35_1.0']
                            profile_110_BESS = DeratingDict['Cha-UE_35_1.1']

                    elif temperature == "40degC":
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-UE_40_0.9']
                            profile_100_BESS = DeratingDict['Dis-UE_40_1.0']
                            profile_110_BESS = DeratingDict['Dis-UE_40_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-UE_40_0.9']
                            profile_100_BESS = DeratingDict['Cha-UE_40_1.0']
                            profile_110_BESS = DeratingDict['Cha-UE_40_1.1']

                    else: #"50degC"
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-UE_50_0.9']
                            profile_100_BESS = DeratingDict['Dis-UE_50_1.0']
                            profile_110_BESS = DeratingDict['Dis-UE_50_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-UE_50_0.9']
                            profile_100_BESS = DeratingDict['Cha-UE_50_1.0']
                            profile_110_BESS = DeratingDict['Cha-UE_50_1.1']  

                # Estimate the Qlim (individual Qlim) at differnt terminal voltage levels (0.9, 1.0 and 1.1pu) at a given P power level
                Qlim_90_BESS, Smax_90_BESS = iterpolate_derating(profile_90_BESS, P_individual_BESS)
                Qlim_100_BESS, Smax_100_BESS = iterpolate_derating(profile_100_BESS, P_individual_BESS)
                Qlim_110_BESS, Smax_110_BESS = iterpolate_derating(profile_110_BESS, P_individual_BESS)
    
                # Updating the Slim base on the P and Qlim. S suppose to derate linearly with voltage
                Slim_90_BESS = math.sqrt(P_individual_BESS**2 + Qlim_90_BESS**2)
                Slim_100_BESS = math.sqrt(P_individual_BESS**2 + Qlim_100_BESS**2)
                Slim_110_BESS = math.sqrt(P_individual_BESS**2 + Qlim_110_BESS**2)
    
                # Limit the S if it is over the maximum
                if Slim_90_BESS > Smax_90_BESS: Slim_90_BESS = Smax_90_BESS
                if Slim_100_BESS > Smax_100_BESS: Slim_100_BESS = Smax_100_BESS
                if Slim_110_BESS > Smax_110_BESS: Slim_110_BESS = Smax_110_BESS
                
                # Interporating function of the S on INV terminal voltage level
                S_interp_BESS = scipy.interpolate.interp1d([0.9, 1, 1.1], [Slim_90_BESS, Slim_100_BESS, Slim_110_BESS])
                
                # fetch the terminal voltage for estimating the derating
                ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')

                while (delta_Vinv > tolerance) and iters_out < 100: # if the voltage level at inveter varies more than the accepted band
                
                    # Update Sbase with new derating factor based on the terminal voltage level
                    Slim_BESS_ind = update_Slim(S_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
                    if Sbase_BESS > abs(P_BESS): Qlim_BESS = math.sqrt(Sbase_BESS**2 - P_BESS**2) 
                    else: Qlim_BESS = 0

                    # Apply the new base
                    psspy.machine_chng_2(busgen_V_BESS, '1', [_i,_i,_i,_i,_i,_i], [ P_BESS,_f, Qlim_BESS, -Qlim_BESS,_f,_f, Sbase_BESS,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0])    
                
                    # Iteratively converge on maximum operating point by setting Qlim, solving for inverter terminal voltage and re-calculating Qlim                
    #                iters = 0
                    ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                    ierr, Vinv_BESS_original = psspy.busdat(busgen_V_BESS, 'PU')
                    Slim_BESS_ind = update_Slim(S_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
                    if Sbase_BESS > abs(P_BESS): Qlim_BESS = math.sqrt(Sbase_BESS**2 - P_BESS**2) 
                    else: Qlim_BESS = 0
                    
                    ierr, cmpval = psspy.gendat(busgen_V_BESS)
                    Pinv_BESS = cmpval.real
                    Qinv_BESS = cmpval.imag
                    Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2)                    
                    Iinv_BESS = Sinv_BESS / Vinv_BESS
                
                    # If we run into the apparent power limit, start reducing Q, then reducing P.
                    iters = 0
                    print 'SINV_BESS = %f' % Sinv_BESS
                    while (Sinv_BESS > Sbase_BESS) and iters < 500:
                        Qlim_BESS -= Q_step
                        if Qlim_BESS < 0.0: 
                            Qlim_BESS = 0.0
                            if P_BESS > 0: P_BESS -= P_step
                            elif P_BESS < 0: P_BESS += P_step
#                            if P_BESS > Sbase_BESS: P_BESS = Sbase_BESS # Limit P to Sbase
#                            if P_BESS < -Sbase_BESS: P_BESS = -Sbase_BESS # Limit P to Sbase
                        print '*********************************************************'
                        print 'Apparent power limit reached %1.4f, QlimBESS reduced to %1.4f' % (Sinv_BESS, Qlim_BESS)
                        print '*********************************************************'
                        #
                        psspy.machine_chng_2(busgen_V_BESS, '1', [_i,_i,_i,_i,_i,_i], [ P_BESS,_f, Qlim_BESS, -Qlim_BESS,_f,_f, Sbase_BESS,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
                        #if soln_ierr:
                        #    break
                        ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                        ierr, cmpval = psspy.gendat(busgen_V_BESS)
                        Pinv_BESS = cmpval.real
                        Qinv_BESS = cmpval.imag
                        Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2)
                        Iinv_BESS = Sinv_BESS / Vinv_BESS
                        Slim_BESS_ind = update_Slim(S_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                        Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
                        iters += 1
                
                    time.sleep(0)
                    ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                    delta_Vinv_BESS = abs(Vinv_BESS_original - Vinv_BESS)
                    delta_Vinv = delta_Vinv_BESS
                    iters_out += 1
                    
            ###################################################################
            ###################################################################
            else:   # in case both PV and BESS are included
                if Vset_inv >1: # when Qgen is positive (over excited)
                    if temperature == "35degC":
                        profile_90_PV = DeratingDict['PV_OE_35_0.9']
                        profile_100_PV = DeratingDict['PV_OE_35_1.0']
                        profile_110_PV = DeratingDict['PV_OE_35_1.1']
                        
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-OE_35_0.9']
                            profile_100_BESS = DeratingDict['Dis-OE_35_1.0']
                            profile_110_BESS = DeratingDict['Dis-OE_35_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-OE_35_0.9']
                            profile_100_BESS = DeratingDict['Cha-OE_35_1.0']
                            profile_110_BESS = DeratingDict['Cha-OE_35_1.1']

                    elif temperature == "40degC":
                        profile_90_PV = DeratingDict['PV_OE_40_0.9']
                        profile_100_PV = DeratingDict['PV_OE_40_1.0']
                        profile_110_PV = DeratingDict['PV_OE_40_1.1']
                        
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-OE_40_0.9']
                            profile_100_BESS = DeratingDict['Dis-OE_40_1.0']
                            profile_110_BESS = DeratingDict['Dis-OE_40_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-OE_40_0.9']
                            profile_100_BESS = DeratingDict['Cha-OE_40_1.0']
                            profile_110_BESS = DeratingDict['Cha-OE_40_1.1']

                    else: #"50degC"
                        profile_90_PV = DeratingDict['PV_OE_50_0.9']
                        profile_100_PV = DeratingDict['PV_OE_50_1.0']
                        profile_110_PV = DeratingDict['PV_OE_50_1.1']
                        
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-OE_50_0.9']
                            profile_100_BESS = DeratingDict['Dis-OE_50_1.0']
                            profile_110_BESS = DeratingDict['Dis-OE_50_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-OE_50_0.9']
                            profile_100_BESS = DeratingDict['Cha-OE_50_1.0']
                            profile_110_BESS = DeratingDict['Cha-OE_50_1.1']  
                            
                else: # when Qgen is negative (under excited)
                    if temperature == "35degC":
                        profile_90_PV = DeratingDict['PV_UE_35_0.9']
                        profile_100_PV = DeratingDict['PV_UE_35_1.0']
                        profile_110_PV = DeratingDict['PV_UE_35_1.1']
                        
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-UE_35_0.9']
                            profile_100_BESS = DeratingDict['Dis-UE_35_1.0']
                            profile_110_BESS = DeratingDict['Dis-UE_35_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-UE_35_0.9']
                            profile_100_BESS = DeratingDict['Cha-UE_35_1.0']
                            profile_110_BESS = DeratingDict['Cha-UE_35_1.1']

                    elif temperature == "40degC":
                        profile_90_PV = DeratingDict['PV_UE_40_0.9']
                        profile_100_PV = DeratingDict['PV_UE_40_1.0']
                        profile_110_PV = DeratingDict['PV_UE_40_1.1']
                        
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-UE_40_0.9']
                            profile_100_BESS = DeratingDict['Dis-UE_40_1.0']
                            profile_110_BESS = DeratingDict['Dis-UE_40_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-UE_40_0.9']
                            profile_100_BESS = DeratingDict['Cha-UE_40_1.0']
                            profile_110_BESS = DeratingDict['Cha-UE_40_1.1']

                    else: #"50degC"
                        profile_90_PV = DeratingDict['PV_UE_50_0.9']
                        profile_100_PV = DeratingDict['PV_UE_50_1.0']
                        profile_110_PV = DeratingDict['PV_UE_50_1.1']
                        
                        if P_individual_BESS >=0: #when discharging
                            profile_90_BESS = DeratingDict['Dis-UE_50_0.9']
                            profile_100_BESS = DeratingDict['Dis-UE_50_1.0']
                            profile_110_BESS = DeratingDict['Dis-UE_50_1.1']
                        else: # When charging
                            profile_90_BESS = DeratingDict['Cha-UE_50_0.9']
                            profile_100_BESS = DeratingDict['Cha-UE_50_1.0']
                            profile_110_BESS = DeratingDict['Cha-UE_50_1.1']  

                # + At each active power level, determine the associated Qlim for differnt voltage level (from individual PQ curve) Qlim_0.9, Qlim_1.0 and Qlim_1.1
                # + From Qlim, calculate the Slim. Slim_0.9, Slim_1.0 and Slim_1.1. This is needed for derating interporation (S derate linearly with V)
                # + Check and make sure the Slim is capped by Smax
                # + 
                # Estimate the Qlim (individual Qlim) at differnt terminal voltage levels (0.9, 1.0 and 1.1pu) at a given P power level
                Qlim_90_PV, Smax_90_PV = iterpolate_derating(profile_90_PV, P_individual_PV)
                Qlim_100_PV, Smax_100_PV = iterpolate_derating(profile_100_PV, P_individual_PV)
                Qlim_110_PV, Smax_110_PV = iterpolate_derating(profile_110_PV, P_individual_PV)                        
                            
                Qlim_90_BESS, Smax_90_BESS = iterpolate_derating(profile_90_BESS, P_individual_BESS)
                Qlim_100_BESS, Smax_100_BESS = iterpolate_derating(profile_100_BESS, P_individual_BESS)
                Qlim_110_BESS, Smax_110_BESS = iterpolate_derating(profile_110_BESS, P_individual_BESS)
    
                # Updating the Slim base on the P and Qlim. S suppose to derate linearly with voltage
                Slim_90_PV = math.sqrt(P_individual_PV**2 + Qlim_90_PV**2)
                Slim_100_PV = math.sqrt(P_individual_PV**2 + Qlim_100_PV**2)
                Slim_110_PV = math.sqrt(P_individual_PV**2 + Qlim_110_PV**2)
                
                Slim_90_BESS = math.sqrt(P_individual_BESS**2 + Qlim_90_BESS**2)
                Slim_100_BESS = math.sqrt(P_individual_BESS**2 + Qlim_100_BESS**2)
                Slim_110_BESS = math.sqrt(P_individual_BESS**2 + Qlim_110_BESS**2)
    
                # Limit the S if it is over the maximum
                if Slim_90_PV > Smax_90_PV: Slim_90_PV = Smax_90_PV
                if Slim_100_PV > Smax_100_PV: Slim_100_PV = Smax_100_PV
                if Slim_110_PV > Smax_110_PV: Slim_110_PV = Smax_110_PV
    
                if Slim_90_BESS > Smax_90_BESS: Slim_90_BESS = Smax_90_BESS
                if Slim_100_BESS > Smax_100_BESS: Slim_100_BESS = Smax_100_BESS
                if Slim_110_BESS > Smax_110_BESS: Slim_110_BESS = Smax_110_BESS
                
                # Interporating function of the S on INV terminal voltage level
                S_interp_PV = scipy.interpolate.interp1d([0.9, 1, 1.1], [Slim_90_PV, Slim_100_PV, Slim_110_PV])            
                S_interp_BESS = scipy.interpolate.interp1d([0.9, 1, 1.1], [Slim_90_BESS, Slim_100_BESS, Slim_110_BESS])
                
                # fetch the terminal voltage for estimating the derating
                ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
            


                while (delta_Vinv > tolerance) and iters_out < 100: # if the voltage level at inveter varies more than the accepted band
                
                    # Update Sbase with new derating factor based on the terminal voltage level
                    Slim_BESS_ind = update_Slim(S_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
                    if Sbase_BESS > abs(P_BESS): Qlim_BESS = math.sqrt(Sbase_BESS**2 - P_BESS**2) 
                    else: Qlim_BESS = 0
    
                    Slim_PV_ind = update_Slim(S_interp_PV,Vinv_PV,Slim_110_PV,Slim_90_PV) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                    if Sbase_PV > abs(P_PV): Qlim_PV = math.sqrt(Sbase_PV**2 - P_PV**2) 
                    else: Qlim_PV = 0
    
                    # Apply the new base
                    psspy.machine_chng_2(busgen_V_PV, '1', [_i,_i,_i,_i,_i,_i], [ P_PV,_f, Qlim_PV, -Qlim_PV,_f,_f, Sbase_PV,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.machine_chng_2(busgen_V_BESS, '1', [_i,_i,_i,_i,_i,_i], [ P_BESS,_f, Qlim_BESS, -Qlim_BESS,_f,_f, Sbase_BESS,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0])    
                
                    # Iteratively converge on maximum operating point by setting Qlim, solving for inverter terminal voltage and re-calculating Qlim                
                    ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                    ierr, Vinv_BESS_original = psspy.busdat(busgen_V_BESS, 'PU')
                    Slim_BESS_ind = update_Slim(S_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
                    if Sbase_BESS > abs(P_BESS): Qlim_BESS = math.sqrt(Sbase_BESS**2 - P_BESS**2) 
                    else: Qlim_BESS = 0
                    
                    ierr, cmpval = psspy.gendat(busgen_V_BESS)
                    Pinv_BESS = cmpval.real
                    Qinv_BESS = cmpval.imag
                    Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2)                    
                    Iinv_BESS = Sinv_BESS / Vinv_BESS
                    
                    ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                    ierr, Vinv_PV_original = psspy.busdat(busgen_V_PV, 'PU')
                    Slim_PV_ind = update_Slim(S_interp_PV,Vinv_PV,Slim_110_PV,Slim_90_PV) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                    Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                    if Sbase_PV > abs(P_PV): Qlim_PV = math.sqrt(Sbase_PV**2 - P_PV**2) 
                    else: Qlim_PV = 0
    
                    ierr, cmpval = psspy.gendat(busgen_V_PV)
                    Pinv_PV = cmpval.real
                    Qinv_PV = cmpval.imag
                    Sinv_PV = math.sqrt(P_PV**2 + Qinv_PV**2)                    
                    Iinv_PV = Sinv_PV / Vinv_PV
           
                
                    # If we run into the apparent power limit, start reducing Q, then reducing P.
                    iters = 0
                    print 'SINV_BESS = %f' % Sinv_BESS
                    print 'SINV_PV = %f' % Sinv_PV
                    while ((Sinv_BESS > Sbase_BESS) or (Sinv_PV > Sbase_PV)) and iters < 500:
                        if (Sinv_BESS > Sbase_BESS): # If BESS overloaded
                            Qlim_BESS -= Q_step
                            if Qlim_BESS < 0.0: 
                                Qlim_BESS = 0.0
                                if P_BESS > 0: P_BESS -= P_step
                                elif P_BESS < 0: P_BESS += P_step
    #                            if P_BESS > Sbase_BESS: P_BESS = Sbase_BESS # Limit P to Sbase
    #                            if P_BESS < -Sbase_BESS: P_BESS = -Sbase_BESS # Limit P to Sbase
                            print '*********************************************************'
                            print 'Apparent power limit reached %1.4f, QlimBESS reduced to %1.4f' % (Sinv_BESS, Qlim_BESS)
                            print '*********************************************************'
                            #
                            psspy.machine_chng_2(busgen_V_BESS, '1', [_i,_i,_i,_i,_i,_i], [ P_BESS,_f, Qlim_BESS, -Qlim_BESS,_f,_f, Sbase_BESS,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
                            #if soln_ierr:
                            #    break
                            ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                            ierr, cmpval = psspy.gendat(busgen_V_BESS)
                            Pinv_BESS = cmpval.real
                            Qinv_BESS = cmpval.imag
                            Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2)
                            Iinv_BESS = Sinv_BESS / Vinv_BESS
                            Slim_BESS_ind = update_Slim(S_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                            Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
    
                        if (Sinv_PV > Sbase_PV): # If PV overloaded
                            Qlim_PV -= Q_step
                            if Qlim_PV < 0.0: 
                                Qlim_PV = 0.0
                                if P_PV > 0: P_PV -= P_step
    #                            elif P_PV < 0: P_PV += P_step
                                if P_PV < 0: P_PV = 0 # PV do not absorb P
                            print '*********************************************************'
                            print 'Apparent power limit reached %1.4f, QlimPV reduced to %1.4f' % (Sinv_PV, Qlim_PV)
                            print '*********************************************************'
                            #
                            psspy.machine_chng_2(busgen_V_PV, '1', [_i,_i,_i,_i,_i,_i], [ P_PV,_f, Qlim_PV, -Qlim_PV,_f,_f, Sbase_PV,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                            soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
                            #if soln_ierr:
                            #    break
                            ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                            ierr, cmpval = psspy.gendat(busgen_V_PV)
                            Pinv_PV = cmpval.real
                            Qinv_PV = cmpval.imag
                            Sinv_PV = math.sqrt(Pinv_PV**2 + Qinv_PV**2)
                            Iinv_PV = Sinv_PV / Vinv_PV
                            Slim_PV_ind = update_Slim(S_interp_PV,Vinv_PV,Slim_110_PV,Slim_90_PV) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
                            Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                        iters += 1
                
    
                    time.sleep(0)
                    ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                    delta_Vinv_BESS = abs(Vinv_BESS_original - Vinv_BESS)
                    ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                    delta_Vinv_PV = abs(Vinv_PV_original - Vinv_PV)                
                    delta_Vinv = max(delta_Vinv_PV,delta_Vinv_BESS)
                    iters_out += 1




            # Prepare outputs
            ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
            ierr, cmpval = psspy.gendat(busgen_V_BESS)
            try:
                Pinv_BESS = cmpval.real
                Qinv_BESS = cmpval.imag
                Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2)
                Iinv_BESS = Sinv_BESS / Vinv_BESS
            except:
                Pinv_BESS = 0
                Qinv_BESS = 0
                Sinv_BESS = 0
                Iinv_BESS = 0

            ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
            ierr, cmpval = psspy.gendat(busgen_V_PV)
            try:
                Pinv_PV = cmpval.real
                Qinv_PV = cmpval.imag
                Sinv_PV = math.sqrt(Pinv_PV**2 + Qinv_PV**2)
                Iinv_PV = Sinv_PV / Vinv_PV
            except:
                Pinv_PV = 0
                Qinv_PV = 0
                Sinv_PV = 0
                Iinv_PV = 0
                
#            ierr, cmpval = psspy.gendat(bus_inf)
            ierr, cmpval = psspy.brnflo(bus_POC_fr, bus_POC_to, r"""1""")
            Pcp = cmpval.real
            Qcp = cmpval.imag
            #
            print '*****************************'
            print ' Vinv, Pinv, Qinv, Sinv, Iinv (pu)'
#            print ' %1.4f, %1.2f, %1.4f, %1.4f, %1.4f' % (Vinv_BESS, P, Qinv_BESS, Sinv_BESS, Iinv_BESS / Sbase_BESS)
            print '*****************************'
            #
#            Vbus[idx].append(Vinv)
#            Pbus[idx].append(Pinv)
#            Qbus[idx].append(Qinv)
#            Sbus[idx].append(Sinv)
#            Ibus[idx].append(Iinv)
#            Pconn[idx].append(Pcp)
#            Qconn[idx].append(Qcp)


            # Monitoring parameters
#            P_poc, Q_poc, V_poc, P_inv_gen1, Q_inv_gen1, V_inv_gen1, P_inv_gen2, Q_inv_gen2, V_inv_gen2 = monitor_SF(bus_POC_fr, bus_POC_to, bus_POC_fr, busgen1, busgen2)
            P_inv_gen1 = Pinv_PV
            Q_inv_gen1 = Qinv_PV
            V_inv_gen1 = Vinv_PV
            P_inv_gen2 = Pinv_BESS
            Q_inv_gen2 = Qinv_BESS
            V_inv_gen2 = Vinv_BESS
            P_poc = Pcp
            Q_poc = Qcp
            V_poc = Vcp
            Sinv_gen1 = Sinv_PV
            Iinv_gen1 = Iinv_PV
            Sinv_gen2 = Sinv_BESS
            Iinv_gen2 = Iinv_BESS            

            if Vset_inv > 1: # Q positive
                P_inv_gen1_all_pos.append(P_inv_gen1)
                Q_inv_gen1_all_pos.append(Q_inv_gen1)
                V_inv_gen1_all_pos.append(V_inv_gen1)
                P_inv_gen2_all_pos.append(P_inv_gen2)
                Q_inv_gen2_all_pos.append(Q_inv_gen2)
                V_inv_gen2_all_pos.append(V_inv_gen2)
                P_poc_all_pos.append(P_poc)
                Q_poc_all_pos.append(Q_poc)
                V_poc_all_pos.append(V_poc)
                Sinv_gen1_all_pos.append(Sinv_gen1)
                Iinv_gen1_all_pos.append(Iinv_gen1)
                Sinv_gen2_all_pos.append(Sinv_gen2)
                Iinv_gen2_all_pos.append(Iinv_gen2)
            else: # Q negative
                P_inv_gen1_all_neg.append(P_inv_gen1)
                Q_inv_gen1_all_neg.append(Q_inv_gen1)
                V_inv_gen1_all_neg.append(V_inv_gen1)
                P_inv_gen2_all_neg.append(P_inv_gen2)
                Q_inv_gen2_all_neg.append(Q_inv_gen2)
                V_inv_gen2_all_neg.append(V_inv_gen2)
                P_poc_all_neg.append(P_poc)
                Q_poc_all_neg.append(Q_poc)
                V_poc_all_neg.append(V_poc)
                Sinv_gen1_all_neg.append(Sinv_gen1)
                Iinv_gen1_all_neg.append(Iinv_gen1)
                Sinv_gen2_all_neg.append(Sinv_gen2)
                Iinv_gen2_all_neg.append(Iinv_gen2)


    P_inv_gen1_all = P_inv_gen1_all_pos + P_inv_gen1_all_neg
    Q_inv_gen1_all = Q_inv_gen1_all_pos + Q_inv_gen1_all_neg
    V_inv_gen1_all = V_inv_gen1_all_pos + V_inv_gen1_all_neg
    P_inv_gen2_all = P_inv_gen2_all_pos + P_inv_gen2_all_neg
    Q_inv_gen2_all = Q_inv_gen2_all_pos + Q_inv_gen2_all_neg
    V_inv_gen2_all = V_inv_gen2_all_pos + V_inv_gen2_all_neg
    P_poc_all = P_poc_all_pos + P_poc_all_neg
    Q_poc_all = Q_poc_all_pos + Q_poc_all_neg
    V_poc_all = V_poc_all_pos + V_poc_all_neg
    Sinv_gen1_all = Sinv_gen1_all_pos + Sinv_gen1_all_neg
    Iinv_gen1_all = Iinv_gen1_all_pos + Iinv_gen1_all_neg
    Sinv_gen2_all = Sinv_gen2_all_pos + Sinv_gen2_all_neg
    Iinv_gen2_all = Iinv_gen2_all_pos + Iinv_gen2_all_neg

    
#    # Open csv file and write the title for each case
#    csv_file_name = 'S5251_%s_%s.csv' % (temp_name, vol_case_name)
##    csv_file_name = 'S5251_result.csv'
#    f = open(os.path.join(output_loc, csv_file_name), 'w')  # 'w' write mode
#    # f.write('P_Inv_1101, Q_Inv_1101, V_Inv_1101,P_Inv_1102, Q_Inv_1102, V_Inv_1102,P_Poc,Q_Poc,V_Poc, S_limit\n')
#    f.write('V_Inv, P_Inv, Q_Inv, S_Inv, I_Inv, P_Poc, Q_Poc\n')
#
#    for idx in range(len(P_inv_gen1_all) / 2):
#        # Export the data to csv file        
#        # f.write('%1.3f, %1.3f, %1.5f,%1.3f, %1.3f, %1.5f, %1.3f, %1.3f, %1.5f,%1.3f\n' % (P_inv_gen1_all[idx], Q_inv_gen1_all[idx], V_inv_gen1_all[idx], P_inv_gen2_all[idx], Q_inv_gen2_all[idx], V_inv_gen2_all[idx],P_poc_all[idx], Q_poc_all[idx], V_poc_all[idx], Sinv_all[idx]))
#        f.write('%1.4f, %1.4f, %1.4f, %1.4f, %1.6f, %1.4f, %1.4f\n' % (V_inv_gen1_all[idx], P_inv_gen1_all[idx], Q_inv_gen1_all[idx], Sinv_all[idx], Iinv_all[idx], P_poc_all[idx], Q_poc_all[idx]))
#
#        # Update the dataset to be plotted
#        P_POC_plot.append(P_poc_all[idx])
#        Q_POC_plot.append(Q_poc_all[idx])
#
#    for idx in range(len(P_inv_gen1_all) - 1, (len(P_inv_gen1_all) / 2) - 1, -1):
#        # Export the data to csv file        
#        # f.write('%1.3f, %1.3f, %1.5f,%1.3f, %1.3f, %1.5f, %1.3f, %1.3f, %1.5f,%1.3f\n' % (P_inv_gen1_all[idx], Q_inv_gen1_all[idx], V_inv_gen1_all[idx], P_inv_gen2_all[idx], Q_inv_gen2_all[idx], V_inv_gen2_all[idx],P_poc_all[idx], Q_poc_all[idx], V_poc_all[idx], Sinv_all[idx]))
#        f.write('%1.4f, %1.4f, %1.4f, %1.4f, %1.6f, %1.4f, %1.4f\n' % (V_inv_gen1_all[idx], P_inv_gen1_all[idx], Q_inv_gen1_all[idx], Sinv_all[idx], Iinv_all[idx], P_poc_all[idx], Q_poc_all[idx]))
#        # Update the dataset to be plotted
#        P_POC_plot.append(P_poc_all[idx])
#        Q_POC_plot.append(Q_poc_all[idx])
#    f.close()
##    P_POC_plot_final.append(P_POC_plot)
##    Q_POC_plot_final.append(Q_POC_plot)



    #Writing summary results into one file
    
    
#    data_out = {}
#    data_out['1 V_Inv'] = []
#    data_out['2 P_Inv'] = []
#    data_out['3 Q_Inv'] = []
#    data_out['4 S_Inv'] = []
#    data_out['5 I_Inv'] = []
#    data_out['6 P_Poc'] = []
#    data_out['7 Q_Poc'] = []    
#    for idx in range(len(P_inv_gen1_all) / 4):
#        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
#        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
#        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
#        data_out['4 S_Inv'].append(Sinv_all[idx])
#        data_out['5 I_Inv'].append(Iinv_all[idx])
#        data_out['6 P_Poc'].append(P_poc_all[idx])
#        data_out['7 Q_Poc'].append(Q_poc_all[idx]) 
#    for idx in range((3*len(P_inv_gen1_all) / 4) - 1, len(P_inv_gen1_all) / 2, -1):
#        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
#        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
#        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
#        data_out['4 S_Inv'].append(Sinv_all[idx])
#        data_out['5 I_Inv'].append(Iinv_all[idx])
#        data_out['6 P_Poc'].append(P_poc_all[idx])
#        data_out['7 Q_Poc'].append(Q_poc_all[idx])
#    for idx in range((len(P_inv_gen1_all) / 4), len(P_inv_gen1_all) / 2 - 1, 1):
#        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
#        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
#        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
#        data_out['4 S_Inv'].append(Sinv_all[idx])
#        data_out['5 I_Inv'].append(Iinv_all[idx])
#        data_out['6 P_Poc'].append(P_poc_all[idx])
#        data_out['7 Q_Poc'].append(Q_poc_all[idx]) 
#    for idx in range(len(P_inv_gen1_all) - 1, (3*len(P_inv_gen1_all) / 4) - 1, -1):
#        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
#        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
#        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
#        data_out['4 S_Inv'].append(Sinv_all[idx])
#        data_out['5 I_Inv'].append(Iinv_all[idx])
#        data_out['6 P_Poc'].append(P_poc_all[idx])
#        data_out['7 Q_Poc'].append(Q_poc_all[idx]) 


    data_out = {}
    data_out['1 V_Inv'] = []
    data_out['2 P_Inv'] = []
    data_out['3 Q_Inv'] = []
    data_out['4 S_Inv'] = []
    data_out['5 I_Inv'] = []
    data_out['1 V_Inv2'] = []
    data_out['2 P_Inv2'] = []
    data_out['3 Q_Inv2'] = []
    data_out['4 S_Inv2'] = []
    data_out['5 I_Inv2'] = []
    data_out['6 P_Poc'] = []
    data_out['7 Q_Poc'] = []    
    
    for idx in range(len(P_inv_gen1_all) / 2):
        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
        data_out['4 S_Inv'].append(Sinv_gen1_all[idx])
        data_out['5 I_Inv'].append(Iinv_gen1_all[idx])
        data_out['1 V_Inv2'].append(V_inv_gen2_all[idx])
        data_out['2 P_Inv2'].append(P_inv_gen2_all[idx])
        data_out['3 Q_Inv2'].append(Q_inv_gen2_all[idx])
        data_out['4 S_Inv2'].append(Sinv_gen2_all[idx])
        data_out['5 I_Inv2'].append(Iinv_gen2_all[idx])
        data_out['6 P_Poc'].append(P_poc_all[idx])
        data_out['7 Q_Poc'].append(Q_poc_all[idx]) 

    for idx in range(len(P_inv_gen1_all) - 1, (len(P_inv_gen1_all) / 2) - 1, -1):
        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
        data_out['4 S_Inv'].append(Sinv_gen1_all[idx])
        data_out['5 I_Inv'].append(Iinv_gen1_all[idx])
        data_out['1 V_Inv2'].append(V_inv_gen2_all[idx])
        data_out['2 P_Inv2'].append(P_inv_gen2_all[idx])
        data_out['3 Q_Inv2'].append(Q_inv_gen2_all[idx])
        data_out['4 S_Inv2'].append(Sinv_gen2_all[idx])
        data_out['5 I_Inv2'].append(Iinv_gen2_all[idx])
        data_out['6 P_Poc'].append(P_poc_all[idx])
        data_out['7 Q_Poc'].append(Q_poc_all[idx]) 


        
#    data_out = {}
#    data_out['1 V_Inv'] = V_inv_gen1_all
#    data_out['2 P_Inv'] = P_inv_gen1_all
#    data_out['3 Q_Inv'] = Q_inv_gen1_all
#    data_out['4 S_Inv'] = Sinv_all
#    data_out['5 I_Inv'] = Iinv_all
#    data_out['6 P_Poc'] = P_poc_all
#    data_out['7 Q_Poc'] = Q_poc_all
    
    df_out = pd.DataFrame.from_dict(data = data_out)
    df_out.to_excel(writer, sheet_name = str(temp_name +"_"+ vol_case_name))

#    data_out_all.append(data_out)
    
writer.close() 


#writer = pd.ExcelWriter("S5251_all_" +str(temp_name)+".xlsx", engine = "xlsxwriter") # Preparing for exporting the result
#df_out = pd.DataFrame.from_dict(data = data_out_all)
#df_out.to_excel(writer, sheet_name = str(temp_name +"_all"))
#writer.close()

            
#psspy.lines_per_page_one_device(2,10000000)
#psspy.report_output(1,"",[0,0])
#psspy.lines_per_page_one_device(2,10000000)
#psspy.progress_output(1,"",[0,0])
#psspy.lines_per_page_one_device(2,10000000)
#psspy.prompt_output(1,"",[0,0])
#psspy.lines_per_page_one_device(2,10000000)
#psspy.alert_output(1,"",[0,0])
#
#
#filename = 'Registered Q Capability1'+'.csv'#This needs to be adjusted to work for mutil-report setup
#csvfile = output_loc + '\\' + filename
#f = open(csvfile, 'w')
##    f.write('Filename,Vsettle,Qsettle,Qrise,Vrise\n')
#f.write('Registered Q Capability\n')
#
#print 'Vcp = 90%,,,Vcp = 100%,,,Vcp = 110%'
#print 'Vinv90,Pinv90,Qinv90,Sinv90,Pcp90,Qcp90,Vinv100,Pinv100,Qinv100,Sinv100,Pcp100,Qcp100,Vinv110,Pinv110,Qinv110,Sinv110,Pcp110,Qcp110'
#
#f.write('Vcp = 90%,,,Vcp = 100%,,,Vcp = 110%\n')
#f.write('Vinv90,Pinv90,Qinv90,Sinv90,Pcp90,Qcp90,Vinv100,Pinv100,Qinv100,Sinv100,Pcp100,Qcp100,Vinv110,Pinv110,Qinv110,Sinv110,Pcp110,Qcp110\n')
#
#
#for idx in range(len(Vbus[0]) / 2):
#    line = ''
#    for n in range(len(Vinfinite)):
#        line += '%1.4f, %1.4f, %1.4f, %1.4f, %1.4f, %1.4f, ' % (Vbus[n][idx], Pbus[n][idx], Qbus[n][idx], Sbus[n][idx], -Pconn[n][idx], -Qconn[n][idx])
##        f.write('{:1.4f}, {:1.4f}, {:1.4f}, {:1.4f}, {:1.4f}, {:1.4f}\n'.format(Vbus[n][idx], Pbus[n][idx], Qbus[n][idx], Sbus[n][idx], -Pconn[n][idx], -Qconn[n][idx]))
#        f.write(line)
##    print line[:-2]
#
#for idx in range(len(Vbus[0]) - 1, (len(Vbus[0]) / 2) - 1, -1):
#    line = ''
#    for n in range(len(Vinfinite)):
#        line += '%1.4f, %1.4f, %1.4f, %1.4f, %1.4f, %1.4f, ' % (Vbus[n][idx], Pbus[n][idx], Qbus[n][idx], Sbus[n][idx], -Pconn[n][idx], -Qconn[n][idx])
##        f.write('{:1.4f}, {:1.4f}, {:1.4f}, {:1.4f}, {:1.4f}, {:1.4f}\n'.format(Vbus[n][idx], Pbus[n][idx], Qbus[n][idx], Sbus[n][idx], -Pconn[n][idx], -Qconn[n][idx]))
##    print line[:-2]
#
#'''
#for idx in range((len(Vbus[0]) / 2)):
#for n in range(3):
#    print 'Vinf, Vinv, P, Q, S, I, Pcp, Qcp'    
#    for idx in range((len(Vbus[0]) / 2)):
#        print '%1.4f, %1.4f, %1.2f, %1.4f, %1.4f, %1.4f, %1.4f, %1.4f' % (Vinfinite[n], Vbus[n][idx], Pbus[n][idx], Qbus[n][idx], Sbus[n][idx], Ibus[n][idx] / Sbase, -Pconn[n][idx], -Qconn[n][idx])
#
#    for idx in range(len(Vbus[0]) - 1, (len(Vbus[0]) / 2) - 1, -1):
#        print '%1.4f, %1.4f, %1.2f, %1.4f, %1.4f, %1.4f, %1.4f, %1.4f' % (Vinfinite[n], Vbus[n][idx], Pbus[n][idx], Qbus[n][idx], Sbus[n][idx], Ibus[n][idx] / Sbase, -Pconn[n][idx], -Qconn[n][idx])
#    print ' '
#'''