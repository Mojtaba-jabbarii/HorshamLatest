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
    + Check and adapt the mode: PV/BESS or 30degC/50degC
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
source_type = "BESS" #"BESS" # "PV"
temperature = "30degC" #"30degC" # "50degC"


temp_name = temperature + "_" + source_type #30degC_BESS"

PV_INV_num = 26.0
PV_INV_rate = 4.2 #MVA

BESS_INV_num = 36.0
BESS_INV_rate = 3.6
    

tapstep = 1 
inf_bus = 9900
Dum_bus = 9910
POC_bus = 9920

busgen1 = 9942
busTx1 = 9941
busgen2 = 9944
busTx2 = 9943

busMV = 9940
busHV = 9930


if source_type == "BESS":
    Sbase_analysis = BESS_INV_num * BESS_INV_rate
    INV_num_analysis = BESS_INV_num
    Qmax = 0.7*Sbase_analysis #Qmax is limited to 70% of the Sbase when voltage is equal or greater than 1pu. It is 0.9*70% = 63% when voltage is 0.9pu
    bus_dis1 = busgen1
    bus_dis2 = busTx1
    busgen_V = busgen2
    Pmax = 90   
    Pmin = -90
    
else:
    Sbase_analysis = PV_INV_num * PV_INV_rate
    INV_num_analysis = PV_INV_num
    Qmax = 0.6*Sbase_analysis
    bus_dis1 = busgen2
    bus_dis2 = busTx2
    busgen_V = busgen1
    Pmax = 90 
    Pmin = 0
    

import readtestinfo as readtestinfo

TestDefinitionSheet = r'20230104_SUM_TESTINFO_V0.xlsx'
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

busPOC = POC_bus
fbus = POC_bus
tbus = Dum_bus

#psspy.case(os.path.join(base_dir, case_name))
psspy.case(base_dir + case_name)
psspy.dscn(bus_dis1)
psspy.dscn(bus_dis2)

psspy.fnsl([1,0,0,1,1,0,99,0])
psspy.fnsl([1,0,0,1,1,0,99,0])
psspy.fnsl([1,0,0,1,1,0,99,0])

    
#
Vinfinite = [0.90, 1.0, 1.10]
#Vbus = [[] for _ in range(len(Vinfinite))]
#Pbus = [[] for _ in range(len(Vinfinite))]
#Qbus = [[] for _ in range(len(Vinfinite))]
#Pconn = [[] for _ in range(len(Vinfinite))]
#Qconn = [[] for _ in range(len(Vinfinite))]
#Sbus = [[] for _ in range(len(Vinfinite))]
#Ibus = [[] for _ in range(len(Vinfinite))]

psspy.lines_per_page_one_device(1,60)
psspy.report_output(2,'test.log',[0,0])
psspy.lines_per_page_one_device(1,60)
psspy.progress_output(2,'test.log',[0,0])
psspy.lines_per_page_one_device(1,60)
psspy.alert_output(2,'test.log',[0,0])
psspy.lines_per_page_one_device(1,60)
psspy.prompt_output(2,'test.log',[0,0])    


def monitor_SF(fbus, tbus, busPOC, busgen1, busgen2):

    ierr, S_inv_gen1 = psspy.gendat(busgen1)
    P_inv_gen1 = S_inv_gen1.real
    Q_inv_gen1 = S_inv_gen1.imag
    P_inv_gen2 = 0
    Q_inv_gen2 = 0
    ierr, cmpval = psspy.brnflo(fbus, tbus, r"""1""")
    P_poc = cmpval.real
    Q_poc = cmpval.imag
    ierr, V_poc = psspy.busdat(busPOC, 'PU')
    ierr, V_inv_gen1 = psspy.busdat(busgen1, 'PU')
    ierr, V_inv_gen2 = psspy.busdat(busgen1, 'PU')

    return P_poc, Q_poc, V_poc, P_inv_gen1, Q_inv_gen1, V_inv_gen1, P_inv_gen2, Q_inv_gen2, V_inv_gen2


def iterpolate_derating(derating_profile, P_individual):
    y_interp = scipy.interpolate.interp1d(derating_profile['x_data'], derating_profile['y_data'])
#    derate_out = 1
    try: derate_out = y_interp(P_individual)
    except: 
        if P_individual == 0: derate_out = derating_profile['y_data'][1]
        else: derate_out = derating_profile['y_data'][-1]
        
    return derate_out


P_POC_plot_final = []
Q_POC_plot_final = []

derating = 1 # if no lookup table is avaiable then derating is 1
Sbase = Sbase_analysis * derating  
if Pmin >= 0: #PV
    Ploop_range = range(0, int(Pmax) + 1, 2) + [Pmax] + range(int(Pmax) + 1, int(Sbase) + 1, 2) + [Sbase]
else: #BESS
    Ploop_range =  [-Sbase] + range(int(-Sbase) - 1, int(Pmin) - 1, 2) + [Pmin] + range(int(Pmin) - 1, int(Pmax) + 1, 2) + [Pmax] + range(int(Pmax) + 1, int(Sbase) + 1, 2) + [Sbase]

writer = pd.ExcelWriter("S5251_" +str(temp_name)+".xlsx", engine = "xlsxwriter") # Preparing for exporting the result

data_out_all = []

for index, Vcp in enumerate(Vinfinite):
    vol_case_name = str(Vcp)
    
    psspy.plant_data(inf_bus,0,[ Vcp,_f])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])    
#    psspy.fnsl([tapstep,0,0,1,1,0,0,0])

    # Force the voltage at POC by changing grid impedance to zero, solve with tap enable 
#        psspy.branch_chng_3(POC_bus,Dum_bus,r"""1""",[1,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s)
#            psspy.two_winding_chng_5(Dum_bus,POC_bus,r"""1""",[0,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[ 80.0, 80.0, 80.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s,_s)
    psspy.branch_chng_3(inf_bus,Dum_bus,r"""1""",[_i,_i,_i,_i,_i,_i],[0.0,0.00001,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
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
    Sinv_all = []
    Iinv_all = []
    
    P_inv_gen1_all_pos = []
    Q_inv_gen1_all_pos = []
    V_inv_gen1_all_pos = []
    P_inv_gen2_all_pos = []
    Q_inv_gen2_all_pos = []
    V_inv_gen2_all_pos = []
    P_poc_all_pos = []
    Q_poc_all_pos = []
    V_poc_all_pos = []
    Sinv_all_pos = []
    Iinv_all_pos = []

    P_inv_gen1_all_neg = []
    Q_inv_gen1_all_neg = []
    V_inv_gen1_all_neg = []
    P_inv_gen2_all_neg = []
    Q_inv_gen2_all_neg = []
    V_inv_gen2_all_neg = []
    P_poc_all_neg = []
    Q_poc_all_neg = []
    V_poc_all_neg = []
    Sinv_all_neg = []
    Iinv_all_neg = []
    
    for P in Ploop_range:
        P_individual = P/INV_num_analysis * 1000 #kW

        # Determine reactive limit based on power output and Sbase
        if Sbase > abs(P): Qlim = math.sqrt(Sbase**2 - P**2)
        else: Qlim = 0.0
        if Qlim > Qmax: Qlim = Qmax
#            Qlim = math.sqrt(Sbase**2 - P**2)
        print 'Qlim = %1.4f' % Qlim
        # Aplly P,Q limit
        psspy.machine_chng_2(busgen_V, '1', [_i,_i,_i,_i,_i,_i], [ P,_f, Qlim, -Qlim,_f,_f, Sbase,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#        psspy.fnsl([1,0,0,1,1,0,0,0])
#        psspy.fnsl([1,0,0,1,1,0,0,0])

        psspy.fnsl([tapstep,0,0,1,1,0,0,0])
        psspy.fnsl([tapstep,0,0,1,1,0,0,0])
            
        # fetch the terminal voltage for estimating the derating
        ierr, Vinv = psspy.busdat(busgen_V, 'PU')

        for Vset_inv in [1.2, 0.7]: # change voltage setpoint of generator to get a Q range. For each case, solve the case with the tap locked
            psspy.plant_data(busgen_V,0,[ Vset_inv,_f])
            soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
            psspy.fnsl([tapstep,0,0,1,1,0,0,0])
            psspy.fnsl([tapstep,0,0,1,1,0,0,0])

            # Update derating factor for 3 terminal voltage levels (0.9, 1.0 and 1.1pu) based on PQ quadrant
            if source_type == "BESS":   
                if temperature == "30degC":
                    if Vset_inv >1: # when Qgen is positive (over excited)
                        if P_individual >=0: #when discharging
                            derate_90 = iterpolate_derating(DeratingDict['Dis-OE_35_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Dis-OE_35_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Dis-OE_35_1.1'], P_individual)
         
                        else: # When charging
                            derate_90 = iterpolate_derating(DeratingDict['Cha-OE_35_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Cha-OE_35_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Cha-OE_35_1.1'], P_individual)
        
                    elif Vset_inv <1: # when Qgen is negative (under excited)
                        if P_individual >=0: #when discharging
                            derate_90 = iterpolate_derating(DeratingDict['Dis-UE_35_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Dis-UE_35_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Dis-UE_35_1.1'], P_individual)
         
                        else: # When charging
                            derate_90 = iterpolate_derating(DeratingDict['Cha-UE_35_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Cha-UE_35_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Cha-UE_35_1.1'], P_individual)

                else:

                    if Vset_inv >1: # when Qgen is positive (over excited)
                        if P_individual >=0: #when discharging
                            derate_90 = iterpolate_derating(DeratingDict['Dis-OE_50_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Dis-OE_50_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Dis-OE_50_1.1'], P_individual)
         
                        else: # When charging
                            derate_90 = iterpolate_derating(DeratingDict['Cha-OE_50_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Cha-OE_50_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Cha-OE_50_1.1'], P_individual)
        
                    elif Vset_inv <1: # when Qgen is negative (under excited)
                        if P_individual >=0: #when discharging
                            derate_90 = iterpolate_derating(DeratingDict['Dis-UE_50_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Dis-UE_50_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Dis-UE_50_1.1'], P_individual)
         
                        else: # When charging
                            derate_90 = iterpolate_derating(DeratingDict['Cha-UE_50_0.9'], P_individual)
                            derate_100 = iterpolate_derating(DeratingDict['Cha-UE_50_1.0'], P_individual)
                            derate_110 = iterpolate_derating(DeratingDict['Cha-UE_50_1.1'], P_individual)
                            
            else:
                if temperature == "30degC":
                    if Vset_inv >1: # when Qgen is positive (over excited)
                        derate_90 = iterpolate_derating(DeratingDict['PV_OE_35_0.9'], P_individual)
                        derate_100 = iterpolate_derating(DeratingDict['PV_OE_35_1.0'], P_individual)
                        derate_110 = iterpolate_derating(DeratingDict['PV_OE_35_1.1'], P_individual)
         
                    elif Vset_inv <1: # when Qgen is negative (under excited)
                        derate_90 = iterpolate_derating(DeratingDict['PV_UE_35_0.9'], P_individual)
                        derate_100 = iterpolate_derating(DeratingDict['PV_UE_35_1.0'], P_individual)
                        derate_110 = iterpolate_derating(DeratingDict['PV_UE_35_1.1'], P_individual)

                else:
                    if Vset_inv >1: # when Qgen is positive (over excited)
                        derate_90 = iterpolate_derating(DeratingDict['PV_OE_50_0.9'], P_individual)
                        derate_100 = iterpolate_derating(DeratingDict['PV_OE_50_1.0'], P_individual)
                        derate_110 = iterpolate_derating(DeratingDict['PV_OE_50_1.1'], P_individual)
         
                    elif Vset_inv <1: # when Qgen is negative (under excited)
                        derate_90 = iterpolate_derating(DeratingDict['PV_UE_50_0.9'], P_individual)
                        derate_100 = iterpolate_derating(DeratingDict['PV_UE_50_1.0'], P_individual)
                        derate_110 = iterpolate_derating(DeratingDict['PV_UE_50_1.1'], P_individual)

            y_interp = scipy.interpolate.interp1d([0.9, 1, 1.1], [derate_90, derate_100, derate_110])            
#            try: y_interp = scipy.interpolate.interp1d([0.9, 1, 1.1], [derate_90, derate_100, derate_110])
#            except: y_interp = scipy.interpolate.interp1d([0.9, 1, 1.1], [0.9, 1.0, 1.0])
            
            # Update Sbase with new derating factor based on the terminal voltage
            delta_Vinv = 1
            tolerance = 0.001
            iters_out = 0
            P_step = 0.5
            Q_step = 0.2            
            while (delta_Vinv > tolerance) and iters_out < 100: # if the voltage level at inveter varies more than the accepted band
            
                # Update Sbase with new derating factor based on the terminal voltage level
                try:
                    derating = y_interp(Vinv)
                except: # if the voltage is outside the range [0.9-1.1pu]
                    if Vinv > 1.1: derating = derate_110
                    elif Vinv < 0.9: derating = derate_90
                Sbase = Sbase_analysis * derating
                Qlim_lowV = Qlim
                if Vinv <1.0: Qlim_lowV = 0.7*Sbase*Vinv #Qmax is limited to 70% of the Sbase when voltage is equal or greater than 1pu. It is 0.9*70% = 63% when voltage is 0.9pu
                if Vinv <0.9: Qlim_lowV = 0.7*Sbase*0.9
                if Qlim > Qlim_lowV: Qlim = Qlim_lowV
                
                # Apply the new base
                psspy.machine_chng_2(busgen_V, '1', [_i,_i,_i,_i,_i,_i], [ P,_f, Qlim, -Qlim,_f,_f, Sbase,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                psspy.fnsl([tapstep,0,0,1,1,0,0,0])                   
                
                # Iteratively converge on maximum operating point by setting Qlim, solving for inverter terminal voltage and re-calculating Qlim                
                iters = 0
                Sinv = math.sqrt(P**2 + Qlim**2)
                ierr, Vinv = psspy.busdat(busgen_V, 'PU')
                ierr, Vinv_original = psspy.busdat(busgen_V, 'PU')
                Iinv = Sinv / Vinv

                # 1. If we run into the apparent power limit, start reducing Q.
                iters = 0
                print 'SINV = %f' % Sinv
                while (Sinv > Sbase) and iters < 100:
                    Qlim -= Q_step
                    if Qlim < 0.0: 
                        Qlim = 0.0
                        if P > 0: P -= P_step
                        elif P < 0: P += P_step
                    print '*********************************************************'
                    print 'Apparent power limit reached %1.4f, Qlim reduced to %1.4f' % (Sinv, Qlim)
                    print '*********************************************************'
                    #
                    psspy.machine_chng_2(busgen_V, '1', [_i,_i,_i,_i,_i,_i], [ P,_f, Qlim, -Qlim,_f,_f, Sbase,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
                    #if soln_ierr:
                    #    break
                    ierr, Vinv = psspy.busdat(busgen_V, 'PU')
                    ierr, cmpval = psspy.gendat(busgen_V)
                    Pinv = cmpval.real
                    Qinv = cmpval.imag
                    Sinv = math.sqrt(Pinv**2 + Qinv**2)
                    Iinv = Sinv / Vinv
                    iters += 1
                #
                # 2. If we run into the current limit, start reducing Q, then start reducing P.
                ierr, Vinv = psspy.busdat(busgen_V, 'PU')
                Iinv = Sinv / Vinv
                iters = 0
                while (Iinv > Sbase) and (Qlim > 0.1) and iters < 100:
                    Qlim -= Q_step
                    if Qlim < 0.0: 
                        Qlim = 0.0
                        if P > 0: P -= P_step
                        elif P < 0: P += P_step
                    print '**************************************************'
                    print 'Current limit reached %1.4f, Qlim reduced to %1.4f' % (Iinv, Qlim)
                    print '**************************************************'
                    #
                    psspy.machine_chng_2(busgen_V, '1', [_i,_i,_i,_i,_i,_i], [ P,_f, Qlim, -Qlim,_f,_f, Sbase,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
    #                if soln_ierr:
    #                    break
                    ierr, Vinv = psspy.busdat(busgen_V, 'PU')
                    ierr, cmpval = psspy.gendat(busgen_V)
                    Pinv = cmpval.real
                    Qinv = cmpval.imag
                    Sinv = math.sqrt(Pinv**2 + Qinv**2)
                    Iinv = Sinv / Vinv
                    iters += 1
                #
                time.sleep(0)
                ierr, Vinv = psspy.busdat(busgen_V, 'PU')
                delta_Vinv = abs(Vinv_original - Vinv)
                iters_out += 1
                
            ierr, Vinv = psspy.busdat(busgen_V, 'PU')
            ierr, cmpval = psspy.gendat(busgen_V)
            Pinv = cmpval.real
            Qinv = cmpval.imag
            Sinv = math.sqrt(Pinv**2 + Qinv**2)
            Iinv = Sinv / Vinv
#            ierr, cmpval = psspy.gendat(inf_bus)
            ierr, cmpval = psspy.brnflo(fbus, tbus, r"""1""")
            Pcp = cmpval.real
            Qcp = cmpval.imag
            #
            print '*****************************'
            print ' Vinv, Pinv, Qinv, Sinv, Iinv (pu)'
            print ' %1.4f, %1.2f, %1.4f, %1.4f, %1.4f' % (Vinv, P, Qinv, Sinv, Iinv / Sbase)
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
#            P_poc, Q_poc, V_poc, P_inv_gen1, Q_inv_gen1, V_inv_gen1, P_inv_gen2, Q_inv_gen2, V_inv_gen2 = monitor_SF(fbus, tbus, busPOC, busgen1, busgen2)
            P_inv_gen1 = Pinv
            Q_inv_gen1 = Qinv
            V_inv_gen1 = Vinv
            P_inv_gen2 = Pinv
            Q_inv_gen2 = Qinv
            V_inv_gen2 = Vinv
            P_poc = Pcp
            Q_poc = Qcp
            V_poc = Vcp
            Sinv = Sinv
            Iinv = Iinv
            
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
                Sinv_all_pos.append(Sinv)
                Iinv_all_pos.append(Iinv)

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
                Sinv_all_neg.append(Sinv)
                Iinv_all_neg.append(Iinv)


    P_inv_gen1_all = P_inv_gen1_all_pos + P_inv_gen1_all_neg
    Q_inv_gen1_all = Q_inv_gen1_all_pos + Q_inv_gen1_all_neg
    V_inv_gen1_all = V_inv_gen1_all_pos + V_inv_gen1_all_neg
    P_inv_gen2_all = P_inv_gen2_all_pos + P_inv_gen2_all_neg
    Q_inv_gen2_all = Q_inv_gen2_all_pos + Q_inv_gen2_all_neg
    V_inv_gen2_all = V_inv_gen2_all_pos + V_inv_gen2_all_neg
    P_poc_all = P_poc_all_pos + P_poc_all_neg
    Q_poc_all = Q_poc_all_pos + Q_poc_all_neg
    V_poc_all = V_poc_all_pos + V_poc_all_neg
    Sinv_all = Sinv_all_pos + Sinv_all_neg
    Iinv_all = Iinv_all_pos + Iinv_all_neg

    
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
    data_out['6 P_Poc'] = []
    data_out['7 Q_Poc'] = []    
    
    for idx in range(len(P_inv_gen1_all) / 2):
        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
        data_out['4 S_Inv'].append(Sinv_all[idx])
        data_out['5 I_Inv'].append(Iinv_all[idx])
        data_out['6 P_Poc'].append(P_poc_all[idx])
        data_out['7 Q_Poc'].append(Q_poc_all[idx]) 

    for idx in range(len(P_inv_gen1_all) - 1, (len(P_inv_gen1_all) / 2) - 1, -1):
        data_out['1 V_Inv'].append(V_inv_gen1_all[idx])
        data_out['2 P_Inv'].append(P_inv_gen1_all[idx])
        data_out['3 Q_Inv'].append(Q_inv_gen1_all[idx])
        data_out['4 S_Inv'].append(Sinv_all[idx])
        data_out['5 I_Inv'].append(Iinv_all[idx])
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