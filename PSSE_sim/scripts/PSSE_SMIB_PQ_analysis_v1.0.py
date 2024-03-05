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
import os, sys
import math
import time
import pandas as pd
import scipy.interpolate
from contextlib import contextmanager
from win32com.client import Dispatch
timestr = time.strftime("%Y%m%d-%H%M%S") 


###############################################################################
#USER INPUTS
###############################################################################

TestDefinitionSheet = r'20230828_SUM_TESTINFO_V2.xlsx'

case_name = 'HSFBESS_SMIB_V1.sav' #to match with the model name provided in \PSSE_sim\base_model\SMIB
source_type = "BESS_PV" #"BESS" # "PV" #"BESS_PV"
temp_cases = ["35degC","50degC"] #"35degC" # "50degC" #"40degC"
Vinfinite = [0.90, 1.0, 1.10] #0.90, 1.0, 1.10
BESS_pct = 1.0 #[0.0-1.0]BESS setpoint to be a proportion of POC active power setpoint

simulation_batches=['S5251']

try: testRun = timestr + '_' + simulation_batches[0] #define a test name for the batch or configuration that is being tested -> link to time stamp for auto update
except: testRun = timestr

# Generator (OEM) inputs
PV_INV_num = 0.0 # PV INV quantity aggregated
PV_INV_rate = 4.2 # MVA INV individual rating
BESS_INV_num = 37.0 # BESS INV quantity aggregated
BESS_INV_rate = 3.6 # MVA INV individual rating
    
# SMIB model inputs
tapstep = 1 
bus_inf = 10000 # INF bus to control the voltage for over and under excitation mode
bus_inf_scr = 334081  # To update the SCR to very big during analysis
bus_POC_fr = 334081 # POC bus from: connect the branch for measuring power flow
bus_POC_to = 334090 # POC bus to: connect the branch for measuring power flow
pq_drt = -1 # direction of the power measured =-1 if measuring bus_POC_to to bus_POC_fr

# Branch 1 - PV leg
busgen1 = 334094 # Generator bus
busTx1 = 334092
# Branch 2 - BESS leg
busgen2 = 334095
busTx2 = 334093

Sbase_analysis_PV = PV_INV_num * PV_INV_rate
Sbase_analysis_BESS = BESS_INV_num * BESS_INV_rate

busgen_V_PV = busgen1
busgen_V_BESS = busgen2
Prated = 119.0
Pmax_POC = 100 
Pmin_POC = -100
q_poc_ner = 0.395*Prated # Q required at POC as NER = 0.395*Prated
    

Pmax_loop = Pmax_POC + 1.9 # account for the losses from POC to INV terminal
Pmin_loop = Pmin_POC # account for the losses from POC to INV terminal
# Define the range for looping P
Ploop_range =  [Pmin_loop] + range(int(math.ceil(Pmin_loop)), int(math.floor(Pmax_loop)), 2) + [Pmax_loop]  
#Ploop_range =  [Pmin_loop] # for debuging only

# Capbanks to be turnoff
capbanks = []
capbanks = [
        {"bus": 334091, "id": "4"},
        {"bus": 334091, "id": "3"}
        ]

# Buses to be deactivated
bus_dis_list = []
#bus_dis_list = [busgen1, busTx1] #[busgen1, busTx1] [busgen1, busgen2] all the buses in the PV/BESS leg to be disconnected

# Gen bus list
bus_gen_list = [busgen1, busgen2]  #bus_gen_list = [busgen1] [busgen2] generator buses to be mornitored voltage and power flow

###############################################################################
# Supporting functions
###############################################################################
#from __future__ import with_statement
#from contextlib import contextmanager

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
    import math
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
#    Smax_out = abs(derating_profile['x_data'][-1])
    S_individual = [math.sqrt(math.pow(p,2) + math.pow(q,2)) for p, q in zip(derating_profile['x_data'], derating_profile['y_data'])]
    Smax_out = max(S_individual)
    return Qlim_out, Smax_out


def update_Slim(y_interp_BESS,Vinv_BESS,Slim_110_BESS,Slim_90_BESS):
    # Update Sbase with new derating factor based on the terminal voltage level
    try:
        Slim_ind = y_interp_BESS(Vinv_BESS)
    except: # if the voltage is outside the range [0.9-1.1pu]
        if Vinv_BESS >= 1.1: Slim_ind = Slim_110_BESS
        elif Vinv_BESS <= 0.9: Slim_ind = Slim_90_BESS
    return Slim_ind

def apply_derating(P_individual_PV, Vinv_PV, case_text): 
    # Alocate the derating profile (individual PQ curve) at differnt terminal voltage levels (0.9, 1.0 and 1.1pu)
    profile_90_PV = DeratingDict[case_text+'_0.9']
    profile_100_PV = DeratingDict[case_text+'_1.0']
    profile_110_PV = DeratingDict[case_text+'_1.1']

    # + At each active power level, determine the associated Qlim for differnt voltage level (from individual PQ curve) Qlim_0.9, Qlim_1.0 and Qlim_1.1
    # + From Qlim, calculate the Slim. Slim_0.9, Slim_1.0 and Slim_1.1. This is needed for derating interporation (S derate linearly with V)
    # + Check and make sure the Slim is capped by Smax
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
    # if Q is zero then set the Slim to Smax (Smax is at Q = 0)-> consider Q = 0 in all voltages for iterpolation
    if Qlim_90_PV == 0:
        Slim_90_PV = Smax_90_PV
        Slim_100_PV = Smax_100_PV
        Slim_110_PV = Smax_110_PV
    # Interporating function of the S on INV terminal voltage level
    S_interp_PV = scipy.interpolate.interp1d([0.9, 1, 1.1], [Slim_90_PV, Slim_100_PV, Slim_110_PV])            
    Slim_PV_ind = update_Slim(S_interp_PV,Vinv_PV,Slim_110_PV,Slim_90_PV) # MVA - interpolate Slim level each individual inverter based on updated INV voltage
    
    return Slim_PV_ind

def turn_off_capbank(capbanks):
    if capbanks != []:
        for capbank in capbanks:
            psspy.shunt_chng(capbank["bus"],capbank["id"],0,[_f,_f])

                    
###############################################################################
# Define Project Paths
###############################################################################
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
ResultsDir = OutputDir+"\\PQ_curve"
make_dir(ResultsDir)


outputResultPath=ResultsDir+"\\"+testRun
make_dir(outputResultPath)

###############################################################################
# Import additional functions, # Initialise PSSE
###############################################################################
import misc_functions as mf
import auxiliary_functions as af
import readtestinfo as readtestinfo
import psspy
import redirect
import shutil
import psse34
import pssarrays
import bsntools
import readtestinfo as readtestinfo
redirect.psse2py()

sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
sys.path.append(sys_path_PSSE)
os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE


with silence():
    psspy.psseinit(80000)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()

###############################################################################
# Reactive Power Infor
###############################################################################

# return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['ProjectDetails', 'SimulationSettings', 'ModelDetailsPSSE', 'SetpointsDict', 'ScenariosSMIB', 'Profiles'])
return_dict =  readtestinfo.readTestdef(testDefinitionDir+"\\"+TestDefinitionSheet, ['PowerCapability'])
#ProjectDetailsDict = return_dict['ProjectDetails']
#PSSEmodelDict = return_dict['ModelDetailsPSSE']
DeratingDict = return_dict['PowerCapability']

out_file_name = "S5251_PQ curve results.xlsx" # for storing the results
writer = pd.ExcelWriter(outputResultPath+'\\'+out_file_name,engine = 'xlsxwriter')

psspy.case(base_model_workspace+"\\" + case_name) # Loac the case
for bus_dis_i in bus_dis_list:#if source_type in ["BESS", "PV"] deactivate the irrelavent buses
    psspy.dscn(bus_dis_i)
psspy.fnsl([1,0,0,1,1,0,99,0])
psspy.fnsl([1,0,0,1,1,0,99,0])
psspy.fnsl([1,0,0,1,1,0,99,0])

# turn off the capbank if needed:
turn_off_capbank(capbanks)


# Force the voltage at POC by changing grid impedance to zero, solve with tap enable 
psspy.branch_chng_3(bus_inf,bus_inf_scr,r"""1""",[_i,_i,_i,_i,_i,_i],[0.0,0.00001,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
psspy.fdns([1,0,0,1,1,1,99,0])
psspy.fdns([1,0,0,1,1,1,99,0])

# Save the ini case into the working (model copies) folder    
wrkg_spc = createPath(ModelCopyDir+"\\"+testRun) # working space folder
psspy.save(wrkg_spc +"\\" +case_name) #save the case with intial conditions


del_Vinv_tol = 0.001
iters_out = 0
P_step = 0.25
Q_step = 0.5 
                    
for temperature in temp_cases: # loop through different temperature cases
    if temperature == "35degC":temp_text = '35'
    elif temperature == "40degC":temp_text = '40'
    else:temp_text = '50' #"50degC"
    
    for index, Vcp in enumerate(Vinfinite): # Loop through different voltage cases
        vol_case_name = str(Vcp)

        psspy.plant_data(bus_inf,0,[ Vcp,_f])
        psspy.fnsl([1,0,0,1,1,0,0,0])
        psspy.fnsl([1,0,0,1,1,0,0,0])
                
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
        data_out_test = pd.DataFrame()
        data_out_test['1 V_Inv'] = []
        inv_gen1 = pd.DataFrame()
        inv_gen1['P'] = []
        inv_gen1['P_pos'] = []
        
        for P in Ploop_range: # loop through different value of active power

            for Vset_inv in [1.2, 0.7]: # change voltage setpoint of generator to get a Q range. For each case, solve the case with the tap locked
                if Vset_inv >1:mode_text = 'OE' # when Qgen is positive (over excited)
                else:mode_text = 'UE' # when Qgen is negative (under excited)
                
                # Apply the voltage target to all gens
                for gen_i in bus_gen_list:
                    psspy.plant_data(gen_i,0,[ Vset_inv,_f])
                soln_ierr = psspy.fnsl([tapstep,0,0,1,1,0,0,0])
                psspy.fnsl([tapstep,0,0,1,1,0,0,0])
                psspy.fnsl([tapstep,0,0,1,1,0,0,0])
     
                #initialise the maximum apparent power for both PV and BESS
                Sbase_PV = Sbase_analysis_PV
                Sbase_BESS = Sbase_analysis_BESS
                
                # divide the Power between PV and BESS
                # 1. set the BESS following the defined percentage
                # 2. remaining required power comes from PV. If PF is at limit then the pecentage of BESS will be updated to make sure it compensate for the P required at POC
    #            P_BESS_ori = BESS_pct * P #set the PBESS to the defined percentage
                P_BESS = BESS_pct * P #set the PBESS to the defined percentage
                if P_BESS > Sbase_BESS: P_BESS = Sbase_BESS #limit the P_BESS if outside the range
                if P_BESS < -Sbase_BESS: P_BESS = -Sbase_BESS #limit the P_BESS if outside the range
                
                P_PV = P - P_BESS # the remaining required active power to come from PV
                if P_PV > Sbase_PV: P_PV = Sbase_PV #limit the P_PV if outside the range
                if P_PV < 0: P_PV = 0 #limit the P_PV if outside the range
                P_BESS = P - P_PV #after limit the PV, refresh the power contributed by BESS -> make sure P at POC is achieved
    #            if P_BESS != P_BESS_ori: print 'Contribution from BESS is updated to %1.2f' % (P_BESS/P) # if the proportion from BESS changes, note it down
                

                # Calculate the Qlim from the allocated Sbase and P 
                if Sbase_BESS > abs(P_BESS): Qlim_BESS = math.sqrt(Sbase_BESS**2 - P_BESS**2) 
                else: Qlim_BESS = 0
                if Sbase_PV > abs(P_PV): Qlim_PV = math.sqrt(Sbase_PV**2 - P_PV**2) 
                else:Qlim_PV = 0 

                # Apply the PQS to the model
                psspy.machine_chng_2(busgen_V_PV, '1', [_i,_i,_i,_i,_i,_i], [ P_PV,_f, Qlim_PV, -Qlim_PV,_f,_f, Sbase_PV,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                psspy.machine_chng_2(busgen_V_BESS, '1', [_i,_i,_i,_i,_i,_i], [ P_BESS,_f, Qlim_BESS, -Qlim_BESS,_f,_f, Sbase_BESS,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                    

                # Measure V P Q
                ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                ierr, cmpval = psspy.gendat(busgen_V_BESS)
                Pinv_BESS = cmpval.real
                Qinv_BESS = cmpval.imag
                Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2) 

                ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                ierr, cmpval = psspy.gendat(busgen_V_PV)
                Pinv_PV = cmpval.real
                Qinv_PV = cmpval.imag
                Sinv_PV = math.sqrt(P_PV**2 + Qinv_PV**2)  

                # calculate individual power from INV
                if PV_INV_num == 0: P_individual_PV = 0
                else: P_individual_PV = Pinv_PV/PV_INV_num  #MW
                if BESS_INV_num == 0: P_individual_BESS = 0
                else: P_individual_BESS = Pinv_BESS/BESS_INV_num  #MW

                # Iterpolate for new base base on the active power level and voltage level at inv
                if P_individual_BESS >=0: srce_text = 'Dis'#when discharging
                else: srce_text = 'Cha'# When charging
                case_text = srce_text+'-'+mode_text+'_'+temp_text
                Slim_BESS_ind = apply_derating(P_individual_BESS,Vinv_BESS, case_text)
                Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV

                case_text = 'PV'+'_'+mode_text+'_'+temp_text
                Slim_PV_ind = apply_derating(P_individual_PV,Vinv_PV, case_text)
                Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV

                # check if Q POC is sufficient -> only need to check when P is close to Pmin
                if P == Pmin_loop:
                    ierr, cmpval = psspy.brnflo(bus_POC_fr, bus_POC_to, r"""1""")
                    q_poc_check = pq_drt * cmpval.imag
                    q_poc_delta = abs(abs(q_poc_check) - q_poc_ner)
                    k_q_poc = 0.25
                    iters =0
                    while  q_poc_delta>0.2 and iters < 50: 
                        Qlim_BESS -= q_poc_delta*k_q_poc
                        Qlim_PV -= q_poc_delta*k_q_poc
                        if Qlim_BESS < 0.0: Qlim_BESS = 0.0
                        if Qlim_PV < 0.0: Qlim_PV = 0.0
                        
                        # Apply the new base to the model
                        psspy.machine_chng_2(busgen_V_PV, '1', [_i,_i,_i,_i,_i,_i], [ P_PV,_f, Qlim_PV, -Qlim_PV,_f,_f, Sbase_PV,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        psspy.machine_chng_2(busgen_V_BESS, '1', [_i,_i,_i,_i,_i,_i], [ P_BESS,_f, Qlim_BESS, -Qlim_BESS,_f,_f, Sbase_BESS,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                        psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                        psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                        
                        ierr, cmpval = psspy.brnflo(bus_POC_fr, bus_POC_to, r"""1""")
                        q_poc_check = pq_drt * cmpval.imag
                        q_poc_delta = abs(abs(q_poc_check) - q_poc_ner)
                        
                        # Measure V P Q and update S
                        ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                        ierr, cmpval = psspy.gendat(busgen_V_BESS)
                        Pinv_BESS = cmpval.real
                        Qinv_BESS = cmpval.imag
                        Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2) 
        
                        ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                        ierr, cmpval = psspy.gendat(busgen_V_PV)
                        Pinv_PV = cmpval.real
                        Qinv_PV = cmpval.imag
                        Sinv_PV = math.sqrt(P_PV**2 + Qinv_PV**2)  
    
                        # calculate individual power from INV
                        if PV_INV_num == 0: P_individual_PV = 0
                        else: P_individual_PV = Pinv_PV/PV_INV_num  #MW
                        if BESS_INV_num == 0: P_individual_BESS = 0
                        else: P_individual_BESS = Pinv_BESS/BESS_INV_num  #MW
        
                        # Iterpolate for new base base on the active power level and voltage level at inv
                        if P_individual_BESS >=0: srce_text = 'Dis'#when discharging
                        else: srce_text = 'Cha'# When charging
                        case_text = srce_text+'-'+mode_text+'_'+temp_text
                        Slim_BESS_ind = apply_derating(P_individual_BESS,Vinv_BESS, case_text)
                        Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
        
                        case_text = 'PV'+'_'+mode_text+'_'+temp_text
                        Slim_PV_ind = apply_derating(P_individual_PV,Vinv_PV, case_text)
                        Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                        
                        iters += 1
                    
                # check if the base is violated, then reduce Qlim. If Qlim is at limit, then reduce P
                iters = 0
                while ((Sinv_BESS > Sbase_BESS) or (Sinv_PV > Sbase_PV)) and iters < 200:
#                    # check if Q POC is sufficient
#                    ierr, cmpval = psspy.brnflo(bus_POC_fr, bus_POC_to, r"""1""")
#                    q_poc_check = pq_drt * cmpval.imag
#                    q_poc_delta = abs(abs(q_poc_check) - q_poc_ner)
#                    if  q_poc_delta>0.5: 
#                        Qlim_BESS -= q_poc_delta*0.5
#                        Qlim_PV -= q_poc_delta*0.5
#                        if Qlim_BESS < 0.0: Qlim_BESS = 0.0
#                        if Qlim_PV < 0.0: Qlim_PV = 0.0
                
                    if (Sinv_BESS > Sbase_BESS): # If BESS overloaded
                        Qlim_BESS -= Q_step
                        if Qlim_BESS < 0.0: 
                            Qlim_BESS = 0.0
#                            if P_BESS >= 0: P_BESS -= P_step
#                            else: P_BESS += P_step
                            if P_BESS >= 0: P_BESS = Sbase_BESS
                            else: P_BESS = -Sbase_BESS
                            P_PV = P - P_BESS # if change P_BESS, then if PV is not overloaded need to ajust P_PV to make sure P_POC is constant
                            if (P_PV > Sinv_PV): P_PV = Sinv_PV#if PV is not overloaded yet
                            if (P_PV < 0): P_PV = 0
                            
                    if (Sinv_PV > Sbase_PV): # If PV overloaded
                        Qlim_PV -= Q_step
                        if Qlim_PV < 0.0: 
                            Qlim_PV = 0.0
#                            P_PV -= P_step
                            P_PV = Sbase_PV
                            P_BESS = P - P_PV

                    # Apply the new base to the model
                    psspy.machine_chng_2(busgen_V_PV, '1', [_i,_i,_i,_i,_i,_i], [ P_PV,_f, Qlim_PV, -Qlim_PV,_f,_f, Sbase_PV,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.machine_chng_2(busgen_V_BESS, '1', [_i,_i,_i,_i,_i,_i], [ P_BESS,_f, Qlim_BESS, -Qlim_BESS,_f,_f, Sbase_BESS,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                    psspy.fnsl([tapstep,0,0,1,1,0,0,0]) 
                    
                    # Measure V P Q and update S
                    ierr, Vinv_BESS = psspy.busdat(busgen_V_BESS, 'PU')
                    ierr, cmpval = psspy.gendat(busgen_V_BESS)
                    Pinv_BESS = cmpval.real
                    Qinv_BESS = cmpval.imag
                    Sinv_BESS = math.sqrt(Pinv_BESS**2 + Qinv_BESS**2) 
    
                    ierr, Vinv_PV = psspy.busdat(busgen_V_PV, 'PU')
                    ierr, cmpval = psspy.gendat(busgen_V_PV)
                    Pinv_PV = cmpval.real
                    Qinv_PV = cmpval.imag
                    Sinv_PV = math.sqrt(P_PV**2 + Qinv_PV**2)  

                    # calculate individual power from INV
                    if PV_INV_num == 0: P_individual_PV = 0
                    else: P_individual_PV = Pinv_PV/PV_INV_num  #MW
                    if BESS_INV_num == 0: P_individual_BESS = 0
                    else: P_individual_BESS = Pinv_BESS/BESS_INV_num  #MW
    
                    # Iterpolate for new base base on the active power level and voltage level at inv
                    if P_individual_BESS >=0: srce_text = 'Dis'#when discharging
                    else: srce_text = 'Cha'# When charging
                    case_text = srce_text+'-'+mode_text+'_'+temp_text
                    Slim_BESS_ind = apply_derating(P_individual_BESS,Vinv_BESS, case_text)
                    Sbase_BESS = Slim_BESS_ind * BESS_INV_num # Maximum Slim of all INV
    
                    case_text = 'PV'+'_'+mode_text+'_'+temp_text
                    Slim_PV_ind = apply_derating(P_individual_PV,Vinv_PV, case_text)
                    Sbase_PV = Slim_PV_ind * PV_INV_num # Maximum Slim of all INV
                    
                    iters += 1


                
                
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
                Pcp = pq_drt * cmpval.real
                Qcp = pq_drt * cmpval.imag
                #
                print '*****************************'
                print ' Vinv, Pinv, Qinv, Sinv, Iinv (pu)'
    #            print ' %1.4f, %1.2f, %1.4f, %1.4f, %1.4f' % (Vinv_BESS, P, Qinv_BESS, Sinv_BESS, Iinv_BESS / Sbase_BESS)
                print '*****************************'
                #

    
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
#                    data_out_test.append(P_inv_gen1)
#                    inv_gen1['P_pos'] = inv_gen1['P_pos'].append(P_inv_gen1)
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
    
#        inv_gen1['P'] = pd.concat(inv_gen1['P_pos'],inv_gen1['P_pos'])
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
    

    
    
    
        #Writing summary results into one file

        data_out = {}
#        data_out = pd.DataFrame()
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
    
    
            

        temp_name = temperature + "_" + source_type #35degC_BESS"
        df_out = pd.DataFrame.from_dict(data = data_out)
        df_out.to_excel(writer, sheet_name = str(temp_name +"_"+ vol_case_name), index=True )

#        results[df_to_sheet['volt_levels']['df']].to_excel(writer,sheet_name = df_to_sheet['volt_levels']['sht']+'_'+snapshot_name+'_'+config, index=True )

        
writer.close() 

