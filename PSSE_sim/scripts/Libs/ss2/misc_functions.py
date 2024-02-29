# -*- coding: utf-8 -*-
"""
Created on Tue Nov 14 16:53:58 2023

@author: 341510davu

FUNCTIONALITY:
    The funtions built in this script aim to set up a load folow to match with the reactive control model of a specific power plant. Control model considered are:
        1. Voltage droop control
        2. Power factor control
        3. Reactive power control
COMMENTS:
    Added funtion which can be used to search for a set of paramters aiming for more robust load folow solver step_solv_paras and lckd_solv_paras
    
"""

#PSS/E path
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

from datetime import datetime
start_time = datetime.now()

# Locating required existing folder paths
main_folder_path=os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
sys.path.append(main_folder_path+"\\PSSE_sim\\scripts\\Libs")
import auxiliary_functions as af
import numpy as np


# Try differernt set of parameters to solve the load flow with tap change
def step_solv_paras():
    """
    # Copy a temporary model to a temporary path and work on this model
    # Identify significant parameters to Load Flow FNSL: ACCN, DVLIM
#    ACCN            # POM: Change to 0.1 if tuff cases
#    TOLN
#    ITMXN
#    DVLIM = 0.99 # largest | (delta v)/|v| | for Newton solutions (0.99 by default) # POM: change to 0.05 in tuff cases
#    NDVFCT = 0.99
#    VCTOLQ
#    VCTOLV
    others:
#    ITMXN = 315 # Newton-Raphson maximum number of iterations (20 by default)
#    MXTPSS = 99 # maximum number of times taps and/or switched shunts are adjusted during power flow solutions (100 by default)
#    ACCN = 0.5 # Newton-Raphson acceleration factor (1.0 by default)
#    TOLN = 0.15 # Newton-Raphson mismatch convergence tolerance (default Newton power flow solution tolerance option setting)
#    BLOWUP = 5.0 # blow-up threshold (5.0 by default)
#    VCTOLQ = 0.1 # Newton-Raphson voltage controlled bus reactive power mismatch convergence tolerance (default Newton power flow solution tolerance option setting)
#    VCTOLV = 0.1E-04 # Newton-Raphson voltage controlled bus voltage error convergence tolerance (0.00001 by default)
#    SWVBND = 100.0 # percent of voltage controlling band mode switched shunts to be adjusted per power flow iteration (100.0 by default)
#  
    # Create a loop to find the best parameters for the load flow
        Two ways for conditions to loop: soluion mismatch or range of the paramters
        Aplly the setting
        run the load flow
        Check crash
        Check duaration of simulation
        Continue the loop if not sufficient
    # If sucessfully, save the model to original or required folder; delete the temporary path.
    """
    currntModelPath = os. getcwd()
    casefile = "temp"
    psspy.save(currntModelPath + "\\"+casefile + '_temp.sav')
    ierr, ACCN = psspy.prmdat('ACCN')
    ierr, DVLIM = psspy.prmdat('DVLIM')
    iter_num = 500
    Duration_req = 100
#    mismatch = 100
    psspy.solution_parameters_4([_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    mismatch=psspy.sysmsm()
    
    while mismatch > 1 and iter_num > 0:
        if ACCN > 0.05: 
            print "load flow does not converge - try a different set of solution parameters:"
            print("ACCN_old = ", ACCN)
            ACCN -= 0.05
            print("ACCN_new = ", ACCN)
        elif DVLIM > 0.05: 
            print "load flow does not converge - try a different set of solution parameters:"
            print("ACCN = ", ACCN, "DVLIM_old = ", DVLIM)
            DVLIM -= 0.05
            print("ACCN = ", ACCN, "DVLIM_new = ", DVLIM)
        else: 
            print "Cannot get the case converged"
            break
        
        psspy.case(currntModelPath + "\\"+casefile + '_temp.sav') #load the case again
        psspy.solution_parameters_4([_i,_i,_i,_i,_i],[_f,_f,_f,_f, ACCN,_f,_f,_f,_f,_f,_f,_f,_f,_f, DVLIM,_f,_f,_f,_f])
#            psspy.solution_parameters_4([1,315,20,99,10],[ 1.0, 1.6, 1.0, 0.0001,ACCN, 0.15, 1.0, 0.1E-04, 5.0, 0.7, 0.0001, 0.005, 1.0, 0.05, DVLIM, 0.99, 0.1, 0.1E-04, 100.0])
        start_time = datetime.now()
        psspy.fnsl([1,0,0,1,1,0,0,0])
        psspy.fnsl([1,0,0,1,1,0,0,0])
        psspy.fnsl([1,0,0,1,1,0,0,0])
        end_time = datetime.now()
#            Duration = end_time - start_time
        print('Duration: {}'.format(end_time - start_time))
        mismatch=psspy.sysmsm()
        iter_num -= 1
   
    if (abs(mismatch) <= 1): # load flow converged
        os.remove(currntModelPath + "\\"+casefile + '_temp.sav')
    else:
        af.print_volts(lower=0.9, upper=1.1)
        raise
    return mismatch, ACCN, DVLIM
    
    

# Try differernt set of parameters to solve the load flow with tap fixed
def lckd_solv_paras():
    currntModelPath = os. getcwd()
    casefile = "temp"
    psspy.save(currntModelPath + "\\"+casefile + '_temp.sav')
#    ierr, [ACCN, DVLIM] = psspy.prmdat('ACCN','DVLIM')
    ierr, ACCN = psspy.prmdat('ACCN')
    ierr, DVLIM = psspy.prmdat('DVLIM')    
    iter_num = 500
    Duration_req = 100
#    mismatch = 100
    psspy.solution_parameters_4([_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    mismatch=psspy.sysmsm()
    
    while mismatch > 1 and iter_num > 0:
        if ACCN > 0.05: 
            print "load flow does not converge - try different set of solution parameters:"
            print("ACCN_old = ", ACCN)
            ACCN -= 0.05
            print("ACCN_new = ", ACCN)
        elif DVLIM > 0.05: 
            print "load flow does not converge - try different set of solution parameters:"
            print("ACCN = ", ACCN, "DVLIM_old = ", DVLIM)
            DVLIM -= 0.05
            print("ACCN = ", ACCN, "DVLIM_new = ", DVLIM)
        else: 
            print "Cannot get the case converged"
            break
        
        psspy.case(currntModelPath + "\\"+casefile + '_temp.sav') #load the case again
        psspy.solution_parameters_4([_i,_i,_i,_i,_i],[_f,_f,_f,_f, ACCN,_f,_f,_f,_f,_f,_f,_f,_f,_f, DVLIM,_f,_f,_f,_f])
#            psspy.solution_parameters_4([1,315,20,99,10],[ 1.0, 1.6, 1.0, 0.0001,ACCN, 0.15, 1.0, 0.1E-04, 5.0, 0.7, 0.0001, 0.005, 1.0, 0.05, DVLIM, 0.99, 0.1, 0.1E-04, 100.0])
        start_time = datetime.now()
        psspy.fnsl([0,0,0,1,0,0,0,0])
        psspy.fnsl([0,0,0,1,0,0,0,0])
        psspy.fnsl([0,0,0,1,0,0,0,0])
        end_time = datetime.now()
#            Duration = end_time - start_time
        print('Duration: {}'.format(end_time - start_time))
        mismatch=psspy.sysmsm()
        iter_num -= 1

    if (abs(mismatch) <= 1): # load flow converged
        os.remove(currntModelPath + "\\"+casefile + '_temp.sav')
    else:
        af.print_volts(lower=0.9, upper=1.1)
        raise
    return mismatch, ACCN, DVLIM

# Initialise gens with voltage droop control mode:
def init_gens_vdc(gens_with_vdc, err_code = 0, auto_script = ''): 
    if len(gens_with_vdc['gens']) != 0:
        for i in range(0,20): #If multiple generator participate in the QV droop, then repeat the process to make sure the actual voltage settle well
            for gen in gens_with_vdc['gens']:
                ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
                ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
                if ival == 0 or ival2 == 4:
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
                    k_factor = 0.20 # regression factor
                    iter_num = 15
                    
                    ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                    q_gen_new = q_gen + delta_q*k_factor
                    if q_gen_new > gen['gen_q_max']: q_gen_new = gen['gen_q_max']
                    if q_gen_new < -gen['gen_q_max']: q_gen_new = -gen['gen_q_max']
                    psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new,q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"

#                    while abs(delta_q) > tol_q and iter_num > 0: # This loop may be deactivated to increase the convergence of the case
#                        ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
#                        q_gen_new = q_gen + delta_q*k_factor
#                        if q_gen_new > gen['gen_q_max']: q_gen_new = gen['gen_q_max']
#                        if q_gen_new < -gen['gen_q_max']: q_gen_new = -gen['gen_q_max']
#                        psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new,q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#                        auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
#                        psspy.fnsl([1,0,0,1,1,0,0,0])
#                        psspy.fnsl([1,0,0,1,1,0,0,0])
#                        auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
#                        auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
#                        af.print_volts(lower=0.9, upper=1.1) #for debugging
##                        step_solv_paras()
#                        if(af.test_convergence(method='fnsl',taps='step')>1.0):
#                            err_code = 1
#                        s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
#                        q_poc = -s_poc.imag # poc q MVAr
#                        delta_q = q_poc_req - q_poc
#                        iter_num -=1
    return err_code,auto_script

# Initialise gens with fix power factor control mode:
def init_gens_pf(gens_with_pf, err_code = 0, auto_script = ''):
    if len(gens_with_pf['gens']) != 0: 
        for gen in gens_with_pf['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
            if ival == 0 or ival2 == 4:
                print('GEN is OFF')
            else:
                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #Update the Control Mode to Not a wind machine so Qmax, Qmin can be updated
                auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1] #power measured at POC, from POC back to the SF (reversed power flow)
                p_poc = -s_poc.real # poc p MW generated from SF
                q_poc = -s_poc.imag # poc q MVAr generated from SF
                q_poc_req = np.tan(np.arccos(abs(gen['poc_pf'])))*np.sign(gen['poc_pf']) * gen['poc_p_gen']
#                 q_poc_req = np.tan(np.arccos(abs(gen['poc_pf'])))*np.sign(gen['poc_pf']) * p_poc
                if q_poc_req > gen['poc_q_max']: # Limit the compensation to Qcorner at POC
                    q_poc_req = gen['poc_q_max']
                if q_poc_req < -gen['poc_q_max']: 
                    q_poc_req = -gen['poc_q_max']
                        
#                delta_p = gen['poc_p_gen'] - p_poc
                delta_p = 0 
                delta_q = q_poc_req - q_poc
                tol_p = 0.5 #MW
                tol_q = 1.0 #MVAr
                k_factor = 1.00 # regression factor
                iter_num = 15
                
                ierr,p_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'P') # gen p MW
                ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                p_gen_new = p_gen + delta_p*k_factor
                q_gen_new = q_gen + delta_q*k_factor
                psspy.machine_chng_2(int(gen['gen_bus']),r"""1""",[_i,_i,_i,_i,_i,_i],[p_gen_new,_f, q_gen_new, q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],["+str(p_gen_new)+",_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"

#                while (abs(delta_p) > tol_p or abs(delta_q) > tol_q) and iter_num > 0:
                while abs(delta_q) > tol_q and iter_num > 0:
                    ierr,p_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'P') # gen p MW
                    ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                    p_gen_new = p_gen + delta_p*k_factor
                    q_gen_new = q_gen + delta_q*k_factor
                    psspy.machine_chng_2(int(gen['gen_bus']),r"""1""",[_i,_i,_i,_i,_i,_i],[p_gen_new,_f, q_gen_new, q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],["+str(p_gen_new)+",_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    af.print_volts(lower=0.9, upper=1.1) #for debugging
#                    step_solv_paras()
                    if(af.test_convergence(method='fnsl',taps='step')>1.0):
                        af.print_volts(lower=0.9, upper=1.1)
                        err_code = 1
                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    p_poc = -s_poc.real # poc p MW
                    q_poc = -s_poc.imag # poc q MVAr
#                     q_poc_req = np.tan(np.arccos(abs(gen['poc_pf'])))*np.sign(gen['poc_pf']) * p_poc
                    delta_p = gen['poc_p_gen'] - p_poc
                    delta_q = q_poc_req - q_poc
                    iter_num -=1
    return err_code,auto_script

#                # set the voltage setpoint of the generator to actual voltage at POC -> this will not impact the PQ level in pf control; but will prepare for Vcontrol after contingency when plant in hygrid control
#                poc_v_actual = af.get_bus_info(gen['poc_bus'],'PU')
#                poc_v_spnt = poc_v_actual[gen['poc_bus']]['PU'] #poc volt level
#                psspy.plant_data_4(gen['gen_bus'],0,[_i,_i],[ poc_v_spnt,gen['poc_bus']])
                    
# Initialise gens with fix reactive power control mode:
def init_gens_qfix(gens_with_qfix, err_code = 0, auto_script = ''):
    
    if len(gens_with_qfix['gens']) != 0: 
        for gen in gens_with_qfix['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
            if ival == 0 or ival2 == 4:
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
                    auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    psspy.fnsl([1,0,0,1,1,0,0,0])
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
#                    af.print_volts(lower=0.9, upper=1.1) #for debugging
#                    step_solv_paras()
                    if(af.test_convergence(method='fnsl',taps='step')>1.0):
                        err_code = 1
                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    q_poc = -s_poc.imag # poc q MVAr
                    delta_q = q_poc_req - q_poc
                    iter_num -=1
    return err_code,auto_script

# Initialise statcom - for LSF but can be used for other Statcorms
def ini_statcom(statcom, err_code = 0, auto_script = ''):
    if len(statcom['gens']) != 0: 
        for gen in statcom['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
            if ival == 0 or ival2 == 4:
                print('STATCOM is OFF')
            else:

                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f,gen['gen_q_ini'],gen['gen_q_ini'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #Initialise statcom at 0MVAr
                auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(gen['gen_q_ini'])+","+str(gen['gen_q_ini'])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                vlt_lvl = af.get_bus_info(gen['poc_bus'],'PU')
                vlt_lvl = vlt_lvl[gen['poc_bus']]['PU']
                psspy.plant_data_4(gen['gen_bus'],0,[_i,_i],[vlt_lvl,gen['poc_bus']])
#                psspy.fnsl([1,0,0,1,1,0,0,0])
#                psspy.fnsl([1,0,0,1,1,0,0,0])
#                auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
#                auto_script+="psspy.fnsl(psspy.fnsl([1,0,0,1,1,0,0,0])\n"
##                step_solv_paras()
#                if(af.test_convergence(method='fnsl',taps='step')>1.0):
#                    err_code = 1
    return err_code,auto_script

# Responsse of gens with voltage droop control mode when contingency occurs:
def lckd_gens_vdc(gens_with_vdc, err_code = 0, auto_script = ''): 
    if len(gens_with_vdc['gens']) != 0:
        for i in range(0,20): #If multiple generator participate in the QV droop, then repeat the process to make sure the actual voltage settle well
            for gen in gens_with_vdc['gens']:
                ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
                ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
                if ival == 0 or ival2 == 4:
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
                    k_factor = 0.20 # regression factor
                    iter_num = 15

                    ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
                    q_gen_new = q_gen + delta_q*k_factor
                    if q_gen_new > gen['gen_q_max']: q_gen_new = gen['gen_q_max']
                    if q_gen_new < -gen['gen_q_max']: q_gen_new = -gen['gen_q_max']
                    psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new,q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    auto_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    
#                    while abs(delta_q) > tol_q and iter_num > 0: # this loop will work with only one generator in Vdc mode. If multiple gens are in Vdc mode it may not work as when once Q is adjusted in one Gen, it will change the voltage of other gens next loop
#                        ierr,q_gen = psspy.macdat(gen['gen_bus'],gen['gen_id'],'Q') # gen q MVAr
#                        q_gen_new = q_gen + delta_q*k_factor
#                        if q_gen_new > gen['gen_q_max']: q_gen_new = gen['gen_q_max']
#                        if q_gen_new < -gen['gen_q_max']: q_gen_new = -gen['gen_q_max']
#                        psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, q_gen_new,q_gen_new,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
#                        auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(q_gen_new)+","+str(q_gen_new)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
#                        psspy.fnsl([0,0,0,1,0,0,0,0])
#                        psspy.fnsl([0,0,0,1,0,0,0,0])
#                        auto_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
#                        auto_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
##                        lckd_solv_paras()
#                        mismatch=psspy.sysmsm()
#                        if(mismatch>1.0):
##                        if(af.test_convergence(method='fnsl',taps='locked')>1.0):
#                            err_code = 1
#                        s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
#                        q_poc = -s_poc.imag # poc q MVAr
#                        delta_q = q_poc_req - q_poc
#                        iter_num -=1
    return err_code,auto_script
                        
# Responsse of gens with hybrid (PF and V_PI) or direct voltage control mode when contingency occurs:
def lckd_gens_hc(gens_with_pf_vc, err_code = 0, auto_script = ''):
    if len(gens_with_pf_vc['gens']) != 0:
        for gen in gens_with_pf_vc['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
            if ival == 0 or ival2 == 4:
                print('GEN is OFF')
            else:
                gen_q_max = gen['gen_q_max']
                gen_q_min = -gen['gen_q_max']
                ierr, mc_q_max=psspy.macdat(gen['gen_bus'],gen['gen_id'],'QMAX') #current maximum reactive power of the machine
                ierr, mc_q_min=psspy.macdat(gen['gen_bus'],gen['gen_id'],'QMIN')
                if gen_q_max < mc_q_max: gen_q_max = mc_q_max # keep original value if it provide a wider range
                if gen_q_min > mc_q_min: gen_q_min = mc_q_min
                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, gen_q_max, gen_q_min,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #release the Q capability of the plant
                auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(gen_q_max)+","+str(gen_q_min)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                
#                iter_num = 5 # realease by steps if needed
#                for i in range(1, iter_num):
#                    gen_q_max_ = gen_q_max * i / iter_num
#                    gen_q_min_ = gen_q_min * i / iter_num
#                    psspy.machine_chng_2(gen['gen_bus'],r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, gen_q_max_, gen_q_min_,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) #release the Q capability of the plant
#                    standalone_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(gen_q_max_)+","+str(gen_q_min_)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"

#                    psspy.fnsl([0,0,0,1,0,0,0,0])
#    ##                af.print_volts(lower=0.9, upper=1.1) #for debugging
#                    psspy.fnsl([0,0,0,1,0,0,0,0])
#                    psspy.fnsl([0,0,0,1,0,0,0,0])
#    #                lckd_solv_paras()
#                    if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
    return err_code,auto_script

def lckd_gens_qfix(gens_with_qfix, err_code = 0, auto_script = ''):
    if len(gens_with_qfix['gens']) != 0: 
        for gen in gens_with_qfix['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
            if ival == 0 or ival2 == 4:
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
                    auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[_i,_i,_i,_i,_i,_i],[_f,_f,"+str(gen_q_max)+","+str(gen_q_min)+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    psspy.fnsl([0,0,0,1,0,0,0,0])
                    auto_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    auto_script+="psspy.fnsl(psspy.fnsl([0,0,0,1,0,0,0,0])\n"
                    af.print_volts(lower=0.9, upper=1.1) #for debugging
#                    lckd_solv_paras()
                    if(af.test_convergence(method='fnsl',taps='locked')>1.0):
                        af.print_volts(lower=0.9, upper=1.1)
                        err_code = 1

                    s_poc = psspy.brnflo(gen['poc_bus'],gen['ibus'],'1')[1]
                    q_poc = -s_poc.imag # poc q MVAr
                    delta_q = q_poc_req - q_poc
                    iter_num -=1
    return err_code,auto_script

# Reponse of the statcom
def lckd_statcom(statcom, err_code = 0, auto_script = ''):
    if len(statcom['gens']) != 0: 
        for gen in statcom['gens']:
            ierr,ival = psspy.macint(gen['gen_bus'],gen['gen_id'],'STATUS')
            ierr,ival2 = psspy.busint(gen['gen_bus'],'TYPE')
            if ival == 0 or ival2 == 4:
                print('STATCOM is OFF')
            else:
                psspy.machine_chng_2(gen['gen_bus'],r"""1""",[1,_i,_i,_i,_i,_i],[_f,_f,gen['gen_q_max'],gen['gen_q_min'],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f]) # release the capacity of the statcom for voltage regulation
                auto_script+="psspy.machine_chng_2(" + str(gen['gen_bus'])+",'1',[1,_i,_i,_i,_i,_i],[_f,_f,"+str(gen['gen_q_max'])+","+str(gen['gen_q_min'])+",_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])\n"
#                psspy.fnsl([0,0,0,1,0,0,0,0])
#                psspy.fnsl([0,0,0,1,0,0,0,0])
##                lckd_solv_paras()
#                if(af.test_convergence(method='fnsl',taps='locked')>1.0):raise
    return err_code,auto_script

###############################################################################                        
def main():

#    gens_with_pf = {'gens':[{'gen_bus':30714,'gen_id':'1','poc_bus':36711,'ibus':36712,'poc_pf':-0.96,'gen_q_max':21.1857,'poc_q_max':21.1857,'poc_p_gen':66.00},
#                             {'gen_bus':30715,'gen_id':'1','poc_bus':36711,'ibus':36713,'poc_pf':-0.96,'gen_q_max':17.6143,'poc_q_max':17.6143,'poc_p_gen':34.00},
#                             {'gen_bus':30709,'gen_id':'1','poc_bus':36716,'ibus':36717,'poc_pf':-0.990001,'gen_q_max':19.25625,'poc_q_max':19.25625,'poc_p_gen':48.75},
#                             {'gen_bus':30710,'gen_id':'1','poc_bus':36716,'ibus':36718,'poc_pf':-0.990001,'gen_q_max':10.36875,'poc_q_max':10.36875,'poc_p_gen':26.25},
#                             {'gen_bus':302,'gen_id':'1','poc_bus':1,'ibus':102,'poc_pf':-0.995,'gen_q_max':30.0276,'poc_q_max':30.0276,'poc_p_gen':76},
#                             #{'gen_bus':367169,'gen_id':'1','poc_bus':36721,'ibus':367167,'poc_pf':-0.96,'gen_q_max':15.8,'gen_p_gen':-40},
#                             #{'gen_bus':367166,'gen_id':'1','poc_bus':36721,'ibus':367164,'poc_pf':-0.96,'gen_q_max':15.8,'gen_p_gen':-40}
#                             ]}
#
#    gens_with_vdc = {'gens':[{'gen_bus':367169,'gen_id':'1','poc_bus':36721,'ibus':367167,'poc_v_spt':0.97,'poc_v_dbn':0.0,'poc_q_max':15.8,'poc_drp_pct':3.3},
#                         {'gen_bus':367166,'gen_id':'1','poc_bus':36721,'ibus':367164,'poc_v_spt':0.97,'poc_v_dbn':0.0,'poc_q_max':15.8,'poc_drp_pct':3.3},
#                         ]}
#    
#    gens_with_pf_vc = {'gens':[{'gen_bus':30714,'gen_id':'1','poc_bus':36711,'ibus':36712,'poc_pf':-0.98,'gen_q_max':21.1857,'gen_p_gen':-66.00},
#                             {'gen_bus':30715,'gen_id':'1','poc_bus':36711,'ibus':36713,'poc_pf':-0.98,'gen_q_max':17.6143,'gen_p_gen':-34.00},
#                             {'gen_bus':302,'gen_id':'1','poc_bus':1,'ibus':102,'poc_pf':-0.995,'gen_q_max':30.0276,'gen_p_gen':-76},
#                             #{'gen_bus':100001,'gen_id':'1','poc_bus':37703,'ibus':100001,'poc_pf':0.0,'gen_q_max':3.5,'gen_p_gen':0.0}
#                             #{'gen_bus':367169,'gen_id':'1','poc_bus':36721,'ibus':367167,'poc_pf':-0.985,'gen_q_max':15.8,'gen_p_gen':-40},
#                             #{'gen_bus':367166,'gen_id':'1','poc_bus':36721,'ibus':367164,'poc_pf':-0.985,'gen_q_max':15.8,'gen_p_gen':-40}
#                             ]}
#
#    
#    init_gens_pf(gens_with_pf)
    currntModelPath = r"C:\work\3. Grid - LSF\1. Main Test Environment\20220525\PSSE_sim\base_model\LowLoad testing"
    casefile = "NoStatCom_LANSF_low_genoff"
    psspy.case(currntModelPath +"\\"+ casefile + '.sav')
    
    mismatch, ACCN, DVLIM = step_solv_paras()
    print [mismatch, ACCN, DVLIM]
    
    psspy.branch_chng_3(36700,36719,r"""1""",[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[ 79.4, 90.3,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],_s)

#    mismatch, ACCN, DVLIM = step_solv_paras(currntModelPath,casefile)
    mismatch, ACCN, DVLIM = lckd_solv_paras()
    print [mismatch, ACCN, DVLIM]
    
if __name__ == '__main__':
    main()