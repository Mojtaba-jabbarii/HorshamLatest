# -*- coding: utf-8 -*-
"""
Created on Mon Jun 29 09:44:38 2020

@author: Mervin Kall
"""
#making some mods

import pandas as pd

def readTestdef(testdefSheetPath, relevant_tabs='all'):
    return_dict={}
    #--------------------------------------------------------------------------
    #PROJECT DETAILS
    if('ProjectDetails' in relevant_tabs or relevant_tabs=='all'):
        #PROJECT DETAILS
        ProjectDetailsSheet=pd.read_excel(testdefSheetPath, sheet_name="ProjectDetails", usecols="B,C,D", keep_default_na=False)
        #print(ProjectDetailsSheet)
        ProjectDetailsDict=dict(zip(ProjectDetailsSheet.ParamIdentifier, ProjectDetailsSheet.Value))
        if('' in ProjectDetailsDict.keys()):
            del ProjectDetailsDict['']
        #print(ProjectDetailsDict)
        print("Read project details")
        return_dict['ProjectDetails']=ProjectDetailsDict
        
    #--------------------------------------------------------------------------
    #SIMULATION SETTINGS
    if('SimulationSettings' in relevant_tabs or relevant_tabs=='all'):    
        SimulationSettingsSheet=pd.read_excel(testdefSheetPath, sheet_name="SimulationSettings", usecols="A,B,C", keep_default_na=False, skiprows=0)
        SimulationSettingsDict=dict(zip(SimulationSettingsSheet.Parameter, SimulationSettingsSheet.Value))
        if('' in SimulationSettingsDict.keys()):
            del SimulationSettingsDict['']
        #print(ProjectDetailsDict)
        print("Read simulations setting details")
        return_dict['SimulationSettings']=SimulationSettingsDict
    
    #-------------------------------------------------------------------------
    #PSCAD-SPECIFIC SETTINGS
    if('ModelDetailsPSCAD' in relevant_tabs or relevant_tabs=='all'):  
        PSCADmodelSheet=pd.read_excel(testdefSheetPath, sheet_name="ModelDetailsPSCAD", usecols="A,B", keep_default_na=False)
        #print(FileSettingsSheet)
        PSCADmodelDict=dict(zip(PSCADmodelSheet.ParamIdentifier, PSCADmodelSheet.Value))
        if('' in PSCADmodelDict.keys()):
            del PSCADmodelDict['']
        #print(FileSettingsDict)
        print("Read PSCAD model settings")
        return_dict['ModelDetailsPSCAD']=PSCADmodelDict
    
    #-------------------------------------------------------------------------
    #PSSE-SPECIFIC SETTINGS
    if('ModelDetailsPSSE' in relevant_tabs or relevant_tabs=='all'):  
        PSSEmodelSheet=pd.read_excel(testdefSheetPath, sheet_name="ModelDetailsPSSE", usecols="A,B", keep_default_na=False)
        #print(FileSettingsSheet)
        PSSEmodelDict=dict(zip(PSSEmodelSheet.ParamIdentifier, PSSEmodelSheet.Value))
        if('' in PSSEmodelDict.keys()):
            del PSSEmodelDict['']
        #print(FileSettingsDict)
        print("Read PSSE model settings")
        return_dict['ModelDetailsPSSE']=PSSEmodelDict
    #-------------------------------------------------------------------------
    #SETPOINTS
    if('Setpoints' in relevant_tabs or relevant_tabs=='all'):      
        # Format: dict with numbers as keys, corresponsing to the setpoint ID.
            #every dict entry contains: { P:, Q:, SCR:, X_R:, V_POC:, componentSettings{Names:[], project_nrs:[], modules:[], pscad_ids:[], symbols:[] , values[]}}
        #retrieve
        SetpointsSheet=pd.read_excel(testdefSheetPath, sheet_name="Setpoints", usecols="B,C,D,E,F", keep_default_na=False, skiprows=0) #
            #detect in which row the component settings start (PSCAD-specific parameters)
        skip_rows=0
        while(SetpointsSheet.iloc[skip_rows]['Unnamed: 2']!='ProjectNo'):
            skip_rows+=1
        skip_rows+=1
            
        SetpointsSheet_compInfo=pd.read_excel(testdefSheetPath, sheet_name="Setpoints", usecols="B,C,D,E,F", keep_default_na=False, skiprows=skip_rows)  
        #print(SetpointsSheet_compInfo)
        SetpointsSheet_stp_info=pd.read_excel(testdefSheetPath, sheet_name="Setpoints", usecols="F:VZ", keep_default_na=False, skiprows=2)  
        #print(SetpointsSheet_stp_info)
        
        SetpointsDict={}
        stp_id=1
        parameter_row_IDs={} #map parameter names against row numbers in data frame
        row_ID=0
        while (SetpointsSheet_stp_info['Setpoint ID'].iloc[row_ID]!='Symbol'): #in Column 'Setpoint ID' we look for entry symbol. Setpoint ID being the content of this column in the first row that is read makes it the column name
            parameter_row_IDs[SetpointsSheet_stp_info['Setpoint ID'].iloc[row_ID]]=row_ID  # determine keys based on what is in the list. Naming format P_, Q_ and LOC_ must be respected for the power locations. 
            row_ID+=1
        while(stp_id in SetpointsSheet_stp_info.columns):#iterate over columns of setpoints
            SetpointsDict[stp_id]={'comp_settings':{}}
            for parameter in parameter_row_IDs.keys():
                SetpointsDict[stp_id][parameter]=SetpointsSheet_stp_info[stp_id].iloc[parameter_row_IDs[parameter]]
                              
            values=[]
            for value_cnt in range (skip_rows-2, len(SetpointsSheet_stp_info[stp_id]) ): values.append(SetpointsSheet_stp_info[stp_id].iloc[value_cnt])#iterate over PSCAD component settings for given setpoint
            SetpointsDict[stp_id]['comp_settings']['values']=values
            for compSetColName in SetpointsSheet_compInfo.columns:
                comp_param_array=[]
                for cnt in range(0, len(SetpointsSheet_compInfo[compSetColName]) ): comp_param_array.append(SetpointsSheet_compInfo[compSetColName][cnt])
                SetpointsDict[stp_id]['comp_settings'][str(compSetColName)]=comp_param_array
                
            stp_id+=1
                
        #print(SetpointsDict)
        print("Read setpoints definition")
        return_dict['Setpoints']=SetpointsDict
    
    #-------------------------------------------------------------------------
#    #TEST TYPES --. not needed anymore I think
#    testTypesSheet=pd.read_excel(testdefSheetPath, sheet_name="TestTypes", usecols=None, keep_default_na=False)
#    #print(testTypesSheet)
#    TestTypesDict={}
#    for row_nr in range (2,len(testTypesSheet['layers (optional)'])):
#        TestTypesDict[str(testTypesSheet['layers (optional)'].iloc[row_nr])]={}
#        for column_name in testTypesSheet.columns[1:]:
#            TestTypesDict [str(testTypesSheet['layers (optional)'].iloc[row_nr])][str(column_name)]=testTypesSheet[column_name].iloc[row_nr]
#            
#    #print(TestTypesDict)    
#    print("Read Test Types list")
    
    #-------------------------------------------------------------------------
    # SCENARIOS
    if('ScenariosSMIB' in relevant_tabs or relevant_tabs=='all'):  
        # dict with Type+Nr as keys, then 
        ScenariosDict={}
        
        SmallDistSheet=pd.read_excel(testdefSheetPath, sheet_name="SmallDist", usecols=None, keep_default_na=False, skiprows=1)
        for row_cnt in range(0, len(SmallDistSheet)):
            scenario_name='small'+str(SmallDistSheet['CaseNr'].iloc[row_cnt])
            ScenariosDict[scenario_name]={}
            for column_name in SmallDistSheet.columns:
                if(column_name != 'CaseNr'):
                    ScenariosDict[scenario_name][column_name]=SmallDistSheet[column_name].iloc[row_cnt] 
        
        LargeDistSheet=pd.read_excel(testdefSheetPath, sheet_name="LargeDist", usecols=None, keep_default_na=False, skiprows=5)
        general_fault_info=['Test Type', 'run in PSCAD?', 'run in PSS/E?', 'Vpoc', 'Active Power', 'Reactive Power', 'SCR', 'X_R', 'SCL_post', 'X_R_post', 'setpoint ID','TimeStep', 'AccFactor', 'Comment/Corresponding DMAT case', 'test group', 'simulation batch' ]
        prev_scenario_name=''
        for row_cnt in range(0, len(LargeDistSheet)):
            scenario_name='large'+str(LargeDistSheet['CaseNr'].iloc[row_cnt])
            
            if(scenario_name=='large'): #no id defined --> nth fault of multifault series with n>0
                scenario_name=prev_scenario_name
                for column_name in LargeDistSheet.columns:
                    if(column_name != 'CaseNr') and (column_name not in general_fault_info):
                        ScenariosDict[scenario_name][column_name].append(LargeDistSheet[column_name].iloc[row_cnt] )
            
            else: #either first element of Multifault, or other fault test
                ScenariosDict[scenario_name]={}
                # 1st element of Multifault series
                if(str(LargeDistSheet['Test Type'].iloc[row_cnt])=='Multifault'): #this would be the 0th element of a multifault series, in that case initialise as arrays
                    for fault_parameter in general_fault_info:
                        ScenariosDict[scenario_name][fault_parameter]=LargeDistSheet[fault_parameter][row_cnt]
                    # ScenariosDict[scenario_name]['Test Type']='Multifault' #There needs tobe only one entry for Multifault, given the 
                    # ScenariosDict[scenario_name]['run in PSCAD?']= LargeDistSheet['run in PSCAD?'][row_cnt] #only needs to be set once per multifuals series
                    # ScenariosDict[scenario_name]['run in PSS/E?']= LargeDistSheet['run in PSS/E?'][row_cnt]   #only needs tobe set once per multifault series.  
                    # ScenariosDict[scenario_name]['Vpoc']=LargeDistSheet['Vpoc'][row_cnt]
                    # ScenariosDict[scenario_name]['Active Power']=LargeDistSheet['Active Power'][row_cnt]
                    # ScenariosDict[scenario_name]['Reactive Power']=LargeDistSheet['Reactive Power'][row_cnt]
                    # ScenariosDict[scenario_name]['SCR']=LargeDistSheet['SCR'][row_cnt]
                    for column_name in LargeDistSheet.columns:
                        
                        if(column_name != 'CaseNr') and (column_name not in general_fault_info):
                            ScenariosDict[scenario_name][column_name]=[LargeDistSheet[column_name].iloc[row_cnt]]
                
                #other fault
                else:            
                    for column_name in LargeDistSheet.columns:
                        
                        if(column_name != 'CaseNr'):
                            ScenariosDict[scenario_name][column_name]=LargeDistSheet[column_name].iloc[row_cnt]   
                prev_scenario_name=scenario_name
            
        OrtSheet=pd.read_excel(testdefSheetPath, sheet_name="ORT", usecols=None, keep_default_na=False, skiprows=1)
        for row_cnt in range(0, len(OrtSheet)):
            scenario_name='ort'+str(OrtSheet['CaseNr'].iloc[row_cnt])
            ScenariosDict[scenario_name]={}
            for column_name in OrtSheet.columns:
                if(column_name != 'CaseNr'):
                    ScenariosDict[scenario_name][column_name]=OrtSheet[column_name].iloc[row_cnt] 
                
        TovSheet=pd.read_excel(testdefSheetPath, sheet_name="TOV", usecols=None, keep_default_na=False, skiprows=5)
        for row_cnt in range(0, len(TovSheet)):
            scenario_name='tov'+str(TovSheet['CaseNr'].iloc[row_cnt])
            ScenariosDict[scenario_name]={}
            for column_name in TovSheet.columns:
                if(column_name != 'CaseNr'):
                    ScenariosDict[scenario_name][column_name]=TovSheet[column_name].iloc[row_cnt] 
                    
        print("Read Scenarios")
        return_dict['ScenariosSMIB']=ScenariosDict
       
    #-------------------------------------------------------------------------
    if('Profiles' in relevant_tabs or relevant_tabs=='all'):
        # PROFILES
        # The profiles always come as two columns defining one profile --> search for columns where the second entry is not empty and interpret as profile name
        # ProfilesDict={'profile1':{'scaling':X 'x_data':[], 'y_data':[] } }
        # scaling factor 
        #       can be either numerical (absolute value will be: scaling factor x parameter base x Y_value) 
        #       or 'nom' (choosing nominal value of underlying parameter, e.g. base voltag or base frequency as scaling factor), or 'p.u.'
        #       or 'normal' (based on normal value of the underlying parameter, e.g. 1.04 pu voltage) 
        ProfilesDict={}
        ProfilesSheet=pd.read_excel(testdefSheetPath, sheet_name="Profiles", usecols=None, keep_default_na=False, skiprows=0)
        column_names=ProfilesSheet.columns
        for col_id in range(0,len(column_names)):
            if(not 'Unnamed' in column_names[col_id]) and (not 'category' in column_names[col_id] ):
                category = column_names[col_id] #Category always has to be defined in the top row of the first column belonging to a new category
            if(ProfilesSheet[column_names[col_id]].iloc[0] !='') and (ProfilesSheet[column_names[col_id]].iloc[0] != 'profile name'): #valid every time a x_data column is reached
                profile_name=ProfilesSheet[column_names[col_id]].iloc[0]
                ProfilesDict[profile_name]={'scaling':ProfilesSheet[column_names[col_id]].iloc[1], 
                                            'scaling_factor_PSCAD':ProfilesSheet[column_names[col_id]].iloc[2], 
                                            'scaling_factor_PSSE':ProfilesSheet[column_names[col_id+1]].iloc[2], 
                                            'offset_PSCAD':ProfilesSheet[column_names[col_id]].iloc[3], 
                                            'offset_PSSE':ProfilesSheet[column_names[col_id+1]].iloc[3], 
                                            'x_data':[], 'y_data':[]}
                for row_nr in range(5, len(ProfilesSheet[column_names[col_id]])):  
                    x=ProfilesSheet[column_names[col_id]].iloc[row_nr]
                    y=ProfilesSheet[column_names[col_id+1]].iloc[row_nr]
                    if(x!=''):
                        ProfilesDict[profile_name]['x_data'].append(float(x))
                    if(y!=''):
                        ProfilesDict[profile_name]['y_data'].append(float(y))
                            
        #print(ProfilesDict)
        print("Read Test Profiles")
        return_dict['Profiles']=ProfilesDict
    #-------------------------------------------------------------------------
    if('NetworkFaults' in relevant_tabs or relevant_tabs=='all'):
    # CONTINGENCIES
    # dict with Type+Nr as keys, then 
        ContingencyDict={}
        
        ConDistSheet=pd.read_excel(testdefSheetPath, sheet_name="NetworkFaults", usecols=None, keep_default_na=False, skiprows=1)
        for row_cnt in range(0, len(ConDistSheet)):
            CaseNr=str(ConDistSheet['CaseNr'].iloc[row_cnt])
            if CaseNr != '':
                scenario_name='con'+str(CaseNr)
                ContingencyDict[scenario_name]=[{}]
                for column_name in ConDistSheet.columns:
                    if(column_name != 'CaseNr'):
                        ContingencyDict[scenario_name][0][column_name]=ConDistSheet[column_name].iloc[row_cnt]    
            else:
                tempDict={}
                for column_name in ConDistSheet.columns:
                    if(column_name != 'CaseNr'):
                        tempDict[column_name]=ConDistSheet[column_name].iloc[row_cnt]   
                
                ContingencyDict[scenario_name].append(tempDict)
        
        print("Read Contingency Scenarios")
        return_dict['NetworkFaults']=ContingencyDict    
    #-------------------------------------------------------------------------
    if('MonitorBuses' in relevant_tabs or relevant_tabs=='all'):
        #MonitorBuses
        BusLibSheet=pd.read_excel(testdefSheetPath, sheet_name="MonitorBuses", usecols="A,B,C", keep_default_na=False)
        #print(FileSettingsSheet)
        BusLibDict=dict(zip(BusLibSheet.bus_number, BusLibSheet.bus_name, BusLibSheet.bus_code ))
        if('' in BusLibDict.keys()):
            del BusLibDict['']
        #print(FileSettingsDict)
        print("Read Branch Lib")
        return_dict['MonitorBuses']=BusLibDict
    
    #-------------------------------------------------------------------------
    if('MonitorBranches' in relevant_tabs or relevant_tabs=='all'):
        #MonitorBuses
        BranchLibSheet=pd.read_excel(testdefSheetPath, sheet_name="MonitorBranches", usecols="A,B,C,D,E", keep_default_na=False)
        #print(FileSettingsSheet)
        BranchLibDict=dict(zip(BranchLibSheet.brch_from, BranchLibSheet.brch_to, BranchLibSheet.brch_id, BranchLibSheet.brch_name, BranchLibSheet.brch_code))
        if('' in BranchLibDict.keys()):
            del BranchLibDict['']
        #print(FileSettingsDict)
        print("Read Branch Lib")
        return_dict['MonitorBuses']=BranchLibDict
    
    #-------------------------------------------------------------------------
    # RETURN ALL DICTS
    print("Done")
    
    return return_dict
    
    
#TEST purposes
def main():
    readTestdef(r"C:\Users\Mervin Kall\OneDrive - ESCO Pacific\Mulwala\20200622_PSCAD_SMIB_STUDIES\20200622_MUL_TESTINFO.xlsx")

if __name__ == '__main__':
    main()