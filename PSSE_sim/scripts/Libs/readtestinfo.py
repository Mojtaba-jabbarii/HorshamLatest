# -*- coding: utf-8 -*-
"""
Created on Mon Jun 29 09:44:38 2020

@author: Mervin Kall

25/3/2022: Making one section to create an auto file from the Infor file and work on it -> not crash when main file is open

29/3/2022: Adding one section to read the channels from ModelDetailsPSSE for plotting graphs

11/5/2022: modify the network read in for the runback to be included

17/1/2023: udate to read powerderating curve

21/6/2023: Update Network contingency list into network Scenarios list

28/2/2024: Update OutputChannels tab

"""
#making some mods

import pandas as pd
import os
from subprocess import call

def readTestdef(testdefSheetPath, relevant_tabs='all'):

    return_dict={}

    #--------------------------------------------------------------------------
    #MAKE A COPY OF THE FILE
    fileName, fileExt = os.path.splitext(testdefSheetPath) #separate file and extention
    autoFile = fileName + "-AUTO" + fileExt # new file will be created
    if os.path.isfile(autoFile): # If the file exits, remove it before creating a new one.
        os.remove(autoFile)
    copycmd = r"echo F|" + "xcopy /Y /R /K /H /C \"" + testdefSheetPath + "\" \"" + autoFile + "\""
    call(copycmd, shell=True)
    testdefSheetPath = autoFile

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
        SimulationSettingsSheet=pd.read_excel(testdefSheetPath, sheet_name="SimulationSettings", usecols="A,B,C", keep_default_na=False)
        SimulationSettingsDict=dict(zip(SimulationSettingsSheet.ParamIdentifier, SimulationSettingsSheet.Value))
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

        # #Plotting channels
        # ChannelDict={}
        # PSSEmodelSheet=pd.read_excel(testdefSheetPath, sheet_name="ModelDetailsPSSE", usecols="G,H,I,J,K,L,M,N", keep_default_na=False)
        # for row_cnt in range(0, len(PSSEmodelSheet)):
        #     chan_name='chan'+str(PSSEmodelSheet['ChanAdd'].iloc[row_cnt])
        #     ChannelDict[chan_name]={}
        #     for column_name in PSSEmodelSheet.columns:
        #         if(column_name != 'ChanAdd'):
        #             ChannelDict[chan_name][column_name]=PSSEmodelSheet[column_name].iloc[row_cnt] 


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
    if('NetworkScenarios' in relevant_tabs or relevant_tabs=='all'):
    # CONTINGENCIES
    # dict with Type+Nr as keys, then 
        NetworkScenDict={}
        
        ConDistSheet=pd.read_excel(testdefSheetPath, sheet_name="NetworkScenarios", usecols=None, keep_default_na=False, skiprows=1)
        scheme = 0
        for row_cnt in range(0, len(ConDistSheet)):
            CaseNr=str(ConDistSheet['CaseNr'].iloc[row_cnt])
            Runback=str(ConDistSheet['Runback?'].iloc[row_cnt]) 
            Test_Type=str(ConDistSheet['Test Type'].iloc[row_cnt])
            if CaseNr != '':
                scheme = 0
                if float(CaseNr) < 10.0: scenario_name='con0'+str(CaseNr) # add zero in if case number is only one digit
                else: scenario_name='con'+str(CaseNr)
                NetworkScenDict[scenario_name]=[{}]
                for column_name in ConDistSheet.columns:
                    if(column_name != 'CaseNr'):
                        NetworkScenDict[scenario_name][0][column_name]=ConDistSheet[column_name].iloc[row_cnt]    
            else:
                scheme += 1
                tempCaseNr=str(ConDistSheet['CaseNr'].iloc[row_cnt-scheme]) # Name belongs to the previous contingency case
                if (Runback == 'yes') or (Runback == '1') or (Test_Type == 'Multi_fault'): #Check to see if it is a runback or multifault studies
#                    tempCaseNr=str(ConDistSheet['CaseNr'].iloc[row_cnt-scheme]) # Name belongs to the previous contingency case
                    scenario_name='con'+str(tempCaseNr)
                    tempDict={}
                    for column_name in ConDistSheet.columns:
                        if(column_name != 'CaseNr'):
                            tempDict[column_name]=ConDistSheet[column_name].iloc[row_cnt]   
                    
                    NetworkScenDict[scenario_name].append(tempDict)
        
        print("Read Network Study Scenarios")
        return_dict['NetworkScenarios']=NetworkScenDict    


    #-------------------------------------------------------------------------
    if('SteadyStateStudies' in relevant_tabs or relevant_tabs=='all'):
    # CONTINGENCIES - Steady State
        SteadyStateDict={}
        SteadyStateSheet=pd.read_excel(testdefSheetPath, sheet_name="SteadyStateStudies", usecols=None, keep_default_na=False, skiprows=1)
        for row_cnt in range(0, len(SteadyStateSheet)):
            CaseNr=str(SteadyStateSheet['CaseNr'].iloc[row_cnt])
            if CaseNr != '':
                if float(CaseNr) < 10.0: scenario_name='con0'+str(CaseNr) # add zero in if case number is only one digit
                else: scenario_name='con'+str(CaseNr)
                SteadyStateDict[scenario_name]=[{}]
                for column_name in SteadyStateSheet.columns:
                    if(column_name != 'CaseNr'):
                        SteadyStateDict[scenario_name][0][column_name]=SteadyStateSheet[column_name].iloc[row_cnt]    
            else:
                tempDict={}
                for column_name in SteadyStateSheet.columns:
                    if(column_name != 'CaseNr'):
                        tempDict[column_name]=SteadyStateSheet[column_name].iloc[row_cnt]   
                
                SteadyStateDict[scenario_name].append(tempDict)
        
        print("Read Contingency Scenarios - Steady State")
        return_dict['SteadyStateStudies']=SteadyStateDict
        
    #-------------------------------------------------------------------------
    if('MonitorBuses' in relevant_tabs or relevant_tabs=='all'):
#        #MonitorBuses
#        BusLibSheet=pd.read_excel(testdefSheetPath, sheet_name="MonitorBuses", usecols="A,B,C", keep_default_na=False)
#        #print(FileSettingsSheet)
#        BusLibDict=zip(BusLibSheet.bus_number, BusLibSheet.bus_name, BusLibSheet.bus_code )
##        BusLibDict=zip(BusLibSheet.bus_number, BusLibSheet.bus_name)
##        if('' in BusLibDict.keys()):
##            del BusLibDict['']
#        #print(FileSettingsDict)
#        print("Read MonitorBuses")
#        return_dict['MonitorBuses']=BusLibDict

        BusLibSheet=pd.read_excel(testdefSheetPath, sheet_name="MonitorBuses", usecols="A,B,C", keep_default_na=False)
        bus_numbers = BusLibSheet.bus_number.to_list()
        bus_numbers = [int(x) for x in bus_numbers]
        bus_names =BusLibSheet.bus_name.to_list()
#        print(bus_numbers)
        BusLibDict = [bus_numbers, bus_names]
        print("Read MonitorBuses")
        return_dict['MonitorBuses']=BusLibDict

    #-------------------------------------------------------------------------
    if('MonitorBranches' in relevant_tabs or relevant_tabs=='all'):
#        #MonitorBranches
#        BranchLibSheet=pd.read_excel(testdefSheetPath, sheet_name="MonitorBranches", usecols="A,B,C,D,E", keep_default_na=False)
#        #print(FileSettingsSheet)
#        BranchLibDict=zip(BranchLibSheet.brch_from, BranchLibSheet.brch_to, BranchLibSheet.brch_id, BranchLibSheet.brch_name, BranchLibSheet.brch_code)
##        BranchLibDict=zip(BranchLibSheet.brch_from, BranchLibSheet.brch_to, BranchLibSheet.brch_id, BranchLibSheet.brch_name)
##        if('' in BranchLibDict.keys()):
##            del BranchLibDict['']
#        #print(FileSettingsDict)
#        print("Read MonitorBranches")
#        return_dict['MonitorBranches']=BranchLibDict


        BranchLibSheet=pd.read_excel(testdefSheetPath, sheet_name="MonitorBranches", usecols="A,B,C,D,E", keep_default_na=False)
        from_buses = BranchLibSheet.brch_from.to_list()
        from_buses = [int(x) for x in from_buses]
        to_buses = BranchLibSheet.brch_to.to_list()
        to_buses = [int(x) for x in to_buses]
        brch_ids = BranchLibSheet.brch_id.to_list()
        brch_ids = [int(x) for x in brch_ids]
        brch_names = BranchLibSheet.brch_name.to_list()
        BranchLibDict = [from_buses, to_buses, brch_ids, brch_names]
        print("Read MonitorBranches")
        return_dict['MonitorBranches']=BranchLibDict

    #-------------------------------------------------------------------------
    if('PowerCapability' in relevant_tabs or relevant_tabs=='all'):
        # PROFILES
        # The profiles always come as two columns defining one profile --> search for columns where the second entry is not empty and interpret as profile name
        # ProfilesDict={'profile1':{'scaling':X 'x_data':[], 'y_data':[] } }
        # scaling factor 
        #       can be either numerical (absolute value will be: scaling factor x parameter base x Y_value) 
        #       or 'nom' (choosing nominal value of underlying parameter, e.g. base voltag or base frequency as scaling factor), or 'p.u.'
        #       or 'normal' (based on normal value of the underlying parameter, e.g. 1.04 pu voltage) 
        ProfilesDict={}
        ProfilesSheet=pd.read_excel(testdefSheetPath, sheet_name="PowerCapability", usecols=None, keep_default_na=False, skiprows=0)
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
        print("Read Test PowerCapability")
        return_dict['PowerCapability']=ProfilesDict

    if('OutputChannels' in relevant_tabs or relevant_tabs=='all'):
        OutChansDict={}
        OutputChannelsSheet=pd.read_excel(testdefSheetPath, sheet_name="OutputChannels", usecols="A:I", keep_default_na=False, skiprows=1)
        for row_cnt in range(0, len(OutputChannelsSheet)):
            CaseNr=str(OutputChannelsSheet['ChanNum'].iloc[row_cnt])
            if float(CaseNr) < 10.0: scenario_name='ChanNum0'+str(CaseNr) # add zero in if case number is only one digit
            else: scenario_name='ChanNum'+str(CaseNr)
            
            OutChansDict[scenario_name]={}
            for column_name in OutputChannelsSheet.columns:
                if(column_name != 'ChanNum'):
                    OutChansDict[scenario_name][column_name]=OutputChannelsSheet[column_name].iloc[row_cnt] 
                    

#        CaseNr=str(OutputChannelsSheet['Instance'].iloc[row_cnt])
#        Location = OutputChannelsSheet.Location.to_list()
#        Type = OutputChannelsSheet.Type.to_list()
#        Position = OutputChannelsSheet.Position.to_list()
#        Position = [str(x) for x in Position]
#        Name = OutputChannelsSheet.Name.to_list()
#        Legend = OutputChannelsSheet.Legend.to_list()
#        OutChansDict = [Location, Type, Position, Name, Legend]
#        OutChansDict[scenario_name].append(tempDict)
        print("Read OutputChannels")
        return_dict['OutputChannels']=OutChansDict
    #-------------------------------------------------------------------------
    # RETURN ALL DICTS
    print("Done dicts readin")
    
    return return_dict
    
    
#TEST purposes
def main():
    readTestdef(r"C:\Users\Mervin Kall\Documents\GitHub\PowerSystemStudyTool\20220203_APE\test_scenario_definitions\20220310_MUS_TESTINFO_V2b.xlsx", ['Setpoints'])

if __name__ == '__main__':
    main()