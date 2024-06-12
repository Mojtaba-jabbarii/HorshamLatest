"""
ESCOPlot
=====

Patrick Rossiter 2017

14/11/2022: Update GSMG_bands: Plot the error band following PSMG 
15/11/2022: Merge with functions from LSF project                                                                  
26/6/2023: update number of graphs greater than 10; include callout option in markers -> datapoint at 1.5 and 5.5secs
21/7/2023: update the denote for settling time to be consistent with calculated results in csv file 
"""
import sys, os
import numpy as np
import io
from io import StringIO

class ESCOPlot(object):
    def __init__(self):
        self.plotspec = [[] for _ in range(12)]
        self.timearrays = []
        self.dataarrays = []
        self.plotdataarrays = []
        self.filenames = []
        self.offsets = []
        self.timeoffset = []
        self.scales = []
        self.GSMG_arrays = []
        self.settle_arrays = []
        self.settleTimeArrays=[] #included to store settling time data to include time markers in plots and visualise settling time and rise time in plots. Allows to store values for multiple steps per test.
        self.riseTimeArrays=[] #same as above --> mainly to be used for reactive current rise time on fault applicationa as well as reactive power rise time during S5.2.5.13
        self.recTimeArrays=[] #same as above, but for recovery time, mainly used for active power recovery following fault(s)
        self.dVdIq=[] #array to save information about dIq/dV injection during voltage disturbances.
        self.tolerance_band_offset=[]
        self.tolerance_band_base=[]
        self.callout=[]
        self.xlimit = [] #
        self.intervals = [] #this contains information on the relvant index ranges for each data file. It is populated in the plot routine, based on the data in self.xlimit
        self.ylimit = [[] for _ in range(12)]
        self.y2limit = [[] for _ in range(12)]
        self.yspan = np.zeros(12)
        self.y2span = np.zeros(12)
        self.ymaxlim = [99999 for _ in range(12)]
        self.yminlim = [-99999 for _ in range(12)]
        self.channel_names = []
        self.xlabels = [[] for _ in range(12)]
        self.ylabels = [[] for _ in range(12)]
        self.y2labels =[[] for _ in range(12)]
        self.titles = [[] for _ in range(12)]
        self.legends = [[] for _ in range(12)]
    
    def listEntries(self):
        '''
        Lists the entry number and file name of each .csv and .out file which 
        has been read into the class. Useful for looking up which entry applies
        to which file
        '''
        print ("Files in memory are:")
        print ("Entry  Channels   Filename")
        for index, entry in enumerate(self.filenames):
            print ("{:^5} {:^10}  {:<260}".format(index, (len(self.dataarrays[index])), entry))
    
    def read_data(self, filename, pscad_files = 999, timeID='time'):
        '''
        Read data into the method's arrays
        
        Specified files can be either PSCAD output files (.inf) or PSSE output 
        files (.out).
        
        Requires numpy and PSSE v34 for PSSE output files.
        
        '''
        self.timeoffset.append(0.0)
        if filename[-4:] in ('.out', '.OUT'):
            self.filenames.append(filename)
            # PSSE output files
            #########################
            sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
            sys.path.append(sys_path_PSSE)
            os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
            os.environ['PATH'] += ';' + os_path_PSSE
            os.environ['PATH'] += ';' + sys_path_PSSE
            #########################
            import psse34
            import dyntools
            outfile_data = dyntools.CHNF(filename)
            short_title, chanid, chandata = outfile_data.get_data()
    
            time = np.array(chandata['time'])
            data = np.zeros((len(chandata) - 1, len(time)))
            
            chan_ids = []
            print (filename)
            for key in chandata.keys()[:-1]:
                print (key, chanid[key])
                data[key - 1] = chandata[key]
                chan_ids.append(chanid[key])
            print ('\n')
            
            self.timearrays.append(time)
            self.dataarrays.append(data)
            self.scales.append(np.ones(len(data)+1))
            self.GSMG_arrays.append(np.zeros(len(data)+1))
            self.settle_arrays.append([[] for i in range(len(chan_ids)+1)])
            self.offsets.append(np.zeros(len(data)+1))
            self.tolerance_band_offset.append(np.zeros(len(data)+1))
            self.tolerance_band_base.append(np.zeros(len(data)+1))
            self.channel_names.append(chan_ids)

        if filename[-4:] in ('.csv'):
            import pandas as pd
            import glob
            csv_data = pd.DataFrame()
            csv_time = pd.DataFrame()
            
            d_in = pd.read_csv(filename, delimiter = ',',
                           skip_blank_lines = True, header = 0,
                           skiprows = None)
#            d_in = pd.read_csv(filename, delimiter = ',',
#                           skip_blank_lines = True, header = None,
#                           skiprows = [0])
            
            csv_data_indices = d_in.columns 
            
            csv_time = d_in[timeID]
            
            d_in.columns=range(0, len(csv_data_indices))
            
            csv_data = d_in.T

            #csv_data_indices = d_in.columns 
            #csv_data_indices = []

            print (filename)
            for idx, chanid in enumerate(csv_data_indices):
                print (idx+1, chanid)
            print ('\n')
            
            # Add data to arrays in class
            self.timearrays.append(csv_time.values)
            self.dataarrays.append(csv_data.values)
            self.scales.append(np.ones(len(csv_data.values)+1))
            self.GSMG_arrays.append(np.zeros(len(csv_data.values)+1))
#            self.settle_arrays.append([[] for i in range(len(csv_data_indices))])
            self.offsets.append(np.zeros(len(csv_data.values)+1))
            self.tolerance_band_offset.append(np.zeros(len(csv_data.values)+1))
            self.tolerance_band_base.append(np.zeros(len(csv_data.values)+1))
            self.channel_names.append(csv_data_indices)

            self.settle_arrays.append([[] for i in range(len(csv_data) + 1)])
            
            self.settleTimeArrays.append([[] for i in range(len(csv_data)+1)]) #into every empty array cell, a sequence of tuple can be added, each representing start time and end time of a transient event
            self.riseTimeArrays.append([[] for i in range(len(csv_data)+1)])
            self.recTimeArrays.append([[] for i in range(len(csv_data)+1)])
            self.dVdIq.append([{}for i in range(len(csv_data)+1)])
            self.callout.append([{}for i in range(len(csv_data)+1)])
            pass


        if filename[-4:] in ('.inf'):
            import pandas as pd
            import glob
            PSCAD_data = pd.DataFrame()
            PSCAD_time = pd.DataFrame()
            
            if pscad_files < 999: endfile = pscad_files
            else: endfile = pscad_files

            for idx, outfile in enumerate(glob.glob(filename[:-4] + '*.out')[:endfile]):
                # print outfile
                d_in = pd.read_csv(outfile, delim_whitespace = True, 
                                   skip_blank_lines = True, header = None,
                                   skiprows = [0])
                PSCAD_data = pd.concat([PSCAD_data, d_in.T[1:]])
                if not idx:
                    PSCAD_time = d_in.T.values[0]

            PSCAD_data_indices = []
            f = open(filename, 'r')
            for line in f.readlines():
                startindex = line.index('Desc=') + 6
                endindex = line[startindex:].index('"') + startindex
                PSCAD_data_indices.append(line[startindex:endindex])
            #
            # print filename
            for idx, chanid in enumerate(PSCAD_data_indices):
                print (idx+1, chanid)
            print ('\n')
            
            '''
            Old news!! Pandas now working.
            # PSCAD output files
            import glob, re
            if pscad_files == 999:
                pscad_files = len(glob.glob(filename[:-4] + '*.out'))
            #
            outfiles = []
            for n in range(1, pscad_files+0):
                outfiles.append(glob.glob(filename[:-4] + '_*%1.2i*.out' % n)[0])
            #
            chan_ids = []
            f = open(filename, 'r')
            lines = f.readlines()
            f.close()
            for idx, rawline in enumerate(lines):
                line = re.split(r'\s{2,}', rawline)
                chan_ids.append(line[2][6:-1])
            #
            # Setup time array using last outfile
            time = []
            f = open(outfiles[-1], 'r')
            lines = f.readlines()
            f.close()
            for l_idx, rawline in enumerate(lines[1:]):
                line = re.split(r'\s{1,}', rawline)
                try:
                    time.append(float(line[1]))
                except:
                    pass
            time = np.array(time)
            #
            # Setup empty data array using known number of channels and time steps
            data = np.zeros((len(chan_ids), len(time)))
            #
            # Populate data array using all outfiles
            for f_idx, outfile in enumerate(outfiles):
                f = open(outfile, 'r')
                lines = f.readlines()
                f.close()
                for n, rawline in enumerate(lines[1:]):
                    line = re.split(r'\s{1,}', rawline)
                    for m, value in enumerate(line[2:-1]):
                        try:
                            data[(f_idx*10) + m][n] = value
                        except:
                            pass
            for n, chan_id in enumerate(chan_ids):
                print n+1, chan_id
            '''
            
            # Add data to arrays in class
            self.timearrays.append(PSCAD_time)
            self.dataarrays.append(PSCAD_data.values)
            self.scales.append(np.ones(len(PSCAD_data.values)+1))
            self.GSMG_arrays.append(np.zeros(len(PSCAD_data.values)+1))
            self.settle_arrays.append([[] for i in range(len(PSCAD_data_indices)+1)])
            #new added data vectors
            self.settleTimeArrays.append([[] for i in range(len(PSCAD_data_indices)+1)]) #into every empty array cell, a sequence of tuple can be added, each representing start time and end time of a transient event
            self.riseTimeArrays.append([[] for i in range(len(PSCAD_data_indices)+1)])
            self.recTimeArrays.append([[] for i in range(len(PSCAD_data_indices)+1)])
            self.dVdIq.append([{}for i in range(len(csv_data)+1)])
            self.callout.append([{}for i in range(len(csv_data)+1)])
            #
            self.offsets.append(np.zeros(len(PSCAD_data.values)+1))
            self.tolerance_band_offset.append(np.zeros(len(PSCAD_data.values)+1))
            self.tolerance_band_offset.append(np.zeros(len(csv_data.values)+1))
            self.tolerance_band_base.append(np.zeros(len(PSCAD_data.values)+1))
            self.channel_names.append(PSCAD_data_indices)
            
            '''
            self.filenames.append(filename)
            # PSCAD output files
            import glob, re
            #
            outfiles = glob.glob(filename[:-4] + '*.out')
            #
            chan_ids = []
            f = open(outfiles[-1], 'r')
            lines = f.readlines()
            f.close()
            for idx, rawline in enumerate(lines[1:]):
                line = re.split(r'\s{1,}', rawline)
                chan_ids.append(line[2][6:-1])
            #
            # Setup time array using last outfile
            time = []
            f = open(outfiles[-1], 'r')
            lines = f.readlines()
            f.close()
            for l_idx, rawline in enumerate(lines[1:]):
                line = re.split(r'\s{1,}', rawline)
                time.append(float(line[1]))
            time = np.array(time)
            #
            # Setup empty data array using known number of channels and time steps
            print len(chan_ids), len(time)
            data = np.zeros((len(chan_ids), len(time) + 1))
            #
            # Populate data array using all outfiles
            for f_idx, outfile in enumerate(outfiles):
                f = open(outfile, 'r')
                lines = f.readlines()
                f.close()
                n = 0
                for rawline in lines[1:]:
                    line = re.split(r'\s{1,}', rawline)
                    for m, value in enumerate(line[2:-1]):
                        try:
                            data[(f_idx*10) + m][n] = value
                        except:
                            print f_idx, m, n, value
                            print data
                            raise NameError('FUCK')
                    n += 1
            #
            print filename
            for idx, chanid in enumerate(chan_ids):
                print idx+1, chanid
            print '\n'
            # Add data to arrays in class
            self.timearrays.append(time)
            self.dataarrays.append(data)
            self.scales.append(np.ones(len(data)+1))
            self.GSMG_arrays.append(np.zeros(len(data)+1))
            self.settle_arrays.append([[] for i in range(len(chan_ids))])
            self.offsets.append(np.zeros(len(data)+1))
            self.channel_names.append(chan_ids)
            '''
            
    def calcCurrents(self, entry, P=-1, Q=-1, V=-1, I=-1, nameLabel='default', scaling=1):
        import pandas as pd
        """routine can be used to calculate reactive current and active current from either
                Active Power, Reactive Power and Voltage OR from active Power, reactive power and apparent current. 
                for the input signals to be used, the columns IDs must  be provided.
                If Voltage (V) is provided, but apparent current (I) is not provided, the active current vector will be calculated from P and V, if P is provided. The reactive current vector will be provided if Q is specified.
                
                In case the voltage is not specified and, apparent current (I) and P and Q must be specified, otherwise the function cannot calculate I
                "name" is interpreted as suffix for the new current vectors that are calculated. The generated vecotr will be named:
                        Iq_name and Ip_name
                        
                The function will return a dictionary, consisting of {'Iq':[Iq_name, channelID], 'Ip':[Ip_name, channelID] }
                in case Iq and/or Ip cannot be generated due to insufficient input, the entry in the dict will be an empty array, but the keys will still be included
                
        """
        #retreive channel ids in case they are specified as strings
        if type(P) is str:
            chars = len(P)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == P:
                    P = idx + 1  # Channel numbers start at 1
                    break
        if type(Q) is str:
            chars = len(Q)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == Q:
                    Q = idx + 1  # Channel numbers start at 1
                    break
        if type(V) is str:
            chars = len(V)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == V:
                    V = idx + 1  # Channel numbers start at 1
                    break
        if type(I) is str:
            chars = len(I)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == I:
                    I = idx + 1  # Channel numbers start at 1
                    break
                
        if (I!=-1 and Q!=-1 and P!=-1): #Calculate Iq and Ip, if all three vectors are provided. Results will be provided on same basis as I and assumign that P and Q are expressed in the same unit. 
                                        #Any deviation from that needs to be compensated by the scaling factor. 
            pass
            I_vec=self.dataarrays[entry][I-1]
            P_vec=self.dataarrays[entry][P-1]
            Q_vec=self.dataarrays[entry][Q-1]
            ones_vec=np.ones(len(I_vec))            
            
            Q_vec[abs(Q_vec)<0.0000001]=0.0000001
            P_vec[abs(P_vec)<0.0000001]=0.0000001 #avoid possible division by 0 by setting elements 
            #calculate Iq
            Iq=scaling*np.power( (np.power(I_vec, 2)/(ones_vec+(np.power(P_vec/Q_vec, 2))) ), 0.5)*np.sign(Q_vec)
            #calcualte Ip
            Ip=scaling*np.power( (np.power(I_vec, 2)/(ones_vec+(np.power(Q_vec/P_vec, 2))) ), 0.5)*np.sign(P_vec)
            new_columns=pd.Index(['Iq_'+nameLabel, 'Ip_'+nameLabel])
            self.channel_names[entry]=self.channel_names[entry].append(new_columns) 
            self.dataarrays[entry]=np.append(self.dataarrays[entry], np.array([Iq]) , axis=0) #add the Iq and Ip column
            self.dataarrays[entry]=np.append(self.dataarrays[entry], np.array([Ip]) , axis=0) #add the Iq and Ip column
            self.scales[entry]=np.append(self.scales[entry], [1.0, 1.0]) #include 1 as scaling factor. This will later be specified when defining a plot for the vector with the "subplot_spec" routine
            self.GSMG_arrays[entry]=np.append(self.GSMG_arrays[entry], [0.0, 0.0])
            #self.settle_arrays[entry]=np.append(self.settle_arrays[entry], [[],[]])
            self.settle_arrays[entry].append( [] )     
            self.settle_arrays[entry].append( [] )  
            #self.settleTimeArrays=np.append(self.settleTimeArrays[entry], [[],[]])
            self.settleTimeArrays[entry].append([])
            self.settleTimeArrays[entry].append([])
            #self.riseTimeArrays=np.append(self.riseTimeArrays[entry], [[],[]])
            self.riseTimeArrays[entry].append([])
            self.riseTimeArrays[entry].append([])
            #self.recTimeArrays=np.append(self.recTimeArrays[entry], [[],[]])
            self.recTimeArrays[entry].append([])
            self.recTimeArrays[entry].append([])
            self.dVdIq[entry].append({})
            self.dVdIq[entry].append({})
            self.callout[entry].append({})
            self.callout[entry].append({})
            
            self.offsets[entry]=np.append(self.offsets[entry], [0.0, 0.0])
            self.tolerance_band_offset[entry]=np.append(self.tolerance_band_offset[entry], [0.0, 0.0])
            self.tolerance_band_base[entry]=np.append(self.tolerance_band_base[entry], [0.0, 0.0])
            #generate return dict
            chan_id=len(self.channel_names[entry])
            outcome={'Iq':['Iq_'+nameLabel, chan_id-1], 'Ip':['Ip_'+nameLabel, chan_id]}
            
        elif(V!=-1):
            V_vec=self.dataarrays[entry][V-1]
            V_vec[abs(V_vec)<0.0000001]=0.0000001 #avoid division by 0, in case voltage vector reduces to 0.
            #initialise return dict
            outcome={'Iq':[],'Ip':[], 'Itot':[]}
            if(Q!=-1):
                #calculate Iq
                Q_vec=self.dataarrays[entry][Q-1]
                Iq=scaling*Q_vec/V_vec
                
                new_column=pd.Index(['Iq_'+nameLabel])
                self.channel_names[entry]=self.channel_names[entry].append(new_column) 
                self.dataarrays[entry]=np.append(self.dataarrays[entry], np.array([Iq]), axis=0) #add the Iq and Ip column
                self.scales[entry]=np.append(self.scales[entry], [1.0]) #include 1 as scaling factor. This will later be specified when defining a plot for the vector with the "subplot_spec" routine
                self.GSMG_arrays[entry]=np.append(self.GSMG_arrays[entry], [0.0])
                #self.settle_arrays[entry]=np.append(self.settle_arrays[entry], [])
                self.settle_arrays[entry].append( [] )  
                #self.settleTimeArrays=np.append(self.settleTimeArrays[entry], [])
                self.settleTimeArrays[entry].append([])
                #self.riseTimeArrays=np.append(self.riseTimeArrays[entry], [])
                self.riseTimeArrays[entry].append([])
                #self.recTimeArrays=np.append(self.recTimeArrays[entry], [])
                self.recTimeArrays[entry].append([])
                self.dVdIq[entry].append({})
                self.callout[entry].append({})
                self.offsets[entry]=np.append(self.offsets[entry], [0.0])
                self.tolerance_band_offset[entry]=np.append(self.tolerance_band_offset[entry], [0.0])
                self.tolerance_band_base[entry]=np.append(self.tolerance_band_base[entry], [0.0])
                #add entry to return dict
                chan_id=len(self.channel_names[entry])
                outcome['Iq']=['Iq_'+nameLabel, chan_id] #channel is not actual array ID but, first channel at position 0 is actually called '1'
            if(P!=-1):
                #calculate Ip
                P_vec=self.dataarrays[entry][P-1]
                Ip=scaling*P_vec/V_vec
                
                new_column=pd.Index(['Ip_'+nameLabel])
                self.channel_names[entry]=self.channel_names[entry].append(new_column) 
                self.dataarrays[entry]=np.append(self.dataarrays[entry], np.array([Ip]), axis=0) #add the Iq and Ip column
                self.scales[entry]=np.append(self.scales[entry], [1.0]) #include 1 as scaling factor. This will later be specified when defining a plot for the vector with the "subplot_spec" routine
                self.GSMG_arrays[entry]=np.append(self.GSMG_arrays[entry], [0.0])
                #self.settle_arrays[entry]=np.append(self.settle_arrays[entry], [])
                self.settle_arrays[entry].append( [] )  
                #self.settleTimeArrays=np.append(self.settleTimeArrays[entry], [])
                self.settleTimeArrays[entry].append([])
                #self.riseTimeArrays=np.append(self.riseTimeArrays[entry], [])
                self.riseTimeArrays[entry].append([])
                #self.recTimeArrays=np.append(self.recTimeArrays[entry], [])
                self.recTimeArrays[entry].append([])
                self.dVdIq[entry].append({})
                self.callout[entry].append({})
                self.offsets[entry]=np.append(self.offsets[entry], [0.0])
                self.tolerance_band_offset[entry]=np.append(self.tolerance_band_offset[entry], [0.0])
                self.tolerance_band_base[entry]=np.append(self.tolerance_band_base[entry], [0.0])
                #add entry to return dict
                chan_id=len(self.channel_names[entry])
                outcome['Ip']=['Ip_'+nameLabel, chan_id]
            if(P!=-1 and Q!=-1):
                #calculate apaprent current
                Itot=np.power( (np.power(Ip, 2)+np.power(Iq,2)),0.5 )
                new_column=pd.Index(['Itot_'+nameLabel])
                self.channel_names[entry]=self.channel_names[entry].append(new_column) 
                self.dataarrays[entry]=np.append(self.dataarrays[entry], np.array([Itot]), axis=0) #add the Iq and Ip column
                self.scales[entry]=np.append(self.scales[entry], [1.0]) #include 1 as scaling factor. This will later be specified when defining a plot for the vector with the "subplot_spec" routine
                self.GSMG_arrays[entry]=np.append(self.GSMG_arrays[entry], [0.0])
                #self.settle_arrays[entry]=np.append(self.settle_arrays[entry], [])
                self.settle_arrays[entry].append( [] )  
                #self.settleTimeArrays=np.append(self.settleTimeArrays[entry], [])
                self.settleTimeArrays[entry].append([])
                #self.riseTimeArrays=np.append(self.riseTimeArrays[entry], [])
                self.riseTimeArrays[entry].append([])
                #self.recTimeArrays=np.append(self.recTimeArrays[entry], [])
                self.recTimeArrays[entry].append([])
                self.dVdIq[entry].append({})
                self.callout[entry].append({})
                self.offsets[entry]=np.append(self.offsets[entry], [0.0])
                self.tolerance_band_offset[entry]=np.append(self.tolerance_band_offset[entry], [0.0])
                self.tolerance_band_base[entry]=np.append(self.tolerance_band_base[entry], [0.0])
                #add entry to return dict
                chan_id=len(self.channel_names[entry])
                outcome['Itot']=['Itot_'+nameLabel, chan_id]
                
        return outcome
            
    def calPFs(self, entry, P=-1, Q=-1, nameLabel='default', scaling=1):
        import pandas as pd
        """routine can be used to calculate reactive current and active current from either
        """
        #retreive channel ids in case they are specified as strings
        if type(P) is str:
            chars = len(P)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == P:
                    P = idx + 1  # Channel numbers start at 1
                    break
        if type(Q) is str:
            chars = len(Q)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == Q:
                    Q = idx + 1  # Channel numbers start at 1
                    break
 

        outcome={'PF':[]}
        if(Q!=-1 and P!=-1):
            #calculate Iq
            Q_vec=self.dataarrays[entry][Q-1]
            P_vec=self.dataarrays[entry][P-1]
#            PF = (abs(P_vec)/np.power( (np.power(P_vec, 2)+np.power(Q_vec,2)),0.5 ))
#            S_vec=np.power( (np.power(P_vec, 2)+np.power(Q_vec,2)),0.5 )
#            PF = (abs(P_vec)/S_vec)
#            for v in range(len(PF)):
#                if Q_vec[v] >0: # note that Q at POC is measured flowing back to the power plant
#                    PF[v] = 2.0-PF[v]
            
##            Q_vec_sign = [1.0 if v >= 0 else -1.0 for v in Q_vec]
#            Q_vec_sign = 2.0*(Q_vec >= 0) - 1.0
#            PF=np.cos(np.arctan(scaling*Q_vec/P_vec))*Q_vec_sign
#            for v in range(len(PF)):
#                if Q_vec[v] >0: # note that Q at POC is measured flowing back to the power plant
#                    PF[v] = 2.0+PF[v] # for ploting purpose: if pf is negative, then convert it to positive above 1.0
                    
#            PF=np.cos(np.arctan(scaling*Q_vec/P_vec))
#            for v in range(len(PF)):
#                if Q_vec[v] <0: # note that Q at POC is measured flowing back to the power plant
#                    PF[v] = 2.0-PF[v] # for ploting purpose: if pf is negative, then convert it to positive above 1.0

            PF=np.cos(np.arctan(abs(Q_vec/P_vec))) #absolute power factor value
            for v in range(len(PF)):
                if Q_vec[v] <-0.5: # PF to be with the same sign of Q. e.g.if Q is negative, then PF is negative. Factor 0.1 represent the deadband
                    PF[v] = -scaling*PF[v] # take into acount the direction of Q measurement using scaling factor
                else:
                    PF[v] = scaling*PF[v]
                    
            new_column=pd.Index(['PF_'+nameLabel])
            self.channel_names[entry]=self.channel_names[entry].append(new_column) 
            self.dataarrays[entry]=np.append(self.dataarrays[entry], np.array([PF]), axis=0) #add the PF column
            self.scales[entry]=np.append(self.scales[entry], [1.0]) #include 1 as scaling factor. This will later be specified when defining a plot for the vector with the "subplot_spec" routine
            self.GSMG_arrays[entry]=np.append(self.GSMG_arrays[entry], [0.0])
            #self.settle_arrays[entry]=np.append(self.settle_arrays[entry], [])
            self.settle_arrays[entry].append( [] )  
            #self.settleTimeArrays=np.append(self.settleTimeArrays[entry], [])
            self.settleTimeArrays[entry].append([])
            #self.riseTimeArrays=np.append(self.riseTimeArrays[entry], [])
            self.riseTimeArrays[entry].append([])
            #self.recTimeArrays=np.append(self.recTimeArrays[entry], [])
            self.recTimeArrays[entry].append([])
            self.dVdIq[entry].append({})
            self.callout[entry].append({})
            self.offsets[entry]=np.append(self.offsets[entry], [0.0])
            self.tolerance_band_offset[entry]=np.append(self.tolerance_band_offset[entry], [0.0])
            self.tolerance_band_base[entry]=np.append(self.tolerance_band_base[entry], [0.0])
            #add entry to return dict
            chan_id=len(self.channel_names[entry])
            outcome['PF']=['PF_'+nameLabel, chan_id] #channel is not actual array ID but, first channel at position 0 is actually called '1'
           
        return outcome                
    
    def clear_subplot_spec(self):    
        self.plotspec = [[] for _ in range(12)]
        
    def clear_ylabels(self):
        self.ylabels = [[] for _ in range(12)]
        self.y2labels = [[] for _ in range(12)]
        
    def clear_ylimits(self):
        self.ylimit = [[] for _ in range(12)]
        self.y2limit = [[] for _ in range(12)]
        self.ymaxlim = [99999 for _ in range(12)]
        self.yminlim = [-99999 for _ in range(12)]
        
    def clear_yspan(self):
        self.yspan = np.zeros(12)
        self.y2span = np.zeros(12)


#    def qrise(self, entry, channel, starttime = 0.0, endtime = 10.0):
#        if type(channel) is str:
#            chars = len(channel)
#            for idx, name in enumerate(self.channel_names[entry]):
#                if name == channel:
#                    channel = idx + 1  # Channel numbers start at 1
#                    break
#        time = self.timearrays[entry] + self.timeoffset[entry]
#        # Determine location in array of starttime, as it's easier to use array location than actual time in seconds.
#        measstart = np.argmin(abs(starttime - time))
#        measend = np.argmin(abs(endtime - time))
#        
#        if(measstart<measend):
#            data = self.dataarrays[entry][channel-1][measstart:measend]
#            
#            startq = data[0]
#            endq = data[-1]
#            qrange = endq - startq
#            
#            # print startq, endq, qrange
#            
#            qrise_start = startq + (0.1 * qrange)
#            qrise_end = endq - (0.1 * qrange)
#            
#            # print qrise_start, qrise_end
#            #determine position where signal peaks. this This is the latest point in time that should be considered for the rise time 90% mark.  
#            
#            #qrise_start_idx = np.argmin(abs(qrise_start - data[0:np.argmax(data)])) #checks at which position the quantity is closest to the value identified as start value. Problem is if it goes pas the point twice. Define it as threshold instead. 
#            #qrise_end_idx = np.argmin(abs(qrise_end - data[0:np.argmax(data)]))
#            # print time[qrise_end_idx], time[qrise_start_idx]
#            if(np.sign(qrange)<0):
#                data_end_idx=np.argmin(data)
#            else:
#                data_end_idx=np.argmax(data)
#            error_qrise_start=np.sign(qrange)*(data[0:data_end_idx]-qrise_start)
#            qrise_start_idx =np.where(error_qrise_start>0, error_qrise_start, np.inf).argmin()
#            error_qrise_end=np.sign(qrange)*(qrise_end-data[0:data_end_idx]) #calcualte error vector
#            qrise_end_idx =np.where(error_qrise_end>0, error_qrise_end, np.inf).argmin()#make sure when rounding because of limited time resolution, we are not rounding to our disadvantage.
#            print 'Reactive power rise time = {:1.4f} seconds'.format(abs(time[qrise_end_idx] - time[qrise_start_idx]))
#            if(qrise_start_idx==qrise_end_idx):
#                qrise_end_idx+=1 #rise time is never 0. If the result is 0 it will be more accurate to return the time between two successive data points as a result instead.
#            #save rise time in data vectors
#            self.riseTimeArrays[entry][channel-1].append([time[qrise_start_idx+measstart], time[qrise_end_idx+measstart]]) #if no previous rise time has been calculated for that vector, the array is appended to the existing empty array. By adding measstart as offset, the actual time at which the marker should be present in the final plot is stored in the metadata
#            return abs(time[qrise_end_idx] - time[qrise_start_idx])
#
#        else:
#            return 0
            
    def qrise(self, entry, channel, starttime = 0.0, endtime = 10.0):
        if type(channel) is str:
            chars = len(channel)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == channel:
                    channel = idx + 1  # Channel numbers start at 1
                    break
        time = self.timearrays[entry] + self.timeoffset[entry]
        timestep = time[1] - time[0]
        if timestep == 0: timestep = time[2] - time[1]
        # Determine location in array of starttime, as it's easier to use array location than actual time in seconds.
        measstart = np.argmin(abs(starttime - time))
        measend = np.argmin(abs(endtime - time))
        if(measstart<measend):
            data = self.dataarrays[entry][channel-1][measstart:measend]
            time_ext = time[measstart:measend]
            startq = data[0]
            endq = data[-1]
            qrange = endq - startq
            qrise_start = startq + (0.1 * qrange)
            qrise_end = endq - (0.1 * qrange)        
            qrise_start_time = 0.0
            qrise_end_time = 0.0
            # Positive step             
            if startq < endq:            
                for n in range(len(data) - 1):
                    if (data[n] < qrise_start) and (data[n+1] > qrise_start):
                        qrise_start_time = time_ext[n+1]
                for n in range(len(data) - 1):
                    if (data[n] < qrise_end) and (data[n+1] > qrise_end):
                        qrise_end_time = time_ext[n]
                        break
            # Negative step 
            elif startq > endq:            
                for n in range(len(data) - 1):
                    if (data[n] > qrise_start) and (data[n+1] < qrise_start):
                        qrise_start_time = time_ext[n+1]
             
                for n in range(len(data) - 1):
                    if (data[n] > qrise_end) and (data[n+1] < qrise_end):
                        qrise_end_time = time_ext[n] 
                        break
            # print qrise_start, qrise_end
            # print qrise_start_time, qrise_end_time
#            print 'Reactive power rise time = {:1.4f} seconds'.format(abs(qrise_end_time - qrise_start_time))
#            print 'Rise time = {:1.4f} seconds'.format(abs(qrise_end_time - qrise_start_time))
            #save rise time in data vectors
            if(qrise_start_time==qrise_end_time):       
                qrise_end_time += timestep
            self.riseTimeArrays[entry][channel-1].append([qrise_start_time, qrise_end_time]) #if no previous rise time has been calculated for that vector, the array is appended to the existing empty array. By adding measstart as offset, the actual time at which the marker should be present in the final plot is stored in the metadata
            return abs(qrise_end_time - qrise_start_time)
        else:
            return 0

        
#    def prise(self, entry, pchan, pmax=100, vchan=-1, vbase = 1.0, distStartTime=0, distEndTime=-1): 
#        """
#        #added optional parameters vchan, vbase, distStartTime, distEndTime. 
#            if distEndTime i provided, distStartTime does not need to be provided. 
#            if distEndTime is not provided, vchan should be provided, so that the function can detect the end of the disturbance itself
#        vchan can be used to have the function determine at what point in time the recovery period starts (voltage re-enters 0.9 to 1.1 p.u. range)
#        in case the voltage is not provide in p.u., a base needs to be provided, as it is compared against the normal operating range of 0.9 to 1.1
#        """
#        if type(pchan) is str:
#            chars = len(pchan)
#            for idx, name in enumerate(self.channel_names[entry]):
#                if name == pchan:
#                    pchan = idx + 1  # Channel numbers start at 1
#                    break
#        if type(vchan) is str:
#            chars = len(vchan)
#            for idx, name in enumerate(self.channel_names[entry]):
#                if name == vchan:
#                    vchan = idx + 1  # Channel numbers start at 1
#                    break
#        
#
#        time = self.timearrays[entry] + self.timeoffset[entry]
#        power = self.dataarrays[entry][pchan-1]
#        #time = self.timearrays[0]#!!!! WHY IS THE INITIAL TIME VECTOR BEING OVERWRITTEN HERE?? --> BE CAREFUL and DOUBLE-CHECK
#        
#        distStartIndex = np.argmin(abs(distStartTime - time))
#        
#        p_recovery = 0.0
#        #Pinit = power[0]# this only works for PSS/E cases where the system is perfectly initialised already
#        Pinit = power[max(distStartIndex-10, 0)]
#        print Pinit
#        for n in range(len(power) - 1):
##            if (power[n] < Pinit*0.95) and (power[n+1] > Pinit*0.95) and (time[n+1] < 2.0):
#            if (power[n] < Pinit*0.95) and (power[n+1] > Pinit*0.95):
#                p_recovery = time[n+1]
#                
#        if( (vchan!=-1) and (distEndTime==-1)):  
#            if(distStartIndex>10):                 
#                voltage=self.dataarrays[entry][vchan-1]
#                Vinit=voltage[max(distStartIndex-10, 0)] #reac voltage value ten samples before the supposed start time of the disturbance (or instant t=0, if distStartTime is not specified)
#                distEndTime=0.0
#                for n in range(len(voltage) -1):
#                    if( ((voltage[n]<0.9) and (voltage[n+1]>=0.9)) or ( (voltage[n]>1.1) and (voltage[n+1]<=1.1) ) ):
#                        distEndTime = time[n+1]
#        
#
#       
#        if( (Pinit>0.05*pmax) and (min(power[distStartIndex:-1])<0.95*Pinit) and (distEndTime < p_recovery) and (distEndTime!=-1) ): #only add recovery time to results if there is actually a drop of the power to below 0.95 p.u. AND distEndTime is not -1
#            self.recTimeArrays[entry][pchan-1].append([distEndTime, p_recovery]) #--> in plot routine, in case distEndTime==-1--> ignore because it likely means it was not specified
#        
#        return round(p_recovery, 3) #this only returns time at which power recovers, from start of dataset +offset --> if offset is t=1s and fault applied at t=2s for 120 ms and recovery time is 50 ms, then Prise will be 2s+0.12s+0.05s-1s - 1.17s


    def prise(self, entry, pchan, pmax=100, vchan=-1, vbase = 1.0, distStartTime=0, distEndTime=-1): 
        """
        #added optional parameters vchan, vbase, distStartTime, distEndTime. 
            if distEndTime i provided, distStartTime does not need to be provided. 
            if distEndTime is not provided, vchan should be provided, so that the function can detect the end of the disturbance itself
        vchan can be used to have the function determine at what point in time the recovery period starts (voltage re-enters 0.9 to 1.1 p.u. range)
        in case the voltage is not provide in p.u., a base needs to be provided, as it is compared against the normal operating range of 0.9 to 1.1
        """
        if type(pchan) is str:
            chars = len(pchan)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == pchan:
                    pchan = idx + 1  # Channel numbers start at 1
                    break
        if type(vchan) is str:
            chars = len(vchan)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == vchan:
                    vchan = idx + 1  # Channel numbers start at 1
                    break
        time = self.timearrays[entry] + self.timeoffset[entry]
        power = self.dataarrays[entry][pchan-1]
        #time = self.timearrays[0]#!!!! WHY IS THE INITIAL TIME VECTOR BEING OVERWRITTEN HERE?? --> BE CAREFUL and DOUBLE-CHECK
        measstart = np.argmin(abs(distEndTime - time))
#        measend = np.argmax(abs(distEndTime - time))
        power_ext = power[measstart:]
        time_ext = time[measstart:]
        distStartIndex = np.argmin(abs(distStartTime - time))
        p_recovery = 0.0
        #Pinit = power[0]# this only works for PSS/E cases where the system is perfectly initialised already
        Pinit = power[max(distStartIndex-10, 0)]
#        print Pinit
        if Pinit >=0: # find the recovery time if initialised active power being positive
            for n in range(len(power_ext) - 1):
    #            if (power[n] < Pinit*0.95) and (power[n+1] > Pinit*0.95) and (time[n+1] < 2.0):
                if (power_ext[n] < Pinit*0.95) and (power_ext[n+1] > Pinit*0.95):
    #                p_recovery = time[n+1]
                    p_recovery = time_ext[n+1] 
                    break
        else: # cover for the BESS case when power is negative - or the power is measured in opposite direction.
            for n in range(len(power_ext) - 1):
    #            if (power[n] < Pinit*0.95) and (power[n+1] > Pinit*0.95) and (time[n+1] < 2.0):
                if (power_ext[n] > Pinit*0.95) and (power_ext[n+1] < Pinit*0.95):
    #                p_recovery = time[n+1]
                    p_recovery = time_ext[n+1] 
                    break
        if( (vchan!=-1) and (distEndTime==-1)):  
            if(distStartIndex>10):                 
                voltage=self.dataarrays[entry][vchan-1]
                Vinit=voltage[max(distStartIndex-10, 0)] #reac voltage value ten samples before the supposed start time of the disturbance (or instant t=0, if distStartTime is not specified)
                distEndTime=0.0
                for n in range(len(voltage) -1):
                    if( ((voltage[n]<0.9) and (voltage[n+1]>=0.9)) or ( (voltage[n]>1.1) and (voltage[n+1]<=1.1) ) ):
#                        distEndTime = time[n+1]
                        distEndTime = time[n] # Considered to be recovered when moving into the recovery zone
        if Pinit >=0:
            if( (min(power[distStartIndex:-1])<0.95*Pinit) and (distEndTime < p_recovery) and (distEndTime!=-1) ): #only add recovery time to results if there is actually a drop of the power to below 0.95 p.u. AND distEndTime is not -1
                self.recTimeArrays[entry][pchan-1].append([distEndTime, p_recovery]) #--> in plot routine, in case distEndTime==-1--> ignore because it likely means it was not specified
        else:
            if( (max(power[distStartIndex:-1])>0.95*Pinit) and (distEndTime < p_recovery) and (distEndTime!=-1) ): #only add recovery time to results if there is actually a drop of the power to below 0.95 p.u. AND distEndTime is not -1
                self.recTimeArrays[entry][pchan-1].append([distEndTime, p_recovery]) #--> in plot routine, in case distEndTime==-1--> ignore because it likely means it was not specified

        return round(p_recovery, 3) #this only returns time at which power recovers, from start of dataset +offset --> if offset is t=1s and fault applied at t=2s for 120 ms and recovery time is 50 ms, then Prise will be 2s+0.12s+0.05s-1s - 1.17s



    
    def deltaIq(self, entry, Iqchan, Vchan=-1, Iqbase=1,Vbase=1.0, distStartTime=-1, distEndTime=-1, endoffset=50): #think about smart means of detecting end of disturbance automatically
        """
        calculate reactive current delta between pre-fault and during fault (sample taken at the end of the fault, when current has settled)
        if the voltage channel is provided, the voltage profile is used to determine when the disturbance is applied and when it ends. 
        if the distStarttime and distEndTime are explicitly provided, those values are used instead
        
        the starttime is interpreted from the time the "offset" starts. If theoffset for the dataset is set to 4.0s and a disturbance is applied at t=5s, the distStartTime would need to be set to 1.0s
        """
        if type(Iqchan) is str:
            chars = len(Iqchan)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == Iqchan:
                    Iqchan = idx + 1  # Channel numbers start at 1
                    break
                
        if type(Vchan) is str:
            chars = len(Vchan)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == Vchan:
                    Vchan = idx + 1  # Channel numbers start at 1
                    break   
                
        time=self.timearrays[entry] + self.timeoffset[entry] #will offset time by -4

        #If distStartTime not defined, but voltage vector provided, start at 0.0 and try to detect first step
        if(distStartTime==-1):
            distStartTime=0.0
            distStartIndex = np.argmin(abs(distStartTime - time))
            if(Vchan!=-1):
                voltage=self.dataarrays[entry][Vchan-1] 
                Vinit=voltage[max(distStartIndex-10, 0)] 
                for n in range(len(voltage)-1):
                    if( ( (voltage[n]>0.9) and (voltage[n+1]<=0.9) ) or ( (voltage[n]<1.1)and(voltage[n+1]>=1.1) )  ): #stepping into disturbance
                        distStartTime=time[n+1]
                          
        #If distStartTime 
        if(distEndTime==-1):
            distEndTime=distStartTime #distEndTime is at least equal to distStartTime, as the length of the disturbance cannot be negative
            if(Vchan!=-1):
                voltage=self.dataarrays[entry][Vchan-1]
                for n in range(len(voltage)-1):
                    if( ( (voltage[n]<0.9) and (voltage[n+1]>=0.9) ) or ( (voltage[n]>1.1)and(voltage[n+1]<=1.1) )  ): #stepping into disturbance
                        distEndTime=time[n+1]
                
        timestep = time[1] - time[0]        
        if timestep == 0: timestep = time[2] - time[1]
                        
        time=self.timearrays[entry] + self.timeoffset[entry] #will offset time by -4
        Iq= self.dataarrays[entry][Iqchan-1]
        
        startindex = np.argmin(abs(distStartTime - time)) - int(0.001/timestep*endoffset) #read current value 20 ms samples before the fault is applied. 
        endindex= np.argmin(abs(distEndTime - time))-int(0.001/timestep*endoffset) #determine index 20 ms samples before the end of the fault period
        
        Iq_init=Iq[startindex]
        Iq_fault=sum( Iq[endindex-3 : endindex] ) / len(Iq[endindex-3 : endindex])#average over the last few samples of the time period just before fault clearance
        
        delta_Iq=(Iq_fault-Iq_init)/Iqbase
        
        if(distEndTime-distStartTime>0.02): #fault needs to be longer than 20 ms, otherwise the values don't make sense and the metadata entry should not be created.
            self.dVdIq[entry][Iqchan-1]['dIq']=[Iq_init, Iq_fault, distStartTime, distEndTime] #save dIq entry in dataset
        
        return delta_Iq
        
        
    def deltaV(self, entry, channel, distStartTime=-1, distEndTime=-1, HV_calc_threshold=1.2, LV_calc_threshold=0.8, endoffset=50):
        """
        This function determines the voltage drop below the LV threshold or above the HV-threshold in p.u. during a voltage disturbance. 
        This value can be used in conjunction with dIq to verify compliance with current injection requirements per clause S5.2.5.5
        If the voltage is between the HV and LV threshold during the fault, status=0 and dV=0 will be returned.
        If the voltage is outside the limits dV is calculated and status=1 or status =2 is returned
        """
        if type(channel) is str:
            chars = len(channel)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == channel:
                    channel = idx + 1  # Channel numbers start at 1
                    break
        time=self.timearrays[entry] + self.timeoffset[entry] #will offset time by -4
        V= self.dataarrays[entry][channel-1]
        
        if(distStartTime==-1):
            distStartTime=0.0
            distStartIndex = np.argmin(abs(distStartTime - time))
            voltage=self.dataarrays[entry][channel-1]  
            for n in range(distStartIndex, len(voltage)-1):
                if( ( (voltage[n]>0.9) and (voltage[n+1]<=0.9) ) or ( (voltage[n]<1.1)and(voltage[n+1]>=1.1) )  ): #will return the lasst time the voltage stepped outside the normal operating band
                    distStartTime=time[n+1]
                          
        #If distStartTime 
        if(distEndTime==-1):
            distEndTime=distStartTime #distEndTime is at least equal to distStartTime, as the length of the disturbance cannot be negative  
            distEndIndex = np.argmin(abs(distEndTime - time))
            voltage=self.dataarrays[entry][channel-1]
            for n in range(distEndIndex, len(voltage)-1):
                if( ( (voltage[n]<0.9) and (voltage[n+1]>=0.9) ) or ( (voltage[n]>1.1)and(voltage[n+1]<=1.1) )  ): #will return the last time the plant stepped back into normal oeprating band
                    distEndTime=time[n+1]
        
        timestep = time[1] - time[0]
        if timestep == 0: timestep = time[2] - time[1]
        #startindex = np.argmin(abs(starttime - time)) - 10 #read current value 10 samples before the fault is applied. 
        endindex= np.argmin(abs(distEndTime - time))-int(0.001/timestep*endoffset) #determine index 10 samples before the end of the fault period
        
        V_fault=sum( V[endindex-3 : endindex] ) / len(V[endindex-3 : endindex])#average over the last few samples of the time period just before fault clearance
        
        dV=0
        status=0
        if(V_fault>=HV_calc_threshold): #overvoltage case
            dV=V_fault-HV_calc_threshold
            status=1
        
        if(V_fault<=LV_calc_threshold): #undervoltage case
            dV=LV_calc_threshold-V_fault
            status=2
        
        #add test to only create entry if distStartTime and distEndTime are not equal
        if(distEndTime-distStartTime>0.02): #fault neds to be longer than 20 ms, otherwise the values don't make sense and the metadata entry should not be created.
            self.dVdIq[entry][channel-1]['dV']=[HV_calc_threshold, LV_calc_threshold, V_fault, distStartTime, distEndTime]
        return status, dV, V_fault
        

    
#    def settleTime(self, entry, channel, starttime = 0.0, endtime = 10.0, startWindow = 1.0, endWindow = 5.0):
#        '''
#        Calculate settling time as defined in the NER
#        
#        locateStep() should be run for this entry before using this method.
#        
#        The difficulty in performing this task this can sometimes lie in the noise which
#        is encountered when recording responses. This method calculates the starting and
#        final values by taking an average value at the start and end of the period of 
#        interest. By default, the starting value will be the average value for the 
#        second before the step is applied and the final value is the average value 
#        for the five to ten second period after the step is applied. 
#        
#        These values can be overridden using the optional key words as required.
#        
#        Args:
#        -----
#            entry (int): Value of entry. Can be identified using the listEntries 
#                method. Designates the ID of dataset in memory
#            channel (int): Channel number. Can be identified using the listChannels
#                method
#            starttime (float) (optional): Time when the starting value window begins
#            
#            endtime (float) (optional): Time when the final value window ends
#            
#            startWindow (float) (optional): Size of the starting value average window
#                in seconds
#            
#            endWindow (float) (optional): Size of the final value average window
#                in seconds
#        
#        Returns:
#        --------
#            tset (float): Settling time in seconds from t = 0. Caution should be used
#                to ensure the settling time is not taken from the time of the step.
#        
#        Raises:
#        -------
#            None
#        
#        Examples:
#        ---------
#            >>> # Use locateStep to bring the step to t = 0.0 seconds
#            >>> [startV, startP, startQ] = plot1.locateStep(entry = 0, EfdChan = 6, event_number = 0)
#            >>> plot1.settleTime(entry = 0, channel = 2)
#        '''
#        if type(channel) is str:
#            chars = len(channel)
#            for idx, name in enumerate(self.channel_names[entry]):
#                if name == channel:
#                    channel = idx + 1  # Channel numbers start at 1
#                    break
#        
#        time = self.timearrays[entry] + self.timeoffset[entry]
#        data = self.dataarrays[entry][channel-1]
#        
#        timestep = time[1] - time[0]
#        timefreq = int(1. / timestep)
#        #
#        # Determine location in array of starttime, as it's easier to use array location than actual time in seconds.
#        measstart = np.argmin(abs(starttime - time))
#        measstartandwindow = np.argmin(abs(starttime - startWindow - time)) #window should be backward from the measstart (counting point)
#        measend = np.argmin(abs(endtime - time))
#        measendandwindow = np.argmin(abs(endtime + endWindow - time))
#        #
#        # print measend, startWindow, timefreq
#        if(measstart!=measstartandwindow):
#            startVal = sum( data[measstartandwindow : measstart] ) / len(data[measstartandwindow : measstart])
#            endVal = sum( data[measend : measendandwindow] ) / len(data[measend : measendandwindow])
#            maxdev = max( ( max(data[measstart:measend]) - data[measstart] ), ( data[measstart] - min(data[measstart:measend]) ) )
#            stepsize = round( (startVal - endVal) * 100 / 0.25, 0) * 0.25
#            # print measend, measstartandwindow
#            # print measstart, measend, time[measstart], time[measend], data[measstart], data[measend]
#            # print startVal, endVal, maxdev    
#            #
#            tempsettime = 0.00
#            # "voltage" settle time
#            if ( maxdev * 0.5 <= abs(endVal - startVal) ):
#                #
#                # Positive step 
#                if startVal < endVal:
#                    setBand1 = ( endVal + 0.1 * (endVal - startVal) )
#                    setBand2 = ( endVal - 0.1 * (endVal - startVal) )
#                    for j in range(measstart, measend):
#                        if (data[j] > setBand1) or (data[j] < setBand2):
#                            tempsettime = time[j]
#                    if tempsettime > 0.0:
#                        print 'Entry {:1d} channel {:1d} settling time = {:2.4f} sec, bands = {:2.4f}, {:2.4f}'.format(entry, channel, tempsettime, setBand1, setBand2)
#                    else:
#                        print 'Settling time could not be determined for entry {:1d} channel {:1d}'.format(entry, channel)
#                #
#                # Negative step
#                if startVal > endVal:
#                    setBand1 = ( endVal + 0.1 * ( startVal - endVal ) )
#                    setBand2 = ( endVal - 0.1 * ( startVal - endVal ) )
#                    for j in range(measstart, measend):
#                        if (data[j] > setBand1) or (data[j] < setBand2):
#                            tempsettime = time[j]
#                    if tempsettime > 0.0:
#                        print 'Entry {:1d} channel {:1d} settling time = {:2.4f} sec, bands = {:2.4f}, {:2.4f}'.format(entry, channel, tempsettime, setBand1, setBand2)
#                    else:
#                        print 'Settling time could not be determined for entry {:1d} channel {:1d}'.format(entry, channel)
#            #
#            # power settle time
#            if ( maxdev * 0.5 > abs(endVal - startVal) ):
#                setBand1 = ( endVal + (0.1 * maxdev) )
#                setBand2 = ( endVal - (0.1 * maxdev) )
#                for j in range(measstart, measend):
#                    if (data[j] > setBand1) or (data[j] < setBand2):
#                        tempsettime = time[j]
#                if tempsettime > 0.0:
#                    print 'Entry {:1d} channel {:1d} settling time = {:2.4f} sec, bands = {:2.4f}, {:2.4f}'.format(entry, channel, tempsettime, setBand1, setBand2)
#                else:
#                    print 'Settling time could not be determined for entry {:1d} channel {:1d}'.format(entry, channel)
#            
#            self.settle_arrays[entry][channel] = [setBand1, setBand2]
#            self.settleTimeArrays[entry][channel-1].append([time[measstartandwindow], tempsettime]) #time(given the first reference value is retrieved from the section before the step, but tempsettime is actually clacualted from the start of the window, the real settling time is de difference between the two values.
#        #If dataset corrupt
#        else:
#            tempsettime=0
#            
##        return round(tempsettime, 2)
#        return round(tempsettime, 3) # take into account the ms for Iq plot
    
    def settleTime(self, entry, channel, starttime = 0.0, endtime = 10.0, startWindow = 1.0, endWindow = 5.0):
        '''
        Calculate settling time as defined in the NER
        
        locateStep() should be run for this entry before using this method.
        
        The difficulty in performing this task this can sometimes lie in the noise which
        is encountered when recording responses. This method calculates the starting and
        final values by taking an average value at the start and end of the period of 
        interest. By default, the starting value will be the average value for the 
        second before the step is applied and the final value is the average value 
        for the five to ten second period after the step is applied. 
        
        These values can be overridden using the optional key words as required.
        
        Args:
        -----
            entry (int): Value of entry. Can be identified using the listEntries 
                method.
            channel (int): Channel number. Can be identified using the listChannels
                method
            starttime (float) (optional): Time when the starting value window begins
            
            endtime (float) (optional): Time when the final value window ends
            
            startWindow (float) (optional): Size of the starting value average window
                in seconds
            
            endWindow (float) (optional): Size of the final value average window
                in seconds
        
        Returns:
        --------
            tset (float): Settling time in seconds from t = 0. Caution should be used
                to ensure the settling time is not taken from the time of the step.
        
        Raises:
        -------
            None
        
        Examples:
        ---------
            >>> # Use locateStep to bring the step to t = 0.0 seconds
            >>> [startV, startP, startQ] = plot1.locateStep(entry = 0, EfdChan = 6, event_number = 0)
            >>> plot1.settleTime(entry = 0, channel = 2)
        '''
        if type(channel) is str:
            chars = len(channel)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == channel:
                    channel = idx + 1  # Channel numbers start at 1
                    break
        time = self.timearrays[entry] + self.timeoffset[entry]
        data = self.dataarrays[entry][channel-1]
        timestep = time[1] - time[0]
#        timefreq = int(1. / timestep)
        #
        # Determine location in array of starttime, as it's easier to use array location than actual time in seconds.
        measstart = np.argmin(abs(starttime - time))
        measstartandwindow = np.argmin(abs(starttime - startWindow - time)) #window should be backward from the measstart (counting point)
        measend = np.argmin(abs(endtime - time))
        measendandwindow = np.argmin(abs(endtime + endWindow - time))
        data_ext = data[measstart : measend]
        time_ext = time[measstart : measend]
        #
        # print measend, startWindow, timefreq
        if(measstart!=measstartandwindow):
            startVal = sum( data[measstartandwindow : measstart] ) / len(data[measstartandwindow : measstart])
            try: endVal = sum( data[measend : measendandwindow] ) / len(data[measend : measendandwindow])
            except: endVal = data[measend]
            maxdev = max( ( max(data[measstart:measend]) - data[measstart] ), ( data[measstart] - min(data[measstart:measend]) ) )
            stepsize = round( (startVal - endVal) * 100 / 0.25, 0) * 0.25
#            rangeVal =  abs(endVal - startVal) * 0.1
            # print measend, measstartandwindow
            # print measstart, measend, time[measstart], time[measend], data[measstart], data[measend]
            # print startVal, endVal, maxdev    
            tempsettime = 0.00
            # "voltage" settle time
            if ( maxdev * 0.5 <= abs(endVal - startVal) ):
                rangeVal =  abs(endVal - startVal) * 0.1
            else: 
                rangeVal =  maxdev * 0.1
                # Positive and negative step 
            setBand1 = ( endVal + 0.1 * (endVal - startVal) )
            setBand2 = ( endVal - 0.1 * (endVal - startVal) )
            for n in range(len(data_ext) - 1):
                if (abs(data_ext[n] - endVal) > rangeVal) and (abs(data_ext[n+1] - endVal) < rangeVal):
                    tempsettime = time_ext[n]                    
#            if tempsettime > 0.0:
#                print 'Entry {:1d} channel {:1d} settling time = {:2.4f} sec, bands = {:2.4f}, {:2.4f}'.format(entry, channel, tempsettime, setBand1, setBand2)
#            else:
#                print 'Settling time could not be determined for entry {:1d} channel {:1d}'.format(entry, channel)
            self.settle_arrays[entry][channel] = [setBand1, setBand2]
            self.settleTimeArrays[entry][channel-1].append([time_ext[0], tempsettime]) #time(given the first reference value is retrieved from the section before the step, but tempsettime is actually clacualted from the start of the window, the real settling time is de difference between the two values.
#            self.settleTimeArrays[entry][channel-1].append([time[measstartandwindow], tempsettime]) #time(given the first reference value is retrieved from the section before the step, but tempsettime is actually clacualted from the start of the window, the real settling time is de difference between the two values.

        #If dataset corrupt
        else:
            tempsettime=0
#        return round(tempsettime, 2)
        return round(tempsettime, 3) # take into account the ms for Iq plot

    def calloutd(self, entry, channel, callout_times = [0.0]):
        if type(channel) is str:
            chars = len(channel)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == channel:
                    channel = idx + 1  # Channel numbers start at 1
                    break
        time = self.timearrays[entry] + self.timeoffset[entry]
        data = self.dataarrays[entry][channel-1]
#        y_interp = np.interp(callout_time,time,data)
        dstamp = []
        for tstamp in callout_times:
            y_interp = np.interp(tstamp,time,data)
            dstamp.append(y_interp)
            
        return dstamp

    
    def comparison_bands(self, entry, channel, band):
        self.GSMG_arrays[entry][channel] = band

    def GSMG_bands(self, entry, channel, starttime = 0.0, endtime = 10.0, startWindow = 1.0, endWindow = 5.0):
        '''
        Determine GSMG_bands as defined in the PSMG
        
        This method is used to determine the error band following the PSMG. 
        The band is 10% of the induced change, or controlled change depending 
        on the test type. The idea is to first determine the transient window, 
        then within that transient window, look at the absolute deviation of 
        signal to the point prior to event occuring. This would be reasonablly 
        considered as the absolute change (induced or controlled) of the signal. 
        This value is then multiplied by 10% to achieve the error band.
        
        Note this is with assumption that the controlled change does not experience 
        significant overshoot. 
        
        These values can be overridden using the optional key words as required.
        
        Args:
        -----
            entry (int): Value of entry. Can be identified using the listEntries 
                method.
            channel (int): Channel number. Can be identified using the listChannels
                method
            starttime (float) (optional): Time when the starting value window begins
            endtime (float) (optional): Time when the final value window ends
            startWindow (float) (optional): Size of the starting value average window
                in seconds
            endWindow (float) (optional): Size of the final value average window
                in seconds
        
        Returns:
        --------
            error (float): GSMG_error_bands which is calculated from the maximum 
            deviation (induced change or controlled change) of data in transient window.
        
        Raises:
        -------
            None
        
        Examples:
        ---------

        '''
        #convert channel name to column number in case the it is porvided as a string
        if type(channel) is str:
            chars = len(channel)
            for idx, name in enumerate(self.channel_names[entry]):
                if name == channel:
                    channel = idx + 1  # Channel numbers start at 1
                    break
                
        time = self.timearrays[entry] + self.timeoffset[entry]
        data = self.dataarrays[entry][channel-1]
        
        timestep = time[1] - time[0]
        if timestep == 0: timestep = time[2] - time[1]
        timefreq = int(1. / timestep)
        #
        # Determine location in array of starttime, as it's easier to use array location than actual time in seconds.
        measstart = np.argmin(abs(starttime - time))
        measstartandwindow = np.argmin(abs(starttime - startWindow - time)) #window should be backward from the measstart (counting point)
        if endtime == -1: # if endtime = -1, take whole simulation period in consideration
#            measend = np.argmax(abs(endtime - time - endWindow))
            measendandwindow = np.argmax(abs(endtime - time))
            measend = measendandwindow - int(endWindow/timestep)
        else:            
            measend = np.argmin(abs(endtime - time))
            measendandwindow = np.argmin(abs(endtime + endWindow - time))
#        if measend == measendandwindow:
#            measend = measendandwindow - int(endWindow/timestep)
#        data_ext = data[measstart : measend]
#        time_ext = time[measstart : measend]
        #
        # print measend, startWindow, timefreq
        if(measstart!=measend):
            startVal = sum( data[measstartandwindow : measstart] ) / len(data[measstartandwindow : measstart])
            endVal = sum( data[measend : measendandwindow] ) / len(data[measend : measendandwindow])
            maxdev = max( ( max(data[measstart:measend]) - data[measstart] ), ( data[measstart] - min(data[measstart:measend]) ) )
#            stepsize = round( (startVal - endVal) * 100 / 0.25, 0) * 0.25
#            rangeVal =  abs(endVal - startVal) * 0.1
            # print measend, measstartandwindow
            # print measstart, measend, time[measstart], time[measend], data[measstart], data[measend]
            # print startVal, endVal, maxdev    
            #
            GSMG_error_bands = maxdev * 0.1 # 10% of the maximum deviation to be considered as error band
#            print GSMG_error_bands
#            print 'GSMG_error_bands = {:2.4f}, {:2.4f}'.format(GSMG_error_bands, -GSMG_error_bands)

            self.GSMG_arrays[entry][channel] = GSMG_error_bands
#            self.GSMG_arrays[entry][channel-1].append(GSMG_error_bands)
            
        #If dataset corrupt
        else:
            GSMG_error_bands=0
            
        return round(GSMG_error_bands, 3) # take into account the ms for Iq plot

    def export_csv(self, entry, filename):
        f = open(filename, 'w')
        
        line = ''
        for name in self.channel_names[entry]:
            line += '{},'.format(name)
        line += '\n'
        f.write(line)

        for idx, timestep in enumerate(self.timearrays[entry]):
            line += ''
            line += '{},'.format(self.timearrays[entry][idx])
            for n, value in enumerate(self.dataarrays[entry][idx]):
                line += '{},'.format(self.dataarrays[entry][idx][n])
            line += '\n'
            f.write(line)
        
        f.close()
    
    def check_min_max_time(self):
        """returns the highest time time values for each file that has been read by the plot script. 
        This can be used to determine a common time frame for which datas from multiple datasets is available, useful for overlays
        """
        max_times=[]
        min_times=[]
        for m in range(0,len(self.timearrays)):
            timearray =self.timearrays[m]
            max_times.append(max(timearray)+self.timeoffset[m])
            min_times.append(min(timearray))
        return [min_times, max_times]
            

    def subplot_spec(self, subplot, plot_arrays, title = '', ylabel = '', y2label='', scale = 1.0, offset = 0.0, twinX = False, markers=[], colour='', linewidth=2.5, linestyle='-', tolerance_band_offset=0.05, tolerance_band_base=-1):
        """
        Specify which channels are to be plotted

        Args:
            subplot: Number of subplot for which we are specifying files 
            and channels
            
            plot_arrays: The plot_arrays input should be specified as a two element tuple, 
            with the first element being the input file and the second being the channel 
            of that file to be plotted. The second element can be specified either 
            as the channel number (as an int), or the name as shown in 
            self.channel_names (as a str). If a string is provided, only the first
            few letters need to be specified sufficient to be unique from the other
            channel_names. If the string is not unique then the first match will be 
            plotted.

            title: Subplot title (optional)
            
            ylabel: Y axis title (optional)
            
            scale: Constant multiplier for this trace (optional)
            
            offset: Y axis offset for this trace (optional)

        Returns:
            Nothing

        Raises:
            Nothing

        """
        
        if plot_arrays[0] < len(self.dataarrays):
            if title != '':
                self.titles[subplot] = title
            if ylabel != '':
                self.ylabels[subplot] = ylabel
            if y2label != '':
                self.y2labels[subplot] =y2label
            #plot_arrays
            if type(plot_arrays[1]) is str:
                chars = len(plot_arrays[1])
                for idx, name in enumerate(self.channel_names[plot_arrays[0]]):
                    if name == plot_arrays[1]:
                        plot_channel = idx + 1  # Channel numbers start at 1
                        break
            else:
                plot_channel = plot_arrays[1]
            #
            self.plotspec[subplot].append((plot_arrays[0], plot_channel, twinX, markers, colour, linewidth, linestyle)) #specify which markers should be included for that specific data curve in the specific subplot. This can be 
            self.scales[plot_arrays[0]][plot_channel] = scale
            self.offsets[plot_arrays[0]][plot_channel] = offset
            self.tolerance_band_offset[plot_arrays[0]][plot_channel]=tolerance_band_offset
            self.tolerance_band_base[plot_arrays[0]][plot_channel]=tolerance_band_base
        else:
            print ('Specified file number {:d} is not in memory'.format(plot_arrays[0]))
            
        dataset_length=self.timearrays[plot_arrays[0]][-1] + self.timeoffset[plot_arrays[0]]
        return dataset_length 
            
    def plot(self, figname = '', show = 0, legloc = 'best'):#, verticalMarkers='None'): #
        '''
        Create plot of channels and files as specified in self.plotspec.\n
        Manually specify the following plot variables as required:\n
            self.legends[subplot]\n
            self.titles[subplot]\n
            self.ylabels[subplot]\n
            self.ylimit[subplot]\n
            self.xlimit
            #self.verticalMarkers --> the vertical markers allow to visualise settling time/rise time/recovery time in the plot and adds markers and labels showing the time. This is only included for the signals where the corresponding entry exists.
        '''
        import matplotlib.pyplot as plt
        from matplotlib.ticker import ScalarFormatter, FormatStrFormatter

        for idx, i in enumerate(self.plotspec):
            if i != []:
                subplots = idx + 1
        
        if subplots == 1:
            subplot_index = [111]
            plt.figure(figsize = (15 / 2.54, 15 / 2.54))
            plt.subplots_adjust(bottom = 0.10, top = 0.94, right = 0.97, left = 0.12, hspace = 0.25)
        if subplots == 2:
            subplot_index = [211, 212]
            plt.figure(figsize = (15 / 2.54, 15 / 2.54))
            plt.subplots_adjust(bottom = 0.08, top = 0.95, right = 0.97, left = 0.12, hspace = 0.28)
        if subplots == 3:
            subplot_index = [311, 312, 313]
#            plt.figure(figsize = (15 / 2.54, 20.6 / 2.54))
            plt.figure(figsize = (26.0/2.54, 29.7/ 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.97, left = 0.10, hspace = 0.32)
        if subplots == 4:
            subplot_index = [221, 222, 223, 224]
            plt.figure(figsize = (15 / 2.54, 15 / 2.54))
            plt.subplots_adjust(bottom = 0.08, top = 0.95, right = 0.97, left = 0.12, hspace = 0.32, wspace = 0.35)
        if subplots == 5:
            subplot_index = [311, 323, 324, 325, 326]
            plt.figure(figsize = (15 / 2.54, 20.6 / 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.97, left = 0.12, hspace = 0.32, wspace = 0.35)
        if subplots == 6:
            subplot_index = [321, 322, 323, 324, 325, 326]
            plt.figure(figsize = (36.0 / 2.54, 22.5/ 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
        if subplots == 7:
            subplot_index = [421, 423, 424, 425, 426, 427, 428]
            plt.figure(figsize = (36.0 / 2.54, 22.5/ 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
        if subplots == 8:
            subplot_index = [421, 422, 423, 424, 425, 426, 427, 428]
            plt.figure(figsize = (36.0/2.54, 22.5/ 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
#        if subplots in [9,10]:
#            subplot_index = [521, 522, 523, 524, 525, 526, 527, 528, 529]
#            plt.figure(figsize = (22.5/ 2.54, 36.0/2.54))
#            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
#        if subplots == 9:
#            subplot_index = [[(5,2),(0,0)],[(5,2),(1,0)],[(5,2),(2,0)],[(5,2),(3,0)],[(5,2),(4,0)],[(5,2),(0,1)],[(5,2),(1,1)],[(5,2),(2,1)],[(5,2),(3,1)],[(5,2),(4,1)]]
#            plt.figure(figsize = (36.0/2.54, 22.5/ 2.54))
#            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
#        if subplots == 10:
#            subplot_index = [[(5,2),(0,0)],[(5,2),(1,0)],[(5,2),(2,0)],[(5,2),(3,0)],[(5,2),(4,0)],[(5,2),(0,1)],[(5,2),(1,1)],[(5,2),(2,1)],[(5,2),(3,1)],[(5,2),(4,1)]]
#            plt.figure(figsize = (36.0/2.54, 22.5/ 2.54))
#            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
        if subplots == 9:
            subplot_index = [[(5,2),(0,0)],[(5,2),(0,1)],[(5,2),(1,0)],[(5,2),(1,1)],[(5,2),(2,0)],[(5,2),(2,1)],[(5,2),(3,0)],[(5,2),(3,1)],[(5,2),(4,0)],[(5,2),(4,1)]]
            plt.figure(figsize = (36.0/2.54, 22.5/ 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
        if subplots == 10:
            subplot_index = [[(5,2),(0,0)],[(5,2),(0,1)],[(5,2),(1,0)],[(5,2),(1,1)],[(5,2),(2,0)],[(5,2),(2,1)],[(5,2),(3,0)],[(5,2),(3,1)],[(5,2),(4,0)],[(5,2),(4,1)]]
#            plt.figure(figsize = (36.0/2.54, 22.5/ 2.54))
#            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
            plt.figure(figsize = (26.0/2.54, 29.7/ 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.55, right = 0.55, left = 0.05, hspace = 0.20, wspace = 0.08)
        if subplots == 12: #Plot PQV in 3 columns network case
            subplot_index = [[(4,3),(0,0)],[(4,3),(0,1)],[(4,3),(0,2)],[(4,3),(1,0)],[(4,3),(1,1)],[(4,3),(1,2)],[(4,3),(2,0)],[(4,3),(2,1)],[(4,3),(2,2)],[(4,3),(3,0)],[(4,3),(3,1)],[(4,3),(3,2)]]
#            plt.figure(figsize = (36.0/2.54, 22.5/ 2.54))
#            plt.subplots_adjust(bottom = 0.05, top = 0.97, right = 0.99, left = 0.05, hspace = 0.27, wspace = 0.1)
            plt.figure(figsize = (26.0/2.54, 29.7/ 2.54))
            plt.subplots_adjust(bottom = 0.05, top = 0.55, right = 0.55, left = 0.05, hspace = 0.20, wspace = 0.08)
        #
        plt.rc('xtick', labelsize = 10)
        plt.rc('ytick', labelsize = 10)
        #
        if self.xlimit != []:
            self.intervals=[]
            #
            for m in range(len(self.dataarrays)):
                starttime = self.xlimit[0] - self.timeoffset[m]
                endtime = self.xlimit[1] - self.timeoffset[m]
                time = self.timearrays[m]
                measstart = np.argmin(abs(starttime - time)) #identify sample closest to starttime
                measend = np.argmin(abs(endtime - time)) #identify sample closest to endtime
                self.intervals.append([measstart, measend])
#                self.timearrays[m] = self.timearrays[m][measstart:measend] #self.timearrays[m]
#                temp = np.zeros((len(self.dataarrays[m]), measend-measstart))
#                for n in range(len(self.dataarrays[m])):
#                    temp[n] = self.dataarrays[m][n][measstart:measend]            
#                self.plotdataarrays.append(temp)
        else:
            self.plotdataarrays = self.dataarrays  # plotdataarrays seems to be the same as dataarrays, it is just time-limited if x-axis range is specified --> What is the benefit of copying everything?
            for m in range(len(self.dataarrays)):
                measstart=0
                measend = len(self.timearrays[m]-1)
                self.intervals.append([measstart, measend])
        #
        # print subplots
        for idx in range(subplots):
            if len(range(subplots))<=8:
                plt.subplot(subplot_index[idx])
            else:
                plt.subplot2grid(subplot_index[idx][0],subplot_index[idx][1])
#            else:
#                plt.subplot2grid(subplot_index[idx][0],subplot_index[idx][1],subplot_index[idx][2])
            ax1 = plt.gca()
            # ax1.yaxis.set_major_formatter(FormatStrFormatter('%1.{:d}f'.format(2)))            
            ax1.ticklabel_format(useOffset = False)
            #
            line_handles=[]
            for trace in self.plotspec[idx]: #plots data traces for subplot
                datafile = trace[0]
                channel = trace[1]
                twinX = trace[2]
                markers = trace[3]
                colour=trace[4]
                linewidth=trace[5]
                linestyle=trace[6]
                if not twinX:
                    if(colour==''):
                        line_handle=plt.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                                            (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]
                                                            * self.scales[datafile][channel]) 
                                                            + self.offsets[datafile][channel], lw = linewidth, linestyle=linestyle)
                    else:
                        line_handle=plt.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                                            (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]
                                                            * self.scales[datafile][channel]) 
                                                            + self.offsets[datafile][channel], lw = linewidth, color=colour, linestyle=linestyle)
                    line_handles.append(line_handle)
                if twinX:
                    color = 'cyan'
                    ax2 = ax1.twinx()  # instantiate a second axes that shares the same x-axis
                    #ax2.set_ylabel('sin', color=color)  # we already handled the x-label with ax1
                    line_handle=ax2.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                                        (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]] 
                                                        * self.scales[datafile][channel]) 
                                                        + self.offsets[datafile][channel], lw = 2.5, color = color)
                    line_handles.append(line_handle)
                    plt.subplots_adjust(right = 0.92)
                    
                
                        
            #
            for trace in self.plotspec[idx]:#add markers to subplot
                datafile = trace[0]
                channel = trace[1]
                markers = trace[3]
                #
                if('GSMG' in markers):
                    pass
#                    # Plot upper GSMG boundary
#                    if self.GSMG_arrays[datafile][channel] != 0:
#                        ax1.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
#                                ((self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]  + self.GSMG_arrays[datafile][channel]) * self.scales[datafile][channel]) + 
#                                self.offsets[datafile][channel-1], 'm--', lw = 1.5, c='grey')
#                    #
#                    # Plot lower GSMG boundary
#                    if self.GSMG_arrays[datafile][channel] != 0:
#                        ax1.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
#                                ((self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]  - self.GSMG_arrays[datafile][channel]) * self.scales[datafile][channel]) + 
#                                self.offsets[datafile][channel-1], 'm--', lw = 1.5, c='grey')


                if('set_band' in markers):
                    #
                    # Plot upper settling band
                    if self.settle_arrays[datafile][channel] != []:
                        upper_band = (self.settle_arrays[datafile][channel][0] * self.scales[datafile][channel]) + self.offsets[datafile][channel]
                        plt.plot([min(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile]), max(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile])], [upper_band, upper_band], 'k--', lw = 1.5, c='grey')
    
                    #
                    # Plot lower settling band
                    if self.settle_arrays[datafile][channel] != []:
                        lower_band = (self.settle_arrays[datafile][channel][1] * self.scales[datafile][channel]) - self.offsets[datafile][channel]
                        plt.plot([min(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile]), max(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile])], [lower_band, lower_band], 'k--', lw = 1.5, c='grey')
                if('tolerance_bands' in markers):
                    if((self.tolerance_band_offset[datafile][channel])==-1):
                        plt.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                                                (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]
                                                                * self.scales[datafile][channel]*(1+self.tolerance_band_offset[datafile][channel])) 
                                                                + self.offsets[datafile][channel], 'm--', lw = 1.5, c='grey') #upper band
                        plt.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                                                (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]
                                                                * self.scales[datafile][channel]*(1-self.tolerance_band_offset[datafile][channel])) 
                                                                + self.offsets[datafile][channel], 'm--', lw = 1.5, c='grey') #upper band
                    else:
                        plt.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                                                (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]
                                                                * self.scales[datafile][channel])
                                                                + self.offsets[datafile][channel]+(self.tolerance_band_offset[datafile][channel]*self.tolerance_band_base[datafile][channel]), 'm--', lw = 1.5, c='grey') #upper band
                        plt.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                                                (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]
                                                                * self.scales[datafile][channel]) 
                                                                + self.offsets[datafile][channel]-(self.tolerance_band_offset[datafile][channel]*self.tolerance_band_base[datafile][channel]), 'm--', lw = 1.5, c='grey') #upper band
            #
            if self.legends[idx] != []:
                if not twinX:
                    plt.legend(self.legends[idx], prop = {'size': 9}, loc = legloc)
                else:
                    lines=line_handles[0]
                    for line_id in range(1, len(line_handles)):
                        lines+=line_handles[line_id]
                    plt.legend(lines, self.legends[idx], prop = {'size': 9}, loc = legloc)                    
            #
            if self.titles[idx] != []:
                plt.title(self.titles[idx], fontsize = 12)
            #
            if (self.ylabels[idx] != []): #and (twinX == False):
                ax1.set_ylabel(self.ylabels[idx], fontsize = 10)
            if (self.y2labels[idx] != []):#  and (twinX == True):#Add second axis label not as soon as a label is defined. make sure to only define it if a secondary axis ic actually used. 
                ax2.set_ylabel(self.y2labels[idx], fontsize = 10)
            #
            if self.ylimit[idx] != []:
                plt.ylim(self.ylimit[idx])
            elif self.yspan[idx] != 0:
                limits=ax1.get_ylim()
                if(limits[1]-limits[0]<self.yspan[idx]):
                    offset=(self.yspan[idx] - (limits[1]-limits[0]))/2
                    plt.ylim(limits[0]-offset, limits[1]+offset) 
            limits=ax1.get_ylim()
            if(limits[1]>self.ymaxlim[idx]):
                plt.ylim(limits[0], self.ymaxlim[idx])
                limits=ax1.get_ylim()
            if(limits[0]<self.yminlim[idx]):
                plt.ylim(self.yminlim[idx], limits[1])
                limits=ax1.get_ylim()
            #
            if self.y2limit[idx] != []:
                ax2.set_ylim(self.y2limit[idx]) #this is assuming that y2lim is only defined if there is a trace plotted on the secondary axis. 
            elif self.y2span[idx] !=0:
                limits=ax2.get_ylim()
                if(limits[1]-limits[0]<self.y2span[idx]):
                    offset=(self.y2span[idx] - (limits[1]-limits[0]))/2
                    plt.ylim(limits[0]-offset, limits[1]+offset)  
            if self.xlimit != []:
                plt.xlim(self.xlimit)
                #plt.xlim([0,6])
            else:
                xmin, xmax = ax1.get_xlim()
                ticks = ax1.get_xticks()
                plt.xlim([0, ticks[-2]])
            plt.xlabel('Time (sec)', fontsize = 10)
            plt.grid(1)
#            fig.tight_layout()    
            
            #Add markers and time labels per specification in plotspec entry. 
            #This has to be done after the axis ranges are adjusted, because the inclusion arrows depends on that
            for trace in self.plotspec[idx]:
                ax1 = plt.gca()
                datafile = trace[0]
                channel = trace[1]
                twinX = trace[2]
                markers = trace[3]
                label_y_offset=0 #increase after every label added, to avoid overlapping of labels
                label_x_offset=0 #used to determine length of horizontal markers and position of vertical arrows 
                for marker in markers:

                    if('GSMG' in markers):
                        # Plot upper GSMG boundary
                        if self.GSMG_arrays[datafile][channel] != 0:
                            ax1.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                    ((self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]  + self.GSMG_arrays[datafile][channel]) * self.scales[datafile][channel]) + 
                                    self.offsets[datafile][channel-1], 'm--', lw = 1.5, c='grey')
                        #
                        # Plot lower GSMG boundary
                        if self.GSMG_arrays[datafile][channel] != 0:
                            ax1.plot(self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile], 
                                    ((self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]  - self.GSMG_arrays[datafile][channel]) * self.scales[datafile][channel]) + 
                                    self.offsets[datafile][channel-1], 'm--', lw = 1.5, c='grey')

                    if(marker=='set_t'):
                        #check if entry in data structure exists
                        settle_entry=self.settleTimeArrays[datafile][channel-1]
                        if(settle_entry!=[]): #my metainformation is stored starting at position 0
                            for step in settle_entry:
                                starttime=step[0]
                                endtime=step[1]
                                plt.axvline(x=starttime, ymin=0, ymax=1, color='grey',ls='--', lw = 1.5)
                                plt.axvline(x=endtime, ymin=0, ymax=1, color='grey',ls='--', lw = 1.5)
                                x_range=ax1.get_xlim()
                                x_range=x_range[1]-x_range[0] #retrieve x_range of plot
                                y_range=ax1.get_ylim()
                                yloc=y_range[0]+(label_y_offset+1)*(y_range[1]-y_range[0])/10 #allow for 8 labels to be stacked vertically
                                if(endtime-starttime)/x_range>0.01: 
                                    ax1.annotate('', xy=(starttime, yloc), xycoords='data',
                                                 xytext=(endtime, yloc), textcoords='data',
                                                 arrowprops={'arrowstyle': '<->'})
                                    
                                ax1.annotate('set=' +str(round(endtime-starttime, 3))+' s', 
                                             xy=(endtime, yloc), xycoords='data',
                                             xytext=(7, -2), textcoords='offset points', fontsize=7)
                                pass
                                    #add arrow between the two markers, if distance in plot large enough
                                #add label right of the arrow (or above the arrow) stating settling time in s
                                label_y_offset+=1                       
                    if(marker=='rise_t'):
                        #check if entry in data structure exists
                        rise_entry=self.riseTimeArrays[datafile][channel-1]
                        if(rise_entry!=[]): #my metainformation is stored starting at position 0
                            for step in rise_entry:
                                starttime=step[0]
                                endtime=step[1]
                                plt.axvline(x=starttime, ymin=0, ymax=1, color='grey',ls='--', lw = 1.5)
                                plt.axvline(x=endtime, ymin=0, ymax=1, color='grey',ls='--', lw = 1.5)
                                x_range=ax1.get_xlim()
                                x_range=x_range[1]-x_range[0] #retrieve x_range of plot
                                y_range=ax1.get_ylim()
                                yloc=y_range[0]+(label_y_offset+1)*(y_range[1]-y_range[0])/10 #allow for 8 labels to be stacked vertically
                                if(endtime-starttime)/x_range>0.01: 
                                    ax1.annotate('', xy=(starttime, yloc), xycoords='data',
                                                 xytext=(endtime, yloc), textcoords='data',
                                                 arrowprops={'arrowstyle': '<->'})
                                    
                                ax1.annotate('rise=' +str(round(endtime-starttime, 3))+' s', 
                                             xy=(endtime, yloc), xycoords='data',
                                             xytext=(7, -2), textcoords='offset points', fontsize=7)
                                pass
                                    #add arrow between the two markers, if distance in plot large enough
                                #add label right of the arrow (or above the arrow) stating settling time in s
                                label_y_offset+=1

                    if(marker=='rec_t'):    
                        #check if entry in data structure exists
                        rec_entry=self.recTimeArrays[datafile][channel-1]
                        if(rec_entry!=[]): #my metainformation is stored starting at position 0
                            for step in rec_entry:
                                starttime=step[0]
                                endtime=step[1]
                                plt.axvline(x=starttime, ymin=0, ymax=1, color='grey',ls='--', lw = 1.5)
                                plt.axvline(x=endtime, ymin=0, ymax=1, color='grey',ls='--', lw = 1.5)
                                x_range=ax1.get_xlim()
                                x_range=x_range[1]-x_range[0] #retrieve x_range of plot
                                y_range=ax1.get_ylim()
                                yloc=y_range[0]+(label_y_offset+1)*(y_range[1]-y_range[0])/10 #allow for 8 labels to be stacked vertically
                                if(endtime-starttime)/x_range>0.01: 
                                    ax1.annotate('', xy=(starttime, yloc), xycoords='data',
                                                 xytext=(endtime, yloc), textcoords='data',
                                                 arrowprops={'arrowstyle': '<->'})
                                    
                                ax1.annotate('rec=' +str(round(endtime-starttime, 3))+' s', 
                                             xy=(endtime, yloc), xycoords='data',
                                             xytext=(7, -2), textcoords='offset points', fontsize=7)
                                pass
                                    #add arrow between the two markers, if distance in plot large enough
                                #add label right of the arrow (or above the arrow) stating settling time in s
                                label_y_offset+=1

                    if(marker=='dV'):
                        pass
                        #check if entry in data structure exists
                        dVdIq_entry=self.dVdIq[datafile][channel-1]
                        if('dV'in dVdIq_entry.keys()):
                            dV_entry=dVdIq_entry['dV']
                            startTime=dV_entry[3]#-self.timeoffset[datafile]
                            endTime=dV_entry[4]#-self.timeoffset[datafile]
                            if not twinX: #for now, markers can only be placed for trace that is displayed on primary axis
                                #add HV threshold horizontal marker
                                #determine X limits for the line as unite between 0 and 1. 
                                x_range=ax1.get_xlim()
                                x0=x_range[0]
                                x_range=x_range[1]-x_range[0]
                                rel_start=(startTime-x0)/x_range #determine start of disturbance in plot relative coordinates where 0 corresponds to left end of the plot and 1 corresponds to right end of the plot. this is used for positionin the labels and determining the length of the horizontal markers.
                                rel_end=(endTime-x0)/x_range
                                
                                plt.axhline(y=dV_entry[0], xmin=max(0, rel_start-1/10.0), xmax=min(0.9, rel_end+(1+label_x_offset)/10.0), color='grey',ls='--', lw = 1.5) #extend marker 10% of the total width of the plot beyond the fault start and fault end 
                                #add LV threshold horizontal marker
                                plt.axhline(y=dV_entry[1], xmin=max(0, rel_start-1/10.0), xmax= min(0.9, rel_end+(1+label_x_offset)/10.0), color='grey',ls='--', lw = 1.5) #extend marker 10% of the total width of the plot beyond the fault start and fault end 
                                #add line showing voltage during the fault
                                plt.axhline(y=dV_entry[2], xmin=max(0, rel_start-1/10.0), xmax=min(0.9, rel_end+(1+label_x_offset)/10.0), color='grey',ls='--', lw = 1.5) #extend marker 10% of the total width of the plot beyond the fault start and fault end 

                                if(dV_entry[2]>dV_entry[0]): #over voltage case
                                    #add arrow from HV threshold to V_fault
                                    ax1.annotate('', xy=(endTime+(x_range/10)*(1+label_x_offset), dV_entry[0]), xycoords='data',
                                                 xytext=(endTime+(x_range/10)*(1+label_x_offset), dV_entry[2]), textcoords='data',
                                                 arrowprops={'arrowstyle': '<->'})
                                    ax1.annotate('dV=' +str(round(dV_entry[2]-dV_entry[0], 3))+' p.u.', 
                                             xy=(endTime+(x_range/10)*(1+label_x_offset), (dV_entry[0]+dV_entry[2])/2), xycoords='data',
                                             xytext=(7, -2), textcoords='offset points', fontsize=7)
                                elif(dV_entry[2]<dV_entry[1]):
                                    #add arrow from LV threshold to V_fault
                                    ax1.annotate('', xy=(endTime+(x_range/10)*(1+label_x_offset), dV_entry[1]), xycoords='data',
                                                 xytext=(endTime+(x_range/10)*(1+label_x_offset), dV_entry[2]), textcoords='data',
                                                 arrowprops={'arrowstyle': '<->'})
                                    ax1.annotate('dV=' +str(round(dV_entry[2]-dV_entry[1], 3))+' p.u.', 
                                             xy=(endTime+(x_range/10)*(1+label_x_offset), (dV_entry[1]+dV_entry[2])/2), xycoords='data',
                                             xytext=(7, -2), textcoords='offset points', fontsize=7)
                                #label the two lines indicating the calculation threshold.
                                ax1.annotate('LV-lim', 
                                             xy=(startTime-(x_range/10)*(1+label_x_offset), (dV_entry[1])), xycoords='data',
                                             xytext=(-30, -2), textcoords='offset points', fontsize=7)
                                ax1.annotate('HV_lim', 
                                             xy=(startTime-(x_range/10)*(1+label_x_offset), (dV_entry[0]) ), xycoords='data',
                                             xytext=(-30, -2), textcoords='offset points', fontsize=7)
                                label_x_offset+=1

                            #determine which y-axis to use
                            #addhorizontal marker on threshold
                            #addhorizontal marker on value during fault
                            #if distance sufficient: add vertical arrow between two markers
                            #add label right of the marker
                    if (marker=='dIq'):
                        pass
                        dVdIq_entry=self.dVdIq[datafile][channel-1]
                        if('dIq'in dVdIq_entry.keys()):
                            dIq_entry=dVdIq_entry['dIq']
                            startTime=dIq_entry[2]#-self.timeoffset[datafile]
                            endTime=dIq_entry[3]#-self.timeoffset[datafile]
                            if not twinX: #for now, markers can only be placed for trace that is displayed on primary axis
                                x_range=ax1.get_xlim()
                                x0=x_range[0]
                                x_range=x_range[1]-x_range[0]
                                rel_start=(startTime-x0)/x_range #determine start of disturbance in plot relative coordinates where 0 corresponds to left end of the plot and 1 corresponds to right end of the plot. this is used for positionin the labels and determining the length of the horizontal markers.
                                rel_end=(endTime-x0)/x_range
                                
                                plt.axhline(y=dIq_entry[0], xmin=max(0, rel_start-1/10.0), xmax=min(0.9, rel_end+(1+label_x_offset)/10.0), color='grey',ls='--', lw = 1.5) #extend marker 10% of the total width of the plot beyond the fault start and fault end 
                                #add LV threshold horizontal marker
                                plt.axhline(y=dIq_entry[1], xmin=max(0, rel_start-1/10.0), xmax= min(0.9, rel_end+(1+label_x_offset)/10.0), color='grey',ls='--', lw = 1.5) #extend marker 10% of the total width of the plot beyond the fault start and fault end 
                                
                                #add arrow from HV threshold to V_fault
                                ax1.annotate('', xy=(endTime+(x_range/10)*(1+label_x_offset), dIq_entry[0]), xycoords='data',
                                             xytext=(endTime+(x_range/10)*(1+label_x_offset), dIq_entry[1]), textcoords='data',
                                             arrowprops={'arrowstyle': '<->'})
                                ax1.annotate('dIq=' +str(round(dIq_entry[1]-dIq_entry[0], 3))+' p.u.', 
                                         xy=(endTime+(x_range/10)*(1+label_x_offset), (dIq_entry[0]+dIq_entry[1])/2), xycoords='data',
                                         xytext=(7, -2), textcoords='offset points', fontsize=7)

                        #check if entry in data structure exists
                        #if(dV!=>0):
                            #determine which y-axis to use
                            #addhorizontal marker on threshold
                            #addhorizontal marker on value during fault
                            #if distance sufficient: add vertical arrow between two markers
                            #add label right of the marker
                    
                    if (marker=='callout'):
                        #check if entry in data structure exists
                        callout_entry=self.callout[datafile][channel-1]
                        if(callout_entry!=[]): #my metainformation is stored starting at position 0
                            callout_times = [1.5, 5.5] #time stamp on the graph would like to call the data out (by default at 1.5 and 5.5secs)
                            if type(markers[-1]) != str: # if the callout_times input is provided at the end of markers list (type should be list not a string), take it
                                callout_times = markers[-1]
                            for tstamp in callout_times:
                                try:
                                    y_interp = np.interp(tstamp,self.timearrays[datafile][self.intervals[datafile][0]:self.intervals[datafile][1]] + self.timeoffset[datafile],
                                                     (self.dataarrays[datafile][channel-1][self.intervals[datafile][0]:self.intervals[datafile][1]]
                                                                    * self.scales[datafile][channel]) 
                                                                    + self.offsets[datafile][channel])
                                    plt.plot(tstamp,y_interp,'o',color='k')
#                                    ax1.annotate(y_interp,(tstamp,y_interp))
                                    ax1.annotate(str(round(tstamp, 3))+", "+ str(round(y_interp, 3)), 
                                                 xy=(tstamp, y_interp), xycoords='data',
                                                 xytext=(0, 7), textcoords='offset points', fontsize=7)
                                except: 
                                    pass

        if figname != '':
            plt.tight_layout()
            plt.savefig(figname + '.png', format = 'png', dpi=300)
            #plt.savefig(figname + '.emf', format = 'emf')
            plt.savefig(figname + '.svg', format = 'svg')
            # imgdata=StringIO()
            # plt.savefig(imgdata, dpi =200)
            # plt.savefig(imgdata)
            imgdata=io.BytesIO()
            plt.savefig(imgdata)

        if show: 
            plt.show()
            
        plt.clf()
        plt.close()
        
        return imgdata