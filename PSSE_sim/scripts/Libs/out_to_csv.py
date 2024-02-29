# -*- coding: utf-8 -*-
"""
Created on Tue Jun  2 15:42:00 2020

@author: Mervin Kall
"""

from __future__ import with_statement
from contextlib import contextmanager
import pandas as pd

import sys
import os
script_path= os.path.dirname(os.path.abspath(__file__))

sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
sys.path.append(sys_path_PSSE)
os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
os.environ['PATH'] += ';' + os_path_PSSE
os.environ['PATH'] += ';' + sys_path_PSSE

import glob, os, sys, math, csv, time, logging, traceback, exceptions, os.path
import psse34 
import psspy
import redirect
import shutil
import dyntools
from win32com import client
#import numpy as np
import math    
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()
redirect.psse2py()

def save_results(MYOUTFILE, csv_file):
    outfile=dyntools.CHNF(MYOUTFILE)
    short_title, chanName_dict, chandata_dict = outfile.get_data()
    
    dataDict={}
    
    df_time = pd.Series(chandata_dict['time'], name = chanName_dict['time'])
    df=df_time
    for i in sorted(chanName_dict.keys()):
        if i !="time":
            dataDict[chanName_dict[i]]=chandata_dict[i]
            temp = pd.Series(chandata_dict[i], name = chanName_dict[i])
            df=pd.concat([df,temp],axis=1)
            
    dirname, flnm=os.path.split(psspy.sfiles()[0].rstrip(".sav"))
    
    df.to_csv("{}.csv".format(csv_file), sep=',', index=False)
    
def main(outfile, csv_path):  
    save_results(outfile, csv_path)

if __name__ == '__main__':
    main(r'C:\Users\PSCAD\Desktop\mervins stuff\WoreeSynConTemp\20201221_SMIB_testing\PSSE_sim\model_copies\20201221_FaultAlignmentPSSE_6\large40\1. PSSE Model1B_BDAX-10-600_AUTO_large40.out', )