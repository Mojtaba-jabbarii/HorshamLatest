# -*- coding: utf-8 -*-
"""
Created on Tue Jul 11 16:18:25 2023

@author: 341510davu
"""
# import sys
import os, sys
import pandas as pd
import numpy as np
from datetime import datetime
import time
from contextlib import contextmanager
from win32com.client import Dispatch
timestr=time.strftime("%Y%m%d-%H%M%S")

try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO
from io import BytesIO

###############################################################################
#USER INPUTS
###############################################################################
TestDefinitionSheet=r'20230828_SUM_TESTINFO_V1.xlsx'
raw_SS_result_folder = '20240123-112638_S5251'
simulation_batches=['S5251_PQcurve']
simulation_batch_label = simulation_batches[0]

cases = ["35degC_BESS_PV_0.9", "35degC_BESS_PV_1.0"]

Vcp = ["0.9","1.0"]
temperature = ["35degC","50degC"]
    
summary_dfs={'35degC':{}, '50degC':{},} # use input model names

probs={'35degC':0, '50degC':0}

###############################################################################
# Supporting functions
###############################################################################

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

# Generate plots
def generate_plot(x_axis, y_axes, legends, label_x, label_y, title):
#    x_axis = df_vlt_lvl['Bus Name']
#    y_axes = [df_vlt_lvl['Voltage Level(pu) GenON'], df_vlt_lvl['Voltage Level(pu) GenOFF']]
#    legends = ['VL(pu) GenON', 'VL(pu) GenOFF']
#    label_x = 'Voltage(pu)'
#    label_y = 'Bus Names'
#    title = 'Voltage Levels'
    # size and positions
    fig = plt.figure(figsize=(7,5))
    fig.add_axes([0.1,0.1,0.8,0.8])
    colors = ['k', 'r', 'b', 'g']
    markers = ['-', '--', '^','*', '+']
    # axis names and ticks
    for i in range(len(y_axes)):
        ydata = y_axes[i]
        ldata = legends[i]
        if len(x_axis) > 1: xdata = x_axis[i]
        else: xdata = x_axis[0]
        plt.scatter(xdata, ydata, label = ldata, color = colors[i], marker = markers[i], s = 100)
        
#    plt.scatter(x_axis, y_axes[0], label = legends[0], color = 'k', marker = '*', s = 120)
#    plt.scatter(x_axis, y_axes[1], label = legends[1])
    plt.title(title, fontsize =12, color = 'black' )
    plt.ylabel(label_y, fontsize = 10)
    plt.xlabel(label_x, fontsize = 10)
    plt.legend()
#    plt.xticks(rotation = 45)
    plt.margins( tight = True)
    plt.grid()
    plt.minorticks_off()
    #plt.savefig(case +'voltage_levels'+ 'plot.png', bbox_inches = 'tight')
    #imgdata= StringIO.StringIO()
    #imgdata = StringIO()
    imgdata= BytesIO() # version issues
    plt.savefig(imgdata, bbox_inches = 'tight', dpi =200)    
    return imgdata


###############################################################################
# Define Project Paths
###############################################################################

#main_folder_path=os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
script_dir=os.getcwd()
script_dir_up=os.path.abspath(os.path.join(script_dir, os.pardir))
main_folder_path=os.path.abspath(os.path.join(script_dir_up, os.pardir))
sys.path.append(main_folder_path+"\\PSSE_sim\\scripts\\Libs")
# Create directory for storing the results
if "OneDrive - OX2" in main_folder_path: # if the current folder is online (under OneDrive - OX2), create a new directory to store the result
    user = os.path.expanduser('~')
    main_path_out = main_folder_path.replace(user + "\OneDrive - OX2","C:\work") # Change the path from Onedrive to Local in c drive
    main_folder_out = createPath(main_path_out)
else: # if the main folder is not in Onedrive, then store the results in the same location with the model
    main_folder_out = main_folder_path
    
dir_path =  main_folder_out +"\\Plots\\PQ_curve"
make_dir(dir_path)

# Create shortcut linking to result folder if it is not stored in the main folder path
if main_folder_out != main_folder_path:
    createShortcut(main_folder_out, main_folder_path + "\\Plots\\PQ_curve.lnk")
else: # if the output location is same as input location, then delete the links
    try:os.remove(main_folder_path + "\\Plots\\PQ_curve.lnk")
    except: pass

result_sheet_path = main_folder_out +"\\PSSE_sim\\result_data\\PQ_curve\\" + raw_SS_result_folder + "\\S5251_PQ curve results.xlsx"

###############################################################################
# Import additional functions
###############################################################################
import matplotlib.pyplot as plt
sys.path.append(r"C:\ProgramData\Anaconda2\Lib\site-packages")
sys.path.append(r"C:\Python27\Lib\site-packages")
import docx
import openpyxl as xl
import re
from docx import Document, shape
from docx.oxml import OxmlElement, parse_xml
from docx.shared import Inches, Pt
from docx.oxml.ns import qn, nsdecls
from docx.enum.text import WD_BREAK
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.dml import MSO_THEME_COLOR_INDEX
import readtestinfo

###############################################################################
# Main function
###############################################################################

return_dict =  readtestinfo.readTestdef(main_folder_path+"\\test_scenario_definitions\\"+TestDefinitionSheet, ['ProjectDetails','ModelDetailsPSSE'])
ProjectDetailsDict = return_dict['ProjectDetails']
PSSEmodelDict = return_dict['ModelDetailsPSSE']

def main():
    
    wb_in = xl.load_workbook(filename=result_sheet_path)# Read data input
    
    reportname_prefix= timestr+"-"+str(ProjectDetailsDict['NameShrt']+ str(simulation_batch_label))
    writer = pd.ExcelWriter(main_folder_out+"\\Plots\\PQ_curve\\"+reportname_prefix+"_PQcurve.xlsx",engine = 'xlsxwriter') # Preparing for exporting the result
    
#    df_plot={'35degC':{}, '50degC':{},} # use input model names
#    df_plot = {}
    for case in temperature:
        df_out = pd.DataFrame()
#        df_plot{case} = pd.DataFrame()
        for i in range(len(wb_in.worksheets)):
            if case in str(wb_in.worksheets[i]):
                ws_in1 = wb_in.worksheets[i]
                ws_in1.delete_cols(idx=8, amount =4) # current
                ws_in1.delete_cols(idx=1) # index
                data1 = ws_in1.values
                df_out1 = pd.DataFrame.from_dict(data = data1)
                df_out = pd.concat([df_out,df_out1],axis = 1)
                
#                cols = [6,7]
#                df_out2 = df_out1[df_out1.columns[cols]]
#                df_plot = pd.concat([df_plot,df_out2],axis = 1)
##                df_out1[df_out1.columns[7,8]]
                
        df_out.to_excel(writer, sheet_name = str(case))
        
#        x_axis = df_plot[df_plot.columns[[1]]]
#        y_axes = df_plot[df_plot.columns[[2]]]
#        legends = ["0.9"]
#        label_x = "Q MVAr"
#        label_y = "P MW"
#        title = "PQ curve"
#        generate_plot(x_axis, y_axes, legends, label_x, label_y, title)

    writer.close() 


if __name__ == '__main__':
    main()
    
    