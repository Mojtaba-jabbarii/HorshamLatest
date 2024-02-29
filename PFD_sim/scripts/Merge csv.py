# -*- coding: utf-8 -*-
"""
Created on Fri Nov 17 10:31:30 2023

@author: 341510mim3
"""
import pandas as pd
import os
# import csv
import glob    
results_path = r"C:\Users\341510mim3\OneDrive - OX2\Projects\Summerville\1. Main Test Environment\Results data_test for script"
all_files = glob.glob(os.path.join(results_path, "*.csv"))
writer = pd.ExcelWriter(results_path + "PQ Results.xlsx", engine='xlsxwriter')

for f in all_files:
    df = pd.read_csv(f)
    df.to_excel(writer, sheet_name=os.path.basename(f))

writer.save()

#Deleting original files
for f in all_files:
    path = os.path.join(results_path, f)
    os.remove(path)
    
    
    
