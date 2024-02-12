# -*- coding: utf-8 -*-
"""
Created on Tue Dec 19 13:51:36 2023

@author: wateasle

Don't have files open when running merge program, because ghost files will be included.
That might cause data duplication.

"""

import os
import sys
from datetime import date
import numpy as np
import pandas as pd
# Module tkinter for file selection function.
from tkinter import Tk, Label, ttk, filedialog


def user_file_choice(wdir):
    # See tkinter functions here: https://docs.python.org/3/library/dialog.html
    
    # Create an instance of tkinter frame or window
    win = Tk()
    # This might be causing the program to crash.
    #win.withdraw()
    
    # Set the geometry of tkinter frame
    win.geometry("700x350")
    
    def open_file():
        global file_list
        file_paths = filedialog.askopenfilenames(parent=win,
                                                title='Select Files.', 
                                                initialdir=wdir)
        file_list = [os.path.basename(path) for path in file_paths]
        win.destroy()
    
    # Add a Label widget
    label = Label(win, text="Open the file directory and select desired files:", font=('Aerial 11'))
    label.pack(pady=30)
    
    # Add a Button Widget
    ttk.Button(win, text="Select files", command=open_file).pack()
    
    win.mainloop()
    
    return file_list

def main():    
    
    # Get the directory of where the current script is.
    # Note that this won't work unless executing script via Anaconda prompt.
    wdir = os.path.dirname(os.path.abspath(__file__))
    # Directory to use for testing.
    #wdir = r"H:\My Drive\2022 - RSH\RSH (Fred)\Data\Soil Data, Analyses\Total Carbon, Nitrogen"
    #wdir = os.getcwd()
    
    # Set working directory to the script directory
    os.chdir(wdir)
    
    # DEBUG
    # print("wdir: ", wdir)
    # print("cwd: ", os.getcwd())
    
    # Allow user to select desired files using tkinter GUI.
    file_list = user_file_choice(wdir)
    
    # DEBUG
    # print(file_list)
    
    # Date variable for generating files.
    d = date.today().strftime("%Y%m%d")
    
    ###
    
    # Code used prior to implementation of tkinter GUI for file selection.
    
    # # List of files and directories within working directory.
    # names = os.listdir(wdir)
    
    # # Select only excel files (.xlsx and .xlsm) to merge.    
    # file_list = []
    # for f in os.listdir(wdir):
    #     name, ext = os.path.splitext(f)
    #     if ext == ".xlsx" or ext == ".xlsm":
    #         file_name = name + ext
    #         # Note that the "+=" syntax does not work in place of function append here.
    #         file_list.append(file_name)

    ###
    
    # Read in separate excel files and create list of dataframes.
    # In excel, if two rows are part of column header, read first row as header, -
    # then skip row zero to avoid pulling in row of NA values.
    # To do this, chain ".drop(index=0)" after read_excel().
    frames = [pd.read_excel(workbook,
                            header=0,
                            sheet_name="data", 
                            na_values="-").drop(index=0) for workbook in file_list]
    
    df_concat = pd.concat(frames, ignore_index=True)
    # Verify no accidentally duplicated columns in concatenated data frame.
    # df_concat.columns
    
    # Fill in depth "0-15" whereever NAN.
    if 'Depth' not in df_concat.columns:
        df_concat['Depth'] = np.nan
    df_concat['Depth'] =  df_concat['Depth'].fillna(value='0-15')
    
    # Sort dataframe.
    df_sorted = df_concat.sort_values(by=['Year', 'Sample Time', 'Depth', 
                                          'Site', 'Plot'], axis=0)

    # Drop any empty rows and columns.
    df_sorted = df_sorted.dropna(axis=1, how="all").dropna(axis=0, how="all")
    
    # Drop rows that won't be needed in future analyses.
    df_sorted = df_sorted.drop(columns=['Site_Plot'])
    
    # Give column names that will be compatible with R, SAS.
    df_sorted = df_sorted.rename(columns={
                              'Sample ID': 'Sample_ID',
                              'Sample Time': 'Sample_Time',
                              'Carbon (%)': 'Soil_C_pct', 
                              'Nitrogen (%)': 'Soil_N_pct',
                              'Sample Mass (g)': 'Sample_Mass_g'})
    
    # Add columns with calculated data.
    df_sorted['CN Ratio'] = df_sorted['Soil_C_pct'] / df_sorted['Soil_N_pct']
    
    return df_sorted.to_csv(d + "_stats_soil_CN.csv", index=False)
    
if __name__ == '__main__':
    main()
