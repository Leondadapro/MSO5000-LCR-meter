#Lib for all the Settings stuff

from calendar import c
from re import DEBUG
import sys
from turtle import clear
import pandas as pd
import matplotlib.pyplot as plt
import os
import time
import subprocess
import numpy as np
import math
import pyvisa
import msvcrt
from enum import Enum

if (True):  #define Paths
    Lib_Dir = os.path.dirname(os.path.abspath(__file__))
    Base_Dir = os.path.dirname(Lib_Dir)
    Settings_Path = os.path.join(Base_Dir, "Settings")
    Data_Path = os.path.join(Base_Dir, "Data")

if(True): #Functions Layer 1
    def Settings(Type, What, dfData):   # Calling Settings from Excel File
        match Type:
            case "Default":

                file_path = os.path.join(Settings_Path, "Settings_Default.xlsx")
                dfSettings = pd.read_excel(file_path, header=None, index_col=None)
            
                if(What == "Load"): # Load Default Settings
                    Rounded     = int(dfSettings.iloc[0,1])
                    Time_Delay  = float(dfSettings.iloc[1,1])
                    Debug       = str(dfSettings.iloc[2,1])

                    file_path_current = os.path.join(Settings_Path, "Settings_Current.xlsx")
                    dfSettingsCurrent = dfSettings
                    dfSettingsCurrent.to_excel(file_path_current, index = False, header=None)

                    return Rounded, Time_Delay, Debug

            case "Custom":

                file_path = os.path.join(Settings_Path, "Settings_Custom.xlsx")
                dfSettings = pd.read_excel(file_path, header=None, index_col=None)
            
                if(What == "Load"): # Load Custom Settings
                    Rounded     = int(dfSettings.iloc[0,1])
                    Time_Delay  = float(dfSettings.iloc[1,1])
                    Debug       = str(dfSettings.iloc[2,1])

                    file_path_current = os.path.join(Settings_Path, "Settings_Current.xlsx")
                    dfSettingsCurrent = dfSettings
                    dfSettingsCurrent.to_excel(file_path_current, index = False, header=None)

                    return Rounded, Time_Delay, Debug

                if(What == "Save"): # Save Custom Settings
                    dfSettings = dfData
                    file_path_current = os.path.join(Settings_Path, "Settings_Current.xlsx")
                    dfSettings.to_excel(file_path, index = False, header=None)
                    dfSettings.to_excel(file_path_current, index = False, header=None)

            case "Current":

                file_path = os.path.join(Settings_Path, "Settings_Current.xlsx")
                dfSettings = pd.read_excel(file_path, header=None, index_col=None)

                if(What == "Show"): # Show Current Settings
                    print(dfSettings)

                if(What == "Save"): # Save Current Settings
                    dfData.to_excel(file_path, index = False, header=None)

                if(What == "Load"): # Load Current Settings
                    Rounded     = int(dfSettings.iloc[0,1])
                    Time_Delay  = float(dfSettings.iloc[1,1])
                    Debug       = str(dfSettings.iloc[2,1])

                    return Rounded, Time_Delay, Debug

if(False): #Functions Layer 2
    print(1)

if(False): #Functions Layer 3
     print(1)