#Lib for all the Calculation stuff

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
    def Impedance_Calculation(Rounded, Debug):        # Calculating Impedance, Resistance and Blindwiderstand from Voltage, Current, Frequency and Phase Offset
        file_path = os.path.join(Data_Path, "Clean.xlsx")   # Cleaned data from MSO5000
        dfCal = pd.read_excel(file_path)                    # Cleaned data from MSO5000
        dfCalRounded = dfCal.copy() # Copy of dataframe for rounded values
        Xmax = dfCal.shape[1]       # Number of columns
        Ymax = dfCal.shape[0] - 1   # Number of rows
        X = 0
        Y = 0
        if (Debug == "yes"):    # Debug Messages
            print("Xmax =", Xmax)
            print("Ymax =", Ymax)

        dfCal.iloc[:, X+5] = dfCal.iloc[:, X+5].astype(complex, copy=False)
        dfCalRounded.iloc[:, X+5] = dfCalRounded.iloc[:, X+5].astype(complex, copy=False)

        while (Y <= Ymax):                      #Calculating Impedance, Resistance and Blind for each measurement point
            Voltage = dfCal.iloc[Y,X]           #reading Voltage
            Current = dfCal.iloc[Y,X+1]         #reading Current
            Frequenzy = dfCal.iloc[Y,X+2]       #reading Frequency
            PhaseOffset = dfCal.iloc[Y, X+3]    #reading Phase Offset

            Impedance_abs = Voltage / Current                                   #Calculating Impedance in Ohm
            Resistance  = math.cos(math.radians(PhaseOffset)) * Impedance_abs   #Calculating Resistance in Ohm
            Blind       = math.sin(math.radians(PhaseOffset)) * Impedance_abs   #Calculating Blindwiderstand in Ohm
            Impedance   = complex(Resistance, Blind)                            #Calculating Complex Impedance in Ohm

            dfCal.iloc[Y,X+4] = Impedance_abs   #Storing Impedance in Ohm
            dfCal.iloc[Y,X+5] = Impedance       #Complex Impedance in Ohm
            dfCal.iloc[Y,X+6] = Resistance      #Storing Resistance in Ohm
            dfCal.iloc[Y,X+7] = Blind           #Storing Blindwiderstand in Ohm

            # Rounding all values to the specified decimal places in settings
            Rounded_Voltage =       round(Voltage, Rounded)
            Rounded_Current =       round(Current, Rounded)
            Rounded_Frequeny =      round(Frequenzy, Rounded)
            Rounded_PhaseOffset =   round(PhaseOffset, Rounded)
            Rounded_Impedance_abs = round(Impedance_abs, Rounded)
            Rounded_Impedance =     round(Resistance, Rounded) + round(Blind, Rounded)*1j
            Rounded_Resistance =    round(Resistance, Rounded)
            Rounded_Blind =         round(Blind, Rounded)

            # Storing rounded values in new dataframe
            dfCalRounded.iloc[Y,X]   = Rounded_Voltage
            dfCalRounded.iloc[Y,X+1] = Rounded_Current
            dfCalRounded.iloc[Y,X+2] = Rounded_Frequeny
            dfCalRounded.iloc[Y,X+3] = Rounded_PhaseOffset
            dfCalRounded.iloc[Y,X+4] = Rounded_Impedance_abs
            dfCalRounded.iloc[Y,X+5] = Rounded_Impedance
            dfCalRounded.iloc[Y,X+6] = Rounded_Resistance
            dfCalRounded.iloc[Y,X+7] = Rounded_Blind

            Y += 1  #Next Row
        file_path = os.path.join(Data_Path, "Clean_Calc.xlsx")          # Exporting calculated data to Excel File
        dfCal.to_excel(file_path, index = False)
        file_path = os.path.join(Data_Path, "Clean_Calc_Rounded.xlsx")  # Exporting calculated data to Excel File
        dfCalRounded.to_excel(file_path, index = False)

if(False): #Functions Layer 2
    print(1)

if(False): #Functions Layer 3
    print(1)