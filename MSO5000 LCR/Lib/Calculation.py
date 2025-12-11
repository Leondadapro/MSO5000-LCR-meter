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
import Lib.Debug as d
from enum import Enum

if (True):  #define Paths
    Lib_Dir         = os.path.dirname(os.path.abspath(__file__))
    Base_Dir        = os.path.dirname(Lib_Dir)
    Settings_Path   = os.path.join(Base_Dir, "Settings")
    Data_Path       = os.path.join(Base_Dir, "Data")

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

        # Remove the old float64 column
        dfCal.drop(dfCal.columns[X+5], axis=1, inplace=True)
        dfCalRounded.drop(dfCalRounded.columns[X+5], axis=1, inplace=True)

        # Insert a new complex column at the same position
        dfCal.insert(X+5, f"Z_[Ω]", pd.Series(np.zeros(len(dfCal), dtype=np.complex128), index=dfCal.index))
        dfCalRounded.insert(X+5, f"Z_[Ω]", pd.Series(np.zeros(len(dfCalRounded), dtype=np.complex128), index=dfCalRounded.index))

        while (Y <= Ymax):                      
            ( 
            Voltage,        Rounded_Voltage,
            Current,        Rounded_Current,
            Frequenzy,      Rounded_Frequeny,
            PhaseOffset,    Rounded_PhaseOffset,
            Impedance_abs,  Rounded_Impedance_abs,
            Resistance,     Rounded_Resistance,
            Blind,          Rounded_Blind,
            Impedance,      Rounded_Impedance 
            ) = Calc_All(Rounded, dfCal, Y, X)  # Calculate everything

            Save_All(dfCal, dfCalRounded, X, Y,
                 Voltage, Rounded_Voltage,
                 Current, Rounded_Current,
                 Frequenzy, Rounded_Frequeny,
                 PhaseOffset, Rounded_PhaseOffset,
                 Impedance_abs, Rounded_Impedance_abs,
                 Resistance, Rounded_Resistance,
                 Blind, Rounded_Blind,
                 Impedance, Rounded_Impedance)  # Save everything

            Y += 1  # Next Row
        file_path = os.path.join(Data_Path, "Clean_Calc.xlsx")          # Exporting calculated data to Excel File
        dfCal.to_excel(file_path, index = False)
        file_path = os.path.join(Data_Path, "Clean_Calc_Rounded.xlsx")  # Exporting calculated data to Excel File
        dfCalRounded.to_excel(file_path, index = False)

if(True): # Functions Layer 2
    def Calc_All(Rounded, dfCal, Y, X):
        Voltage,        Rounded_Voltage         = Read_Voltage(Rounded, dfCal, Y, X)            # reading Voltage
        Current,        Rounded_Current         = Read_Current(Rounded, dfCal, Y, X)            # reading Current
        Frequenzy,      Rounded_Frequeny        = Read_Frequenzy(Rounded, dfCal, Y, X)          # reading Frequenzy
        PhaseOffset,    Rounded_PhaseOffset     = Read_PhaseOffset(Rounded, dfCal, Y, X)        # reading Phaseoffset
        Impedance_abs,  Rounded_Impedance_abs   = Calc_Impedance(Rounded, dfCal, Y, X)          # calculating impedanc
        Resistance,     Rounded_Resistance      = Calc_Resistance(Rounded, dfCal, Y, X)         # calculating resistance
        Blind,          Rounded_Blind           = Calc_Blind(Rounded, dfCal, Y, X)              # calculating blindwiderstand
        Impedance,      Rounded_Impedance       = Calc_Impedanz_Complex(Rounded, dfCal, Y, X)   # calculating complex impedance
        return  ( 
                Voltage,        Rounded_Voltage,
                Current,        Rounded_Current,
                Frequenzy,      Rounded_Frequeny,
                PhaseOffset,    Rounded_PhaseOffset,
                Impedance_abs,  Rounded_Impedance_abs,
                Resistance,     Rounded_Resistance,
                Blind,          Rounded_Blind,
                Impedance,      Rounded_Impedance 
                )

    def Save_All(dfCal, dfCalRounded, X, Y,
                 Voltage, Rounded_Voltage,
                 Current, Rounded_Current,
                 Frequenzy, Rounded_Frequeny,
                 PhaseOffset, Rounded_PhaseOffset,
                 Impedance_abs, Rounded_Impedance_abs,
                 Resistance, Rounded_Resistance,
                 Blind, Rounded_Blind,
                 Impedance, Rounded_Impedance):

        dfCal, dfCalRounded = Save_Voltage(dfCal, dfCalRounded, X, Y, Voltage, Rounded_Voltage)
        dfCal, dfCalRounded = Save_Current(dfCal, dfCalRounded, X, Y, Current, Rounded_Current)
        dfCal, dfCalRounded = Save_Frequenzy(dfCal, dfCalRounded, X, Y, Frequenzy, Rounded_Frequeny)
        dfCal, dfCalRounded = Save_PhaseOffset(dfCal, dfCalRounded, X, Y, PhaseOffset, Rounded_PhaseOffset)
        dfCal, dfCalRounded = Save_Impedance(dfCal, dfCalRounded, X, Y, Impedance_abs, Rounded_Impedance_abs)
        dfCal, dfCalRounded = Save_Resistance(dfCal, dfCalRounded, X, Y, Resistance, Rounded_Resistance)
        dfCal, dfCalRounded = Save_Blind(dfCal, dfCalRounded, X, Y, Blind, Rounded_Blind)
        dfCal, dfCalRounded = Save_Impedanz_Complex(dfCal, dfCalRounded, X, Y, Impedance, Rounded_Impedance)
        return dfCal, dfCalRounded

if(True): # Functions Layer 3
    # Functions for Reading Values from Dataframe and Calculating stuff
    def Read_Voltage(Rounded, dfCal, Y, X):
        Voltage = dfCal.iloc[Y,X]           #reading Voltage
        Rounded_Voltage = Round_Sig(Voltage, Rounded)
        return Voltage, Rounded_Voltage
    def Read_Current(Rounded, dfCal, Y, X):
        Current = dfCal.iloc[Y,X+1]         #reading Current
        Rounded_Current =       Round_Sig(Current, Rounded)
        return Current, Rounded_Current
    def Read_Frequenzy(Rounded, dfCal, Y, X):
        Frequenzy = dfCal.iloc[Y,X+2]       #reading Frequency
        Rounded_Frequeny =      Round_Sig(Frequenzy, Rounded)
        return Frequenzy, Rounded_Frequeny
    def Read_PhaseOffset(Rounded, dfCal, Y, X):
        PhaseOffset = dfCal.iloc[Y, X+3]    #reading Phase Offset
        Rounded_PhaseOffset =   Round_Sig(PhaseOffset, Rounded)
        return PhaseOffset, Rounded_PhaseOffset
    def Calc_Impedance(Rounded, dfCal, Y, X):
        Voltage, Rounded_Voltage = Read_Voltage(Rounded, dfCal, Y, X) #reading Voltage
        Current, Rounded_Current = Read_Current(Rounded, dfCal, Y, X) #reading Current

        Impedance_abs = Voltage / Current                                   #Calculating Impedance in Ohm
        Rounded_Impedance_abs = Round_Sig(Impedance_abs, Rounded)
        return Impedance_abs, Rounded_Impedance_abs
    def Calc_Resistance(Rounded, dfCal, Y, X):
        PhaseOffset, Rounded_PhaseOffset = Read_PhaseOffset(Rounded, dfCal, Y, X) #reading Phaseoffset
        Impedance_abs, Rounded_Impedance_abs = Calc_Impedance(Rounded, dfCal, Y, X) #calculating impedanc

        Resistance  = math.cos(math.radians(PhaseOffset)) * Impedance_abs   #Calculating Resistance in Ohm
        Rounded_Resistance =    Round_Sig(Resistance, Rounded)
        return Resistance, Rounded_Resistance
    def Calc_Blind(Rounded, dfCal, Y, X):
        PhaseOffset, Rounded_PhaseOffset = Read_PhaseOffset(Rounded, dfCal, Y, X) #reading Phaseoffset
        Impedance_abs, Rounded_Impedance_abs = Calc_Impedance(Rounded, dfCal, Y, X) #calculating impedanc

        Blind       = math.sin(math.radians(PhaseOffset)) * Impedance_abs   #Calculating Blindwiderstand in Ohm
        Rounded_Blind =         Round_Sig(Blind, Rounded)
        return Blind, Rounded_Blind
    def Calc_Impedanz_Complex(Rounded, dfCal, Y, X):
        Resistance, Rounded_Resistance = Calc_Resistance(Rounded, dfCal, Y, X)
        Blind, Rounded_Blind = Calc_Blind(Rounded, dfCal, Y, X)

        Impedance   = complex(Resistance, Blind)       
        Rounded_Impedance =     Rounded_Resistance + Rounded_Blind*1j
        return Impedance, Rounded_Impedance

    # Functions for Saving Values to Dataframe
    def Save_Voltage(dfCal, dfCalRounded, X, Y, Voltage, Rounded_Voltage):
        dfCal.iloc[Y,X]   = Voltage         #Storing Voltage in V
        dfCalRounded.iloc[Y,X]   = Rounded_Voltage
        return dfCal, dfCalRounded
    def Save_Current(dfCal, dfCalRounded, X, Y, Current, Rounded_Current):
        dfCal.iloc[Y,X+1] = Current         #Storing Current in A
        dfCalRounded.iloc[Y,X+1] = Rounded_Current
        return dfCal, dfCalRounded
    def Save_Frequenzy(dfCal, dfCalRounded, X, Y, Frequenzy, Rounded_Frequeny):
        dfCal.iloc[Y,X+2] = Frequenzy
        dfCalRounded.iloc[Y,X+2] = Rounded_Frequeny
        return dfCal, dfCalRounded
    def Save_PhaseOffset(dfCal, dfCalRounded, X, Y, PhaseOffset, Rounded_PhaseOffset):
        dfCal.iloc[Y, X+3] = PhaseOffset
        dfCalRounded.iloc[Y, X+3] = Rounded_PhaseOffset
        return dfCal, dfCalRounded
    def Save_Impedance(dfCal, dfCalRounded, X, Y, Impedance, Rounded_Impedance):
        dfCal.iloc[Y,X+4] = Impedance
        dfCalRounded.iloc[Y,X+4] = Rounded_Impedance
        return dfCal, dfCalRounded
    def Save_Resistance(dfCal, dfCalRounded, X, Y, Resistance, Rounded_Resistance):
        dfCal.iloc[Y,X+6] = Resistance
        dfCalRounded.iloc[Y,X+6] = Rounded_Resistance
        return dfCal, dfCalRounded
    def Save_Blind(dfCal, dfCalRounded, X, Y, Blind, Rounded_Blind):
        dfCal.iloc[Y,X+7] = Blind
        dfCalRounded.iloc[Y,X+7] = Rounded_Blind
        return dfCal, dfCalRounded
    def Save_Impedanz_Complex(dfCal, dfCalRounded, X, Y, Impedance, Rounded_Impedance):
        dfCal.iloc[Y,X+5] = Impedance
        dfCalRounded.iloc[Y,X+5] = Rounded_Impedance
        return dfCal, dfCalRounded

if(True): # Functions Layer 4
    def Round_Sig(x, Rounded):
        if x == 0:
            return 0
        return round(x, Rounded - 1 - int(math.floor(math.log10(abs(x)))))