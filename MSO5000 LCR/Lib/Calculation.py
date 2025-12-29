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
    def Impedance_Calculation(Rounded, Debug): # Main Function for Calculating Impedance and other stuff
        #Reading Data from Excel File and preparing variables
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
        dfCal.drop(dfCal.columns[X+6], axis=1, inplace=True)
        dfCalRounded.drop(dfCalRounded.columns[X+6], axis=1, inplace=True)

        # Insert a new complex column at the same position
        dfCal.insert(X+6, f"Z_[Ω]", pd.Series(np.zeros(len(dfCal), dtype=np.complex128), index=dfCal.index))
        dfCalRounded.insert(X+6, f"Z_[Ω]", pd.Series(np.zeros(len(dfCalRounded), dtype=np.complex128), index=dfCalRounded.index))

        while (Y <= Ymax):                      
            ( 
            Voltage_Ue,     Rounded_Voltage_Ue,
            Voltage_Ua,     Rounded_Voltage_Ua,
            Current,        Rounded_Current,
            Frequenzy,      Rounded_Frequeny,
            PhaseOffset,    Rounded_PhaseOffset,
            Impedance_abs,  Rounded_Impedance_abs,
            Resistance,     Rounded_Resistance,
            Blind,          Rounded_Blind,
            Impedance,      Rounded_Impedance,
            H,              Rounded_H,
            H_db,           Rounded_H_db
            ) = Calc_All(Rounded, dfCal, Y, X)  # Calculate everything

            Save_All    (
                        dfCal, dfCalRounded, X, Y,
                        Voltage_Ue, Rounded_Voltage_Ue,
                        Voltage_Ua, Rounded_Voltage_Ua,
                        Current, Rounded_Current,
                        Frequenzy, Rounded_Frequeny,
                        PhaseOffset, Rounded_PhaseOffset,
                        Impedance_abs, Rounded_Impedance_abs,
                        Resistance, Rounded_Resistance,
                        Blind, Rounded_Blind,
                        Impedance, Rounded_Impedance,
                        H, Rounded_H,
                        H_db, Rounded_H_db
                        )  # Save everything

            Y += 1  # Next Row

        #Creating Excel File with Calculated Data
        file_path = os.path.join(Data_Path, "Clean_Calc.xlsx")          # Exporting calculated data to Excel File
        dfCal.to_excel(file_path, index = False)
        file_path = os.path.join(Data_Path, "Clean_Calc_Rounded.xlsx")  # Exporting calculated data to Excel File
        dfCalRounded.to_excel(file_path, index = False)

if(True): # Functions Layer 2
    def Calc_All(Rounded, dfCal, Y, X):
        Voltage_Ue,     Rounded_Voltage_Ue      = Read_Voltage_Ue(Rounded, dfCal, Y, X)             # reading Voltage Ue
        Voltage_Ua,     Rounded_Voltage_Ua      = Read_Voltage_Ua(Rounded, dfCal, Y, X)             # reading Voltage Ua
        Current,        Rounded_Current         = Read_Current(Rounded, dfCal, Y, X)                # reading Current
        Frequenzy,      Rounded_Frequeny        = Read_Frequenzy(Rounded, dfCal, Y, X)              # reading Frequenzy
        PhaseOffset,    Rounded_PhaseOffset     = Read_PhaseOffset(Rounded, dfCal, Y, X)            # reading Phaseoffset
        Impedance_abs,  Rounded_Impedance_abs   = Calc_Impedance(Rounded, dfCal, Y, X)              # calculating impedanc
        Resistance,     Rounded_Resistance      = Calc_Resistance(Rounded, dfCal, Y, X)             # calculating resistance
        Blind,          Rounded_Blind           = Calc_Blind(Rounded, dfCal, Y, X)                  # calculating blindwiderstand
        Impedance,      Rounded_Impedance       = Calc_Impedanz_Complex(Rounded, dfCal, Y, X)       # calculating complex impedance
        H,              Rounded_H               = Calc_Transferfunction_1(Rounded, dfCal, Y, X)     # calculating transferfunction 1
        H_db,           Rounded_H_db            = Calc_Transferfunction_db(Rounded, dfCal, Y, X)    # calculating transferfunction db
        return  ( 
                Voltage_Ue,     Rounded_Voltage_Ue,
                Voltage_Ua,     Rounded_Voltage_Ua,
                Current,        Rounded_Current,
                Frequenzy,      Rounded_Frequeny,
                PhaseOffset,    Rounded_PhaseOffset,
                Impedance_abs,  Rounded_Impedance_abs,
                Resistance,     Rounded_Resistance,
                Blind,          Rounded_Blind,
                Impedance,      Rounded_Impedance,
                H,              Rounded_H,
                H_db,           Rounded_H_db
                )

    def Save_All(
                dfCal, dfCalRounded, X, Y,
                Voltage_Ue,     Rounded_Voltage_Ue,
                Voltage_Ua,     Rounded_Voltage_Ua,
                Current,        Rounded_Current,
                Frequenzy,      Rounded_Frequeny,
                PhaseOffset,    Rounded_PhaseOffset,
                Impedance_abs,  Rounded_Impedance_abs,
                Resistance,     Rounded_Resistance,
                Blind,          Rounded_Blind,
                Impedance,      Rounded_Impedance,
                H,              Rounded_H,
                H_db,           Rounded_H_db
                ):

        dfCal, dfCalRounded = Save_Voltage_Ue           (dfCal, dfCalRounded, X, Y, Voltage_Ue, Rounded_Voltage_Ue)
        dfCal, dfCalRounded = Save_Voltage_Ua           (dfCal, dfCalRounded, X, Y, Voltage_Ua, Rounded_Voltage_Ua)
        dfCal, dfCalRounded = Save_Current              (dfCal, dfCalRounded, X, Y, Current, Rounded_Current)
        dfCal, dfCalRounded = Save_Frequenzy            (dfCal, dfCalRounded, X, Y, Frequenzy, Rounded_Frequeny)
        dfCal, dfCalRounded = Save_PhaseOffset          (dfCal, dfCalRounded, X, Y, PhaseOffset, Rounded_PhaseOffset)
        dfCal, dfCalRounded = Save_Impedance            (dfCal, dfCalRounded, X, Y, Impedance_abs, Rounded_Impedance_abs)
        dfCal, dfCalRounded = Save_Resistance           (dfCal, dfCalRounded, X, Y, Resistance, Rounded_Resistance)
        dfCal, dfCalRounded = Save_Blind                (dfCal, dfCalRounded, X, Y, Blind, Rounded_Blind)
        dfCal, dfCalRounded = Save_Impedanz_Complex     (dfCal, dfCalRounded, X, Y, Impedance, Rounded_Impedance)
        dfCal, dfCalRounded = Save_Transferfunction_1   (dfCal, dfCalRounded, X, Y, H, Rounded_H)
        dfCal, dfCalRounded = Save_Transferfunction_db  (dfCal, dfCalRounded, X, Y, H_db, Rounded_H_db)
        return dfCal, dfCalRounded

if(True): # Functions Layer 3
    # Functions for Reading Values from Dataframe and Calculating stuff
    def Read_Voltage_Ue(Rounded, dfCal, Y, X):
        Voltage_Ue = dfCal.iloc[Y,X]           #reading Voltage Ue
        Rounded_Voltage_Ue = Round_Sig(Voltage_Ue, Rounded)
        return Voltage_Ue, Rounded_Voltage_Ue
    def Read_Voltage_Ua(Rounded, dfCal, Y, X):
        Voltage_Ua = dfCal.iloc[Y,X+1]           #reading Voltage UA
        Rounded_Voltage_Ua =    Round_Sig(Voltage_Ua, Rounded)
        return Voltage_Ua, Rounded_Voltage_Ua
    def Read_Current(Rounded, dfCal, Y, X):
        Current = dfCal.iloc[Y,X+2]         #reading Current
        Rounded_Current =       Round_Sig(Current, Rounded)
        return Current, Rounded_Current
    def Read_Frequenzy(Rounded, dfCal, Y, X):
        Frequenzy = dfCal.iloc[Y,X+3]       #reading Frequency
        Rounded_Frequeny =      Round_Sig(Frequenzy, Rounded)
        return Frequenzy, Rounded_Frequeny
    def Read_PhaseOffset(Rounded, dfCal, Y, X):
        PhaseOffset = dfCal.iloc[Y, X+4]    #reading Phase Offset
        Rounded_PhaseOffset =   Round_Sig(PhaseOffset, Rounded)
        return PhaseOffset, Rounded_PhaseOffset
    def Calc_Impedance(Rounded, dfCal, Y, X):
        Voltage_Ue, Rounded_Voltage_Ue = Read_Voltage_Ue(Rounded, dfCal, Y, X) #reading Voltage
        Current, Rounded_Current = Read_Current(Rounded, dfCal, Y, X) #reading Current

        Impedance_abs = Voltage_Ue / Current                                   #Calculating Impedance in Ohm
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
    def Calc_Transferfunction_1(Rounded, dfCal, Y, X):
        Voltage_Ue, Rounded_Voltage_Ue = Read_Voltage_Ue(Rounded, dfCal, Y, X)
        Voltage_Ua, Rounded_Voltage_Ua = Read_Voltage_Ua(Rounded, dfCal, Y, X)

        H = Voltage_Ua / Voltage_Ue
        Rounded_H = Round_Sig(Voltage_Ua / Voltage_Ue, Rounded)   
        return H, Rounded_H
    def Calc_Transferfunction_db(Rounded, dfCal, Y, X):
        Voltage_Ue, Rounded_Voltage_Ue = Read_Voltage_Ue(Rounded, dfCal, Y, X)
        Voltage_Ua, Rounded_Voltage_Ua = Read_Voltage_Ua(Rounded, dfCal, Y, X)

        H = Voltage_Ua / Voltage_Ue
        Rounded_H = Round_Sig(Voltage_Ua / Voltage_Ue, Rounded)

        H_db = 20 * math.log10(abs(H))
        Rounded_H_db = Round_Sig(H_db, Rounded)
        return H_db, Rounded_H_db

    # Functions for Saving Values to Dataframe
    def Save_Voltage_Ue(dfCal, dfCalRounded, X, Y, Voltage_Ue, Rounded_Voltage_Ue):
        dfCal.iloc[Y,X]   = Voltage_Ue         #Storing Voltage Ue in V
        dfCalRounded.iloc[Y,X]   = Rounded_Voltage_Ue
        return dfCal, dfCalRounded
    def Save_Voltage_Ua(dfCal, dfCalRounded, X, Y, Voltage_Ua, Rounded_Voltage_Ua):
        dfCal.iloc[Y,X+1]   = Voltage_Ua         #Storing Voltage Ua in V
        dfCalRounded.iloc[Y,X+1]   = Rounded_Voltage_Ua
        return dfCal, dfCalRounded
    def Save_Current(dfCal, dfCalRounded, X, Y, Current, Rounded_Current):
        dfCal.iloc[Y,X+2] = Current         #Storing Current in A
        dfCalRounded.iloc[Y,X+2] = Rounded_Current
        return dfCal, dfCalRounded
    def Save_Frequenzy(dfCal, dfCalRounded, X, Y, Frequenzy, Rounded_Frequeny):
        dfCal.iloc[Y,X+3] = Frequenzy
        dfCalRounded.iloc[Y,X+3] = Rounded_Frequeny
        return dfCal, dfCalRounded
    def Save_PhaseOffset(dfCal, dfCalRounded, X, Y, PhaseOffset, Rounded_PhaseOffset):
        dfCal.iloc[Y, X+4] = PhaseOffset
        dfCalRounded.iloc[Y, X+4] = Rounded_PhaseOffset
        return dfCal, dfCalRounded
    def Save_Impedance(dfCal, dfCalRounded, X, Y, Impedance, Rounded_Impedance):
        dfCal.iloc[Y,X+5] = Impedance
        dfCalRounded.iloc[Y,X+5] = Rounded_Impedance
        return dfCal, dfCalRounded
    def Save_Resistance(dfCal, dfCalRounded, X, Y, Resistance, Rounded_Resistance):
        dfCal.iloc[Y,X+7] = Resistance
        dfCalRounded.iloc[Y,X+7] = Rounded_Resistance
        return dfCal, dfCalRounded
    def Save_Blind(dfCal, dfCalRounded, X, Y, Blind, Rounded_Blind):
        dfCal.iloc[Y,X+8] = Blind
        dfCalRounded.iloc[Y,X+8] = Rounded_Blind
        return dfCal, dfCalRounded
    def Save_Impedanz_Complex(dfCal, dfCalRounded, X, Y, Impedance, Rounded_Impedance):
        dfCal.iloc[Y,X+6] = Impedance
        dfCalRounded.iloc[Y,X+6] = Rounded_Impedance
        return dfCal, dfCalRounded
    def Save_Transferfunction_1(dfCal, dfCalRounded, X, Y, H, Rounded_H):
        dfCal.iloc[Y,X+9] = H
        dfCalRounded.iloc[Y,X+9] = Rounded_H
        return dfCal, dfCalRounded
    def Save_Transferfunction_db(dfCal, dfCalRounded, X, Y, H_db, Rounded_H_db):
        dfCal.iloc[Y,X+10] = H_db
        dfCalRounded.iloc[Y,X+10] = Rounded_H_db
        return dfCal, dfCalRounded
if(True): # Functions Layer 4
    def Round_Sig(x, Rounded):
        if x == 0:
            return 0
        return round(x, Rounded - 1 - int(math.floor(math.log10(abs(x)))))