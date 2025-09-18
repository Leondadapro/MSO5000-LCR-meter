import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import math
import pyvisa

Rounded = 2  #Decimal places for rounding

Voltage =       0  #Voltage for Graphing
Current =       1  #Current for Graphing
Frequency =     2  #Frequency for Graphing
PhaseOffset =   3  #Phase Offset for Graphing
Impedance_abs = 4  #Impedance for Graphing
Impedance =     5  #Complex Impedance for Graphing
Resistance =    6  #Resistance for Graphing
Blind =         7  #Blindwiderstand for Graphing

def ImpedanceCalculation():
    dfCal = pd.read_excel('Clean.xlsx') #Cleaned data from MSO5000
    dfCalRounded = dfCal.copy()         #Copy of dataframe for rounded values
    Xmax = dfCal.shape[1]       #Number of columns
    Ymax = dfCal.shape[0] - 1   #Number of rows
    X = 0
    Y = 0
    print("Xmax =", Xmax)
    print("Ymax =", Ymax)
    while (Y <= Ymax):                      #Calculating Impedance, Resistance and Blind for each measurement point
        Voltage = dfCal.iloc[Y,X]           #reading Voltage
        Current = dfCal.iloc[Y,X+1] / 1000  #reading Current Convert mA to A
        Frequenzy = dfCal.iloc[Y,X+2]       #reading Frequency
        PhaseOffset = dfCal.iloc[Y, X+3]    #reading Phase Offset

        Impedance_abs = Voltage / Current   #Calculating Impedance in Ohm
        Resistance  = math.cos(math.radians(PhaseOffset)) * Impedance_abs   #Calculating Resistance in Ohm
        Blind       = math.sin(math.radians(PhaseOffset)) * Impedance_abs   #Calculating Blindwiderstand in Ohm
        Impedance   = complex(Resistance, Blind)                            #Calculating Complex Impedance in Ohm

        dfCal.iloc[Y,X+4] = Impedance_abs   #Storing Impedance in Ohm
        dfCal.iloc[Y,X+5] = Impedance       #Complex Impedance in Ohm
        dfCal.iloc[Y,X+6] = Resistance      #Storing Resistance in Ohm
        dfCal.iloc[Y,X+7] = Blind           #Storing Blindwiderstand in Ohm

        Rounded_Voltage =       round(Voltage, Rounded)                                 #Rounding Voltage to Rounded decimal places
        Rounded_Current =       round(Current * 1000, Rounded)                          #Rounding Current to Rounded decimal places and convert A to mA
        Rounded_Frequeny =      round(Frequenzy, Rounded)                               #Rounding Frequency to Rounded decimal places
        Rounded_PhaseOffset =   round(PhaseOffset, Rounded)                             #Rounding Phase Offset to Rounded decimal places
        Rounded_Impedance_abs = round(Impedance_abs, Rounded)                           #Rounding Impedance to Rounded decimal places
        Rounded_Impedance =     round(Resistance, Rounded) + round(Blind, Rounded)*1j   #Rounding Complex Impedance to Rounded decimal places
        Rounded_Resistance =    round(Resistance, Rounded)                              #Rounding Resistance to Rounded decimal places
        Rounded_Blind =         round(Blind, Rounded)                                   #Rounding Blindwiderstand to Rounded decimal places

        dfCalRounded.iloc[Y,X]   = Rounded_Voltage
        dfCalRounded.iloc[Y,X+1] = Rounded_Current
        dfCalRounded.iloc[Y,X+2] = Rounded_Frequeny
        dfCalRounded.iloc[Y,X+3] = Rounded_PhaseOffset
        dfCalRounded.iloc[Y,X+4] = Rounded_Impedance_abs
        dfCalRounded.iloc[Y,X+5] = Rounded_Impedance
        dfCalRounded.iloc[Y,X+6] = Rounded_Resistance
        dfCalRounded.iloc[Y,X+7] = Rounded_Blind

        Y += 1  #Next Row
    dfCal.to_excel("Clean_Calc.xlsx", index = False)                    #Saving new data to new file
    dfCalRounded.to_excel("Clean_Calc_Rounded.xlsx", index = False)     #Saving new data to new file


# ImpedanceCalculation()  #Calculate Impedance
# df = pd.read_excel('Clean_Calc_Rounded.xlsx')   #Cleaned, Calculated and Rounded data from MSO5000
# print(df)
# plt.plot(df.iloc[:,Frequency], df.iloc[:,Impedance_abs], label="Impedance")
# plt.show()
while True:

    print("Hello and Welcome to the MSO5000 LCR Measurement Tool\n"
          "This tool helps you to measure and analyze LCR components with the MSO5000\n\n\n")

    print("What do u wanna do? (Pick from List)\n\n")
    print("1: Measure LCR Component\n"
          "2: Analyze existing Measurement\n"
          "3: Exit Program\n\n")
    n = input("Your Input: ")
    print("\nYour Input: ",n)
    
    match n:
        case "1":
            print("sus")
            ImpedanceCalculation()
        case "2":
            print("sus")
        case "3":
            print("sus")