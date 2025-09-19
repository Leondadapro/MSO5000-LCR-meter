from calendar import c
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
from enum import Enum

Rounded     = 2     #Decimal places for rounding
Time_Delay  = 1     #Time Delay for better UX
Repeat      = 0     #Variable for repeating loops

class State(Enum):
    # All of the text dialog variables
    Start_Text = 1
    Pick_Text1 = 2
    Pick_Text2 = 3
    Pick_Text3 = 4
    Pick_Text_Setting = 5

    # Random ah constants
    Voltage =       0  #Voltage for Graphing
    Current =       1  #Current for Graphing
    Frequency =     2  #Frequency for Graphing
    PhaseOffset =   3  #Phase Offset for Graphing
    Impedance_abs = 4  #Impedance for Graphing
    Impedance =     5  #Complex Impedance for Graphing
    Resistance =    6  #Resistance for Graphing
    Blind =         7  #Blindwiderstand for Graphing

    # Random ah variables

def Call_Settings(Type):    #Calling Settings from Excel File
    if   ((Type == "Custom") or (Type == "RND")):
        dfSettings = pd.read_excel('Settings_Custom.xlsx', header = None) #Custom Settings from Excel File

    elif (Type == "Default"):
        dfSettings = pd.read_excel('Settings_Default.xlsx', header = None) #Default Settings from Excel File


    Rounded = int(dfSettings.iloc[0,1])      #Decimal places for rounding
    Time_Delay = dfSettings.iloc[1,1]   #Time Delay for better UX in seconds

    if (Type == "RND"):
        print("\nRounded Decimal Places\t\t: ", Rounded)
        print("\nTime Delay when going back\t: ", Time_Delay)

    return Rounded, Time_Delay

def Save_Settings(Rounded, Time_Delay):
    dfSettings = pd.read_excel('Settings_Custom.xlsx', header = None)
    dfSettings.iloc[0,1] = int(Rounded)      #Decimal places for rounding
    dfSettings.iloc[1,1] = float(Time_Delay)   #Time Delay for better UX in seconds
    dfSettings.to_excel("Settings_Custom.xlsx", index=False, header=False) #Saving Settings to Excel File

def Impedance_Calculation():    #Calculating Impedance, Resistance and Blindwiderstand from Voltage, Current, Frequency and Phase Offset
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

def TXT_Dialog(n):                   #All of the text dialog stuff
    match n:
        case State.Start_Text:
            print(  "Hello and Welcome to the MSO5000 LCR Measurement Tool\n"
                    "This tool helps you to measure and analyze LCR components with the MSO5000\n\n\n")

        case State.Pick_Text1:
            print(  "What do u wanna do? (Pick from List)\n\n")
            print(  "1 : Measure LCR Component\n"
                    "2 : Analyze / Calculate existing Measurement\n"
                    "3 : Settings\n"
                    "99: Exit Program\n\n")

        case State.Pick_Text2:
            print(  "What do u wanna do? (Pick from List)\n\n")
            print(  "1 : Calculate Data and export as Excel Files\n"
                    "2 : Plot Data\n"
                    "3 : Both Calculate and Plot Data\n"
                    "99: Go back\n\n")

        case State.Pick_Text3:
            print(  "Settings Menu\n\n"
                    "1 : Load Default Settings\n"
                    "2 : Load Custom Settings\n"
                    "3 : Show Current Settings\n"
                    "4 : Change Settings\n"
                    "5 : Save Current Settings as Custom Settings\n"
                    "99: Go back\n\n")

        case State.Pick_Text_Setting:
            print(  "What Settings do you wanna change?\n\n"
                    "1 : Decimal places for rounding (Current: " + str(Rounded) + ")\n"
                    "2 : Time Delay when going back (Current: " + str(Time_Delay) + "s)\n"
                    "99: Go back\n\n")

def Clear_CLI():
    print("\033[2J\033[H", end='')  #Clear screen + move cursor to top-left

Rounded, Time_Delay = Call_Settings("Custom")  #Calling Custom Settings at the start of the program
Clear_CLI()

while True:
    Clear_CLI()
    TXT_Dialog(State.Start_Text)    #Starting Text
    TXT_Dialog(State.Pick_Text1)    #pick from list text

    n = input("Your Input: ")       #User Input
    
    match n:
        case "1":   #Measure LCR Component (WIP)

            Clear_CLI()
            print("Measure LCR Component")
            
        case "2":   #Analyze / Calculate existing Measurement

            Clear_CLI()
            print("Analyze / Calculate existing Measurement")

            Repeat = 1
            while (Repeat == 1):
                Clear_CLI()
                TXT_Dialog(State.Pick_Text2)            #pick from list text
                User_Input = input("Your Input: ")      #User Input
                
                match User_Input:

                    case "1": #Calculate Data and export as Excel Files
                        Clear_CLI()
                        print("Calculate Data and export as Excel Files")
                        Impedance_Calculation()

                    case "2": #Plot Data (WIP)
                        Clear_CLI()
                        print("Plot Data")

                    case "3": #Both Calculate and Plot Data (WIP)
                        Clear_CLI()
                        print("Both Calculate and Plot Data")

                    case "99": #Go back
                        Clear_CLI()
                        Repeat = 0
                        print("Exit to Main Menu")
                        time.sleep(Time_Delay)

        case "3":   #Settings
            Clear_CLI()
            Repeat = 1
            while (Repeat == 1):
                Clear_CLI()
                TXT_Dialog(State.Pick_Text3)
                User_Input = input("Your Input: ")

                match User_Input:
                    case "1":   #Load Default Settings
                        Clear_CLI()
                        print("Loaded Default Settings")
                        Rounded, Time_Delay = Call_Settings("Default")
                        time.sleep(Time_Delay)

                    case "2":   #Load Custom Settings
                        Clear_CLI()
                        print("Loaded Custom Settings")
                        Rounded, Time_Delay = Call_Settings("Custom")
                        time.sleep(Time_Delay)

                    case "3":   #Show Current Settings
                        Clear_CLI()
                        Rounded, Time_Delay = Call_Settings("RND") #Calling current settings
                        time.sleep(5)

                    case "4":   #Change Settings
                        again = 1
                        while (again == 1):
                            again = 0
                            Clear_CLI()
                            print("Change Settings")
                            TXT_Dialog(State.Pick_Text_Setting)
                            User_Input = input("Your Input: ")

                            match User_Input:
                                case "1":   #Change Decimal Places
                                    Clear_CLI()
                                    print("Change Decimal Places")
                                    New_Value = input("Enter new value (Current: " + str(Rounded) + "): ")
                                    Rounded = int(New_Value)
                                    print("Decimal Places changed to: ", Rounded)
                                    print("\n\ndo you wana change more settings?")
                                    again = int(input("2 = yes / 1 = no: ")) - 1

                                case "2":   #Change Time Delay
                                    Clear_CLI()
                                    print("Change Time Delay")
                                    New_Value = input("Enter new value (Current: " + str(Time_Delay) + "s): ")
                                    Time_Delay = float(New_Value)
                                    print("Time Delay changed to: ", Time_Delay, "s")
                                    print("\n\ndo you wana change more settings?")
                                    again = int(input("2 = yes / 1 = no: ")) - 1

                                case "99":
                                    Clear_CLI()
                                    again = 0
                                    print("Go back")
                                    time.sleep(Time_Delay)

                    case "5":   #Save Current Settings as Custom Settings
                        Clear_CLI()
                        print("Save Current Settings as Custom Settings")
                        Save_Settings(Rounded, Time_Delay)
                        print("\nCurrent Settings saved as Custom Settings")
                        time.sleep(Time_Delay)

                    case "99":
                        Clear_CLI()
                        Repeat = 0
                        print("Exit to Main Menu")
                        time.sleep(Time_Delay)

        case "99":
            Clear_CLI()
            print("Exit Program")
            time.sleep(Time_Delay) #Short Delay for better UX
            sys.exit()#Exit Program

