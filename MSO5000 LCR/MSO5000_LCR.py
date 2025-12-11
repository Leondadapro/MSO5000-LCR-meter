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
import Lib.Calculation as c
import Lib.Settings as s
import Lib.Debug as d
import Lib.Misc as m

Rounded     = 0     #Decimal places for rounding
Time_Delay  = 0     #Time Delay for better UX
Debug       = 0     #Debug Variable
Repeat      = 0     #Variable for repeating loops

class S(Enum):
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

if (True):
    Base_Dir = os.path.dirname(os.path.abspath(__file__))
    Settings_Path = os.path.join(Base_Dir, "Settings")
    Data_Path = os.path.join(Base_Dir, "Data")

def TXT_Dialog(n):                  # All of the text dialog stuff
    match n:
        case S.Start_Text:          #Starting Text
            print(  "Hello and Welcome to the MSO5000 LCR Measurement Tool\n"
                    "This tool helps you to measure and analyze LCR components with the MSO5000\n\n\n")

        case S.Pick_Text1:          # Main Menu
            print(  "What do u wanna do? (Pick from List)\n\n")
            print(  "1 : Measure LCR Component\n"
                    "2 : Analyze / Calculate existing Measurement\n"
                    "3 : Settings\n"
                    "99: Exit Program\n\n")

        case S.Pick_Text2:          # Analyze / Calculate existing Measurement Menu
            print(  "What do u wanna do? (Pick from List)\n\n")
            print(  "1 : Calculate Data and export as Excel Files\n"
                    "2 : Plot Data\n"
                    "3 : Both Calculate and Plot Data\n"
                    "99: Go back\n\n")

        case S.Pick_Text3:          # Settings Menu
            print(  "Settings Menu\n\n"
                    "1 : Load Default Settings\n"
                    "2 : Load Custom Settings\n"
                    "3 : Show Current Settings\n"
                    "4 : Change Settings\n"
                    "5 : Save Current Settings as Custom Settings\n"
                    "99: Go back\n\n")

        case S.Pick_Text_Setting:   # Settings changing menu
            print(  "What Settings do you wanna change?\n\n"
                    "1 : Decimal places for rounding (Current: " + str(Rounded) + ")\n"
                    "2 : Time Delay when going back (Current: " + str(Time_Delay) + "s)\n"
                    "3 : Debug Messages (Current: " + str(Debug) + ")\n"
                    "99: Go back\n\n")

def Clear_CLI():                    # Clear screen + move cursor to top-left
    print("\033[2J\033[H", end='')
def Wait_for_keypress():          # Wait for a keypress
    print("\nPress anything to continou")
    msvcrt.getch()

Rounded, Time_Delay, Debug = s.Settings("Custom", "Load", 0) #Load Current Settings
Clear_CLI()

d.printDir(Debug)
m.createExcel(0, 0)

while True: # Main Loop
    Clear_CLI()
    TXT_Dialog(S.Start_Text)    #Starting Text
    TXT_Dialog(S.Pick_Text1)    #pick from list text

    n = input("Your Input: ")       #User Input
    
    match n:
        case "1":   # Measure LCR Component (WIP)

            Clear_CLI()
            print("Measure LCR Component")
            
        case "2":   # Analyze / Calculate existing Measurement

            Clear_CLI()
            print("Analyze / Calculate existing Measurement")

            Repeat = 1
            while (Repeat == 1):
                Clear_CLI()
                TXT_Dialog(S.Pick_Text2)            #pick from list text
                User_Input = input("Your Input: ")      #User Input
                
                match User_Input:

                    case "1": #Calculate Data and export as Excel Files
                        Clear_CLI()
                        print("Calculate Data and export as Excel Files")
                        c.Impedance_Calculation(Rounded, Debug)

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

        case "3":   # Settings
            Clear_CLI()
            Repeat = 1
            while (Repeat == 1):
                Clear_CLI()
                TXT_Dialog(S.Pick_Text3)
                User_Input = input("Your Input: ")

                match User_Input:
                    case "1":   #Load Default Settings
                        Clear_CLI()
                        print("Loaded Default Settings")

                        Rounded, Time_Delay, Debug = s.Settings("Default", "Load", 0)

                        time.sleep(Time_Delay)

                    case "2":   #Load Custom Settings
                        Clear_CLI()
                        print("Loaded Custom Settings")

                        Rounded, Time_Delay, Debug = s.Settings("Custom", "Load", 0)

                        time.sleep(Time_Delay)

                    case "3":   #Show Current Settings
                        Clear_CLI()
                        print("Current Settings:\n\n")

                        s.Settings("Current", "Show", 0)
                        Wait_for_keypress()

                    case "4":   #Change Settings
                        file_path = os.path.join(Settings_Path, "Settings_Current.xlsx")
                        dfData = pd.read_excel(file_path, header=None, index_col=None)
                        again = 1
                        while (again == 1):
                            again = 0
                            Clear_CLI()
                            print("Change Settings")
                            TXT_Dialog(S.Pick_Text_Setting)
                            User_Input = input("Your Input: ")

                            match User_Input:
                                case "1":   #Change Decimal Places
                                    Clear_CLI()
                                    print("Change Decimal Places")
                                    New_Value = input("Enter new value (Current: " + str(Rounded) + "): ")
                                    Rounded = int(New_Value)
                                    dfData.iloc[0,1] = Rounded
                                    print("Decimal Places changed to: ", Rounded)
                                    print("\n\ndo you wana change more settings?")
                                    again = int(input("2 = yes / 1 = no: ")) - 1

                                case "2":   #Change Time Delay
                                    Clear_CLI()
                                    print("Change Time Delay")
                                    New_Value = input("Enter new value (Current: " + str(Time_Delay) + "s): ")
                                    Time_Delay = float(New_Value)
                                    dfData.iloc[1,1] = Time_Delay
                                    print("Time Delay changed to: ", Time_Delay, "s")
                                    print("\n\ndo you wana change more settings?")
                                    again = int(input("2 = yes / 1 = no: ")) - 1

                                case "3":   #Debug Messages
                                    Clear_CLI()
                                    print("Do u want Debug Messages")
                                    New_Value = input("Enter new value (Current: " + str(Debug) + ") (yes/no): ")
                                    if (New_Value == "yes") or (New_Value == "no"):
                                        Debug = str(New_Value)
                                        dfData.iloc[2,1] = Debug
                                        print("Debug Messages changed to: ", Debug)
                                    print("\n\ndo you wana change more Settings?")
                                    again = int(input("2 = yes / 1 = no: ")) - 1

                                case "99":
                                    Clear_CLI()
                                    again = 0
                                    print("Go back")
                                    time.sleep(Time_Delay)
                        s.Settings ("Current", "Save", dfData)

                    case "5":   #Save Current Settings as Custom Settings
                        Clear_CLI()
                        print("Save Current Settings as Custom Settings")
                        file_path = os.path.join(Settings_Path, "Settings_Current.xlsx")
                        dfData = pd.read_excel(file_path, header=None, index_col=None)
                        s.Settings("Custom", "Save", dfData)
                        print("\nCurrent Settings saved as Custom Settings")
                        time.sleep(Time_Delay)

                    case "99":
                        Clear_CLI()
                        Repeat = 0
                        print("Exit to Main Menu")
                        time.sleep(Time_Delay)

        case "99":  # Exit Program
            Clear_CLI()
            print("Exit Program")
            time.sleep(Time_Delay)  # Short Delay for better UX
            sys.exit()