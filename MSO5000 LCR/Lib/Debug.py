#Lib for all the Dialog stuff

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
import Lib.Line as line

def Wait_for_keypress():          # Wait for a keypress
    print("\nPress anything to continou")
    msvcrt.getch()

if (True):  #define Paths
    Lib_Dir         = os.path.dirname(os.path.abspath(__file__))
    Base_Dir        = os.path.dirname(Lib_Dir)
    Settings_Path   = os.path.join(Base_Dir, "Settings")
    Data_Path       = os.path.join(Base_Dir, "Data")

if(True): #Functions Layer 1
    def printDir(Debug):
        if(Debug == "yes"):
            line.Print()
            print("All of the important Paths:\n")
            print("Base Dir = \t\t", Base_Dir)
            print("Lib Dir = \t\t", Lib_Dir)
            print("Settings Path = \t", Settings_Path)
            print("Data Path = \t\t", Data_Path)
            print("\b")
            print("End of Debug Paths")
            line.Print()
            Wait_for_keypress()

    def Calculation(Debug, What):
        print(1)

# if(True): #Functions Layer 2
#     print(1)

# if(True): #Functions Layer 3
#     print(1)