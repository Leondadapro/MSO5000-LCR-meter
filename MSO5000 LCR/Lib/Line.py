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
import shutil
from enum import Enum

if (True):  #define Paths
    Lib_Dir         = os.path.dirname(os.path.abspath(__file__))
    Base_Dir        = os.path.dirname(Lib_Dir)
    Settings_Path   = os.path.join(Base_Dir, "Settings")
    Data_Path       = os.path.join(Base_Dir, "Data")

columns, rows = shutil.get_terminal_size()
# print(f"Your CMD is {columns} characters wide and {rows} lines tall.")

if(True): #Functions Layer 1
    def Print():
        print("-" * columns)

# if(True): #Functions Layer 2
#      print(1)

# if(True): #Functions Layer 3
#      print(1)