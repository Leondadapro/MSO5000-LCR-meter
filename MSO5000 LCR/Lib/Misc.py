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



if(True): #Functions Layer 1
    def createExcel(Type, lenght):
        df = pd.DataFrame()
        X = 0

        df.insert(X   , f"Ue_[V]",    pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+1 , f"Ua_[V]",    pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+2 , f"Ie_[A]",    pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+3 , f"F_[Hz]",    pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+4 , f"φ_[°]",     pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+5 , f"|Z|_[Ω]",   pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+6 , f"Z_[Ω]",     pd.Series(np.zeros(lenght, dtype=np.complex128)))
        df.insert(X+7 , f"R_[Ω]",     pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+8 , f"X_[Ω]",     pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+9 , f"H_[]",      pd.Series(np.zeros(lenght, dtype=np.float64)))
        df.insert(X+10, f"H_[db]",    pd.Series(np.zeros(lenght, dtype=np.float64)))

        file_path = os.path.join(Data_Path, "Clean_Test.xlsx")          # Exporting calculated data to Excel File
        df.to_excel(file_path, index = False)


# if(True): #Functions Layer 2
#      print(1)

# if(True): #Functions Layer 3
#      print(1)