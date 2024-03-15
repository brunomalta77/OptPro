import pandas as pd
import numpy as np
import re
import time
import string
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns
import joblib
import glob
import os
import time
from itertools import cycle
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, confusion_matrix, classification_report
from tqdm import tqdm
import MeCab 
import subprocess
import sudachipy
from sudachipy import dictionary
from tqdm import tqdm
import joblib
import os 



df = pd.read_excel(r"C:\Users\BrunoMalta\OneDrive - Brand Delta\Desktop\Testes\Japan_BAT\Power Bi\OptPro_Nomad_FiFi_Example.xlsx",sheet_name="Optimization_sample_MP")



# changing
tv_fifi_min = input("Tell me you min value for TV Fifi")
tv_fifi_max = input("Tell me you max value for TV Fifi")
#vod_fifi_min = input("Tell me you min value for vod Fifi")
#vod_fifi_min = input("Tell me you max value for vod Fifi")
#display_fifi_min = input("Tell me you min value for display Fifi")
#display_fifi_min = input("Tell me you max value for display Fifi")
#social_fifi_min = input("Tell me you min value for social Fifi")
#social_fifi_min = input("Tell me you max value for social Fifi")



def transforming_data():
    # write here the part of the print the mean st.show(df.VolumeOpt.mean())
    
    
    if tv_fifi_min != "None":
        df.loc[(df.Group =="tv.fifi") & (df.Period =="Partial"),"SpendMin"] = tv_fifi_min
    
    if tv_fifi_max != "None":
        df.loc[(df.Group =="tv.fifi") & (df.Period =="Partial"),"SpendMax"] = tv_fifi_max
    

    writer = pd.ExcelWriter(r"C:\Users\BrunoMalta\OneDrive - Brand Delta\Desktop\Testes\Japan_BAT\Power Bi\test.xlsx",
                                engine='xlsxwriter', options={'strings_to_urls': False})

    df.to_excel(writer, index=False,sheet_name='Optimization_sample_MP')
    
    writer.close()


#Using Pyautogui
import os
import pyautogui
import time

def openFile():    
    path = r"C:\Users\BrunoMalta\OneDrive - Brand Delta\Desktop\Testes\Japan_BAT\Power Bi\test.xlsx"
    path = os.path.realpath(path)
    os.startfile(path)
    time.sleep(1.5)

def fullScreen():    
    pyautogui.hotkey('win', 'up')
    time.sleep(1)

def findAndClickRibbon():    
    pyautogui.moveTo(x=400, y=32)
    pyautogui.click()
    time.sleep(1)


def runAddIn():    
    pyautogui.moveTo(x=1385, y=827)
    pyautogui.click()
    time.sleep(1)

def saveFile():    
    pyautogui.hotkey('ctrl', 's')
    time.sleep(1)

def closeFile():    
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)

        
def automate_actions():
    openFile()
    fullScreen()
    findAndClickRibbon()
    runAddIn()
    saveFile()
    closeFile()



def show_test():
    df = pd.read_excel(r"C:\Users\BrunoMalta\OneDrive - Brand Delta\Desktop\Testes\Japan_BAT\Power Bi\test.xlsx",sheet_name="Optimization_sample_MP")
    #st.show(df.VolumeOpt.mean())


if __name__ =="__main__":
    transforming_data()
    automate_actions()


