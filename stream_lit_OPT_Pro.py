# building a stream lite application
import pandas as pd
import numpy as np
import re
import time
import string
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
import joblib
import glob
import requests
import dash
import pandas as pd
import plotly.express as px 
from plotly.subplots import make_subplots
import numpy as np 
import plotly.graph_objects as go
import os 
from PIL import Image
import os
import pyautogui
import time


#page config
st.set_page_config(page_title="OptPro",page_icon="🎩",layout="wide")


df = pd.read_excel(r"C:\Users\BrunoMalta\OneDrive - Brand Delta\Desktop\Testes\Japan_BAT\Power Bi\OptPro_Nomad_FiFi_Example.xlsx",sheet_name="Optimization_sample_MP")



# changing
tv_fifi_min = st.number_input("Tell me you min value for TV Fifi")
tv_fifi_max = st.number_input("Tell me you max value for TV Fifi")
#vod_fifi_min = input("Tell me you min value for vod Fifi")
#vod_fifi_min = input("Tell me you max value for vod Fifi")
#display_fifi_min = input("Tell me you min value for display Fifi")
#display_fifi_min = input("Tell me you max value for display Fifi")
#social_fifi_min = input("Tell me you min value for social Fifi")
#social_fifi_min = input("Tell me you max value for social Fifi")



def transforming_data():
    # write here the part of the print the mean st.show(df.VolumeOpt.mean())
    st.write("old value:")
    st.write(df.VolumeOpt.mean())
    
    if tv_fifi_min != "None":
        df.loc[(df.Group =="tv.fifi") & (df.Period =="Partial"),"SpendMin"] = tv_fifi_min
    
    if tv_fifi_max != "None":
        df.loc[(df.Group =="tv.fifi") & (df.Period =="Partial"),"SpendMax"] = tv_fifi_max
    

    writer = pd.ExcelWriter(r"C:\Users\BrunoMalta\OneDrive - Brand Delta\Desktop\Testes\Japan_BAT\Power Bi\test.xlsx",
                                engine='xlsxwriter', options={'strings_to_urls': False})

    df.to_excel(writer, index=False,sheet_name='Optimization_sample_MP')
    
    writer.close()



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
    try:
        df = pd.read_excel(r"C:\Users\BrunoMalta\OneDrive - Brand Delta\Desktop\Testes\Japan_BAT\Power Bi\test.xlsx",sheet_name="Optimization_sample_MP")
        st.write("new value:")
        st.write(df.VolumeOpt.mean())
    except:
        st.write("There still no test")

if __name__ =="__main__":
    transforming_data()
    if st.button("update data"):
        automate_actions()
        show_test()
    else:
        st.write("Data has not been updated.")