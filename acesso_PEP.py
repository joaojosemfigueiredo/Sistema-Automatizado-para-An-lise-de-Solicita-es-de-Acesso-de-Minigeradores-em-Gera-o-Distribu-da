
#import easygui
import shutil
import os
import time
from datetime import datetime
from openpyxl import load_workbook
from zipfile import ZipFile
from selenium import webdriver
from selenium.webdriver.common.by import By

    
def download_wait(Pasta_dia):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(Pasta_dia):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds
    
