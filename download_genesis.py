import os
import time
import os.path
import pandas as pd
from datetime import datetime
import easygui as sg
import numpy as np
from pandas import Series
import pandapower as pp
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select


def download_genesis(alimentador, sigla, Pasta_mae, Pasta_alimentador):
    try:
        os.makedirs(Pasta_alimentador)
    except:
        print('Essa Pasta Já Existe.')
    else:
        print('A Pasta Foi Criada.')
