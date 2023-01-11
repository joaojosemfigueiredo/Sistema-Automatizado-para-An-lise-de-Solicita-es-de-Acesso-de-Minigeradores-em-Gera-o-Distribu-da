
import os
import time
import os.path
from datetime import datetime
import easygui as sg
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select



alimentador = 'ACA01'
sigla = alimentador[:3]
sigla

# def download(alimentador, sigla):
Pasta_mae ='C:\\Users\\joaof\\OneDrive\\Área de Trabalho\\TCC\\programa atualizado\\Alimentadores\\' + sigla
Pasta_alimentador = Pasta_mae + '\\' + alimentador

try:
    os.makedirs(Pasta_alimentador)
except:
    print('Essa Pasta Já Existe.')
    # continue
else:
    print('A Pasta Foi Criada.')
    # continue
