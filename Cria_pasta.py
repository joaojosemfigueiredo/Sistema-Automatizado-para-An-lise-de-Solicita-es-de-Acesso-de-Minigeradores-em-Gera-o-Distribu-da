#import easygui
import shutil
import os
import time
from datetime import datetime
from openpyxl import load_workbook
from zipfile import ZipFile
from selenium import webdriver
from selenium.webdriver.common.by import By


alimentador = 'ACA01'
sigla = alimentador[:3]
sigla

Pasta_mae ='C:\\Users\\joaof\OneDrive\\Área de Trabalho\\TCC\\programa atualizado\\Alimentadores\\' + sigla
Pasta_alimentador = Pasta_mae + '\\' + alimentador

try:
    os.makedirs(Pasta_alimentador)
except:
    print('Essa Pasta Já Existe.')
    # continue
else:
    print('A Pasta Foi Criada.')
    # continue    
    
# foobars = os.listdir(Pasta_alimentador)
