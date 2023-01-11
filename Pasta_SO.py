
#import easygui
import shutil
import os
import time
from datetime import datetime
from openpyxl import load_workbook
from zipfile import ZipFile
from selenium import webdriver
from selenium.webdriver.common.by import By


def Pasta_SO(numero_so, df_alim_por_reg, alim):
    Regional = df_alim_por_reg.loc[(df_alim_por_reg.Alimentador==alim[0], 'REG')].values[0]
    alimentador = df_alim_por_reg.loc[(df_alim_por_reg.Alimentador==alim[0], 'Alimentador')].values[0]
    
    Pasta_mae =('C:\\Users\\joaof\\OneDrive\\Área de Trabalho\\TCC\\programa atualizado\\SO\\' + Regional)
    Pasta_alimentador_SO = (Pasta_mae + '\\' + alimentador + '\\SO '+ str(numero_so))
    try:
        os.makedirs(Pasta_alimentador_SO)
    except:
        print('Essa Pasta Já Existe.')
        # continue
    else:
        print('A Pasta Foi Criada.')
        # continue    
    return (Pasta_alimentador_SO)





