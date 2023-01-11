import sys
import docx
import os
import time
import os.path
import pandas as pd
from datetime import datetime
import easygui as sg
import re
import numpy as np
from pyproj import Proj
import pyproj
from pandas import Series
import pandapower as pp
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from unidecode import unidecode

#Programas Relacionados
from download_genesis import *
from alimentadores import *
from alimentadorespecifico import *
from word import *
from Pasta_SO import *
from acesso_PEP import *
from init import *
from word2 import *

# utm_x, utm_y, potencia, SO = init()

# SO_automatica =  544219 
# num_so = SO_automatica

# # UTMX, UTMY, potencia = acesso_PEP(num_so)
# UTMX = str(421476.00)
# UTMY = str(7003823.00)

# potencia_conexao = 0.3
# # SO_automatica = 510764
# nome_so = 'nome'

def programa_IA(nome_so, SO_automatica, UTMX, UTMY, potencia_conexao,alimentadorespec):
    # #Encontra a Localização Mais Próxima
    num_so = SO_automatica
    if alimentadorespec == 'ALIMENTADOR MAIS PRÓXIMO':
        alim = ponto_analise(UTMX, UTMY)
    else:
        alim = ponto_analise_alimentador_especifico(UTMX, UTMY,alimentadorespec)
    
    alimentador = alim[0]
    sigla = alimentador[:3]
    sigla
    equip_conectado = str(alim[1])
    
    Pasta_mae ='C:\\Users\\joaof\\OneDrive\\Área de Trabalho\\TCC\\programa atualizado\\Alimentadores\\' + sigla
    Pasta_alimentador = Pasta_mae + '\\' + alimentador
    
    download = download_genesis(alimentador, sigla, Pasta_mae, Pasta_alimentador)
    
    # Dicionário para tradução dos nomes do Gênesis
    dataname = {'Alimentador.Código do Alimentador': 'alim',
                'Trecho Primário Aéreo.Identificador do Objeto':'id',
                'Trecho Primário Subterrâneo.Identificador do Objeto': 'id',
                'Transformador de Distribuição Aéreo.Identificador do Objeto': 'id',
                'Transformador de Distribuição Subterrâneo.Identificador do Objeto': 'id',
                'Banco Capacitor.Identificador do Objeto': 'id',
                'Transformador de Distribuição Aéreo.Número Operacional': 'fu',
                'Banco Capacitor.Número Operacional': 'fu',
                'Transformador de Distribuição Subterrâneo.Número Operacional': 'fu',
                'Transformador de Distribuição Aéreo.Coordenada X': 'coord_x',
                'Transformador de Distribuição Subterrâneo.Coordenada X': 'coord_x',
                'Alimentador.Coordenada X': 'coord_x',
                'Transformador de Distribuição Aéreo.Coordenada Y': 'coord_y',
                'Transformador de Distribuição Subterrâneo.Coordenada Y': 'coord_y',
                'Alimentador.Coordenada Y': 'coord_y',
                'Transformador de Distribuição Aéreo.Potência Nominal (kVA)': 'pot_nom',
                'Transformador de Distribuição Subterrâneo.Potência Nominal (kVA)': 'pot_nom',
                'Banco Capacitor.Potência Nominal': 'pot_nom',
                'Transformador de Distribuição Aéreo.Demanda Total (kVA)': 'pot_dem',
                'Transformador de Distribuição Subterrâneo.Demanda Total (kVA)': 'pot_dem',
                'Transformador de Distribuição Aéreo.Potência instalada GD total (W)': 'pot_ger',
                'Transformador de Distribuição Subterrâneo.Potência instalada GD total (W)': 'pot_ger',
                'Rep. Grafica: Trecho Primário Aéreo.Linha': 'graf',
                'Rep. Grafica: Trecho Primário Subterrâneo.Linha': 'graf',
                'Rep. Grafica: Trecho Primário Aéreo.Rotação': 'rotação',
                'Rep. Grafica: Trecho Primário Subterrâneo.Rotação': 'rotação',
                'Trecho Primário Aéreo.Comprimento Estimado': 'comprimento',
                'Trecho Primário Subterrâneo.Comprimento Estimado': 'comprimento',
                'Trecho Primário Aéreo.Quantidade de Cabos por Fase': 'ncabos',
                'Trecho Primário Subterrâneo.Quantidade de Cabos por Fase': 'ncabos',
                'Trecho Primário Aéreo.Fases Existentes': 'nfases',
                'Trecho Primário Subterrâneo.Fases Existentes': 'nfases',
                'Tipo de Cabo.Código de cabo': 'tipocabo',
                'Tipo de Cabo.Reatância Positiva': 'z1',
                'Tipo de Cabo.Reatância Zero': 'z0',
                'Tipo de Cabo.Resistência Positiva': 'r1',
                'Tipo de Cabo.Resistência Zero': 'r0',
                'Transformador de Distribuição Aéreo.Tensão do TAP ajustado (kV)': 'tap',
                'Transformador de Distribuição Subterrâneo.Tensão do TAP ajustado (kV)': 'tap',
                'Transformador de Distribuição Aéreo.Proprietário': 'proprietario',
                'Transformador de Distribuição Subterrâneo.Proprietário': 'proprietario',
                'Banco Capacitor.Tipo de Banco': 'tipo_banco',
                'Banco Capacitor.Tipo de Controle': 'tipo_controle',
                'Banco Capacitor.Tipo de Ligação': 'tipo_ligacao',
                'Chave Fusível.Identificador do Objeto':'id',
                'Chave Fusível.Número Operacional':'fu',
                'Chave Fusível.Coordenada X':'coord_x',
                'Chave Fusível.Coordenada Y':'coord_y',
                'Chave Fusível.Rotação':'rotacao',
                'Religador.Identificador do Objeto':'id',
                'Religador.Número Operacional':'fu',
                'Religador.Coordenada X':'coord_x',
                'Religador.Coordenada Y':'coord_y',
                'Religador.Rotação':'rotacao',
                'Chave Seccionadora.Identificador do Objeto':'id',
                'Chave Seccionadora.Número Operacional':'fu',
                'Chave Seccionadora.Coordenada X':'coord_x',
                'Chave Seccionadora.Coordenada Y':'coord_y',
                'Chave Seccionadora.Rotação':'rotacao',
                'Banco Regulador de Tensão.Identificador do Objeto':'id',
                'Banco Regulador de Tensão.Número Operacional':'fu',
                'Banco Regulador de Tensão.Coordenada X':'coord_x',
                'Banco Regulador de Tensão.Coordenada Y':'coord_y',
                'Banco Regulador de Tensão.Rotação':'rotacao',
                'Banco Regulador de Tensão.Tensão de Regulação (PU)':'tap',
                'Banco Regulador de Tensão.Nível de Tensão (V)':'tensao',
                'Alimentador.Identificador do Objeto':'id',
                'Alimentador.Tensão Nominal do Alimentador':'tensão'
                }
    
    # Leitura CSV dos Trechos Aéreos
    df_lines_aer = pd.read_csv(Pasta_alimentador +"\\77922.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_lines_aer.rename(columns=dataname, inplace=True)
    # df_lines_aer.dropna(inplace=True)
    
    # Leitura CSV dos Trechos Subterrâneos
    df_lines_sub = pd.read_csv(Pasta_alimentador +"\\77923.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_lines_sub.rename(columns=dataname, inplace=True)
    # df_lines_sub.dropna(inplace=True)
    tamanho = df_lines_sub['graf'][0].count(',')
    
    # Leitura CSV dos Transformadores Aéreos
    df_trafos_aer = pd.read_csv(Pasta_alimentador +"\\77022.csv", skiprows=[0,1,2,3], sep=';', index_col=False, na_filter= False,  encoding='latin_1')
    df_trafos_aer.rename(columns=dataname, inplace=True)
    df_trafos_aer.replace(' ', np.nan, inplace=True)
    
    # Leitura CSV dos Transformadores Subterrâneos
    df_trafos_sub = pd.read_csv(Pasta_alimentador +"\\77122.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_trafos_sub.rename(columns=dataname, inplace=True)
    df_trafos_sub.replace(' ', np.nan, inplace=True)
    
    # Leitura CSV dos Bancos de Capacitores
    df_bancos = pd.read_csv(Pasta_alimentador +"\\77322.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_bancos.rename(columns=dataname, inplace=True)
    df_bancos.replace(' ', np.nan, inplace=True)
    
    # Leitura CSV dos Religadores
    df_religadores = pd.read_csv(Pasta_alimentador +"\\77623.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_religadores.rename(columns=dataname, inplace=True)
    df_religadores.replace(' ', np.nan, inplace=True)
    
    # Leitura CSV dos fusiveis
    df_fusiveis = pd.read_csv(Pasta_alimentador +"\\77622.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_fusiveis.rename(columns=dataname, inplace=True)
    df_fusiveis.replace(' ', np.nan, inplace=True)
    
    # Leitura CSV da Subestação
    df_subestacao = pd.read_csv(Pasta_alimentador +"\\77823.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_subestacao.rename(columns=dataname, inplace=True)
    df_subestacao.replace(' ', np.nan, inplace=True)
    
    # Leitura CSV dos Reguladores
    df_reguladores = pd.read_csv(Pasta_alimentador +"\\78322.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_reguladores.rename(columns=dataname, inplace=True)
    df_reguladores.replace(' ', np.nan, inplace=True)
    
    # Extração das coordenadas das linhas
    df_lines = pd.concat([df_lines_aer, df_lines_sub], axis=0).reset_index(drop=True)
    df_lines.replace(' ', np.nan, inplace=True)
    df_lines.dropna(subset={'graf'}, inplace=True)
    
    # Leitura CSV das Chaves Seccionadoras
    df_chaves_aer = pd.read_csv(Pasta_alimentador +"\\78522.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_chaves_aer.rename(columns=dataname, inplace=True)
    df_chaves_aer.replace(' ', np.nan, inplace=True)
    
    # Leitura CSV das Chaves Seccionadoras
    df_chaves_sub = pd.read_csv(Pasta_alimentador +"\\78622.csv", skiprows=[0,1,2,3], sep=';', index_col=False, encoding='latin_1')
    df_chaves_sub.rename(columns=dataname, inplace=True)
    df_chaves_sub.replace(' ', np.nan, inplace=True)
    
    df_chaves = pd.concat([df_chaves_aer, df_chaves_sub], axis=0).reset_index(drop=True)
    
    
    def parse_linecoord(linecoord):
        if 'LINESTRING' not in linecoord: 
            return None, None    
        
        for i in range(len(df_lines)):
            variavel = linecoord.replace('LINESTRING (', '').replace(')', '').split(',')
            start_x, start_y = [float(i) for i in (variavel[0].split(' '))]
            end_x, end_y = [float(i) for i in (variavel[-1][1:].split(' '))]
        return [(start_x, start_y), (end_x, end_y)]
    
    coord_lines = df_lines['graf'].apply(parse_linecoord)
    
    ####################################################
    ####################################################
    i=0
    coord_lines.dropna(inplace=True)
    coord_lines = coord_lines.reset_index(drop=True)
    for i in range(len(coord_lines)):
        if len(coord_lines[i])==3:
            del coord_lines[i][1]
            
    ####################################################
    ####################################################
    
    df_lines = pd.concat([df_lines, pd.DataFrame(coord_lines.to_list(),
                                                  columns=['coord_start', 'coord_end'],
                                                  index=df_lines.index)], axis=1)
    
    df_lines = df_lines.drop(['graf'], axis=1)
    # df_lines.dropna(axis=0, inplace=True)
    
    # Criação das coordenadas das barras
    coord_buses = pd.concat([df_lines['coord_start'], df_lines['coord_end']], axis=0)
    df_buses = pd.DataFrame(coord_buses.unique(), columns=['coord'])
    df_buses['bus'] = df_buses.index
    dict_buses = df_buses.set_index('coord').bus.to_dict()
    
    # Extração das coordenadas dos trafos
    df_trafos = pd.concat([df_trafos_aer, df_trafos_sub], axis=0).reset_index(drop=True)
    df_trafos['coord_x'] = pd.to_numeric(df_trafos['coord_x'], errors='coerce')
    df_trafos['coord_y'] = pd.to_numeric(df_trafos['coord_y'], errors='coerce')
    df_trafos['coord'] = list(zip(df_trafos.coord_x, df_trafos.coord_y))
    
    # Extração das coordenadas dos bancos
    df_bancos['coord_x'] = pd.to_numeric(df_bancos['coord_x'], errors='coerce')
    df_bancos['coord_y'] = pd.to_numeric(df_bancos['coord_y'], errors='coerce')
    df_bancos['coord'] = list(zip(df_bancos.coord_x, df_bancos.coord_y))
    
    # Extração das coordenadas dos religadores
    df_religadores['coord_x'] = pd.to_numeric(df_religadores['coord_x'], errors='coerce')
    df_religadores['coord_y'] = pd.to_numeric(df_religadores['coord_y'], errors='coerce')
    df_religadores['coord'] = list(zip(df_religadores.coord_x, df_religadores.coord_y))
    df_religadores.dropna(subset={'fu'}, inplace=True)
    
    # Extração das coordenadas dos reguladores
    df_reguladores['coord_x'] = pd.to_numeric(df_reguladores['coord_x'], errors='coerce')
    df_reguladores['coord_y'] = pd.to_numeric(df_reguladores['coord_y'], errors='coerce')
    df_reguladores['coord'] = list(zip(df_reguladores.coord_x, df_reguladores.coord_y))
    
    # Extração das coordenadas dos fusiveis
    df_fusiveis['coord_x'] = pd.to_numeric(df_fusiveis['coord_x'], errors='coerce')
    df_fusiveis['coord_y'] = pd.to_numeric(df_fusiveis['coord_y'], errors='coerce')
    df_fusiveis['coord'] = list(zip(df_fusiveis.coord_x, df_fusiveis.coord_y))
    
    # # Extração das coordenadas da subestação
    df_subestacao['coord_x'] = pd.to_numeric(df_subestacao['coord_x'], errors='coerce')
    df_subestacao['coord_y'] = pd.to_numeric(df_subestacao['coord_y'], errors='coerce')
    df_subestacao['coord'] = list(zip(df_subestacao.coord_x, df_subestacao.coord_y))
    
    # # Extração das coordenadas da subestação
    df_chaves['coord_x'] = pd.to_numeric(df_chaves['coord_x'], errors='coerce')
    df_chaves['coord_y'] = pd.to_numeric(df_chaves['coord_y'], errors='coerce')
    df_chaves['coord'] = list(zip(df_chaves.coord_x, df_chaves.coord_y))
    
    # Mapeamento das barras das linhas
    df_lines['bus_start'] = df_lines['coord_start'].map(dict_buses)
    df_lines['bus_end'] = df_lines['coord_end'].map(dict_buses)
    
    # Mapeamento das barras dos trafos
    df_trafos['bus'] = df_trafos['coord'].map(dict_buses)
    
    # Mapeamento das barras dos bancos
    df_bancos['bus'] = df_bancos['coord'].map(dict_buses)
    
    # Mapeamento das barras dos fusíveis
    df_fusiveis['bus'] = df_fusiveis['coord'].map(dict_buses)
    
    # Mapeamento das barras dos religadores
    df_religadores['bus'] = df_religadores['coord'].map(dict_buses)
    
    # Mapeamento das barras dos reguladores
    df_reguladores['bus'] = df_reguladores['coord'].map(dict_buses)
    
    # Mapeamento das barras da subestação
    df_subestacao['bus'] = df_subestacao['coord'].map(dict_buses)
    
    # Mapeamento das barras da subestação
    df_chaves['bus'] = df_chaves['coord'].map(dict_buses)
    
    # Criando Bus com Número
    # Criação da Rede
    subestacao = df_subestacao['bus']
    subestacao = subestacao[0]
    subestacao = int(subestacao)
    tensao_alimentador = df_subestacao.iloc[0]['tensão']
    tensao_alimentador = tensao_alimentador[:1]
    if tensao_alimentador == '1':
        vn = 13.8
        vn_sub = '13'
        tensao_alimentador = '15'
        bitola = '185'
        tensao_cabo = '15'
        imped_x_caracteristica = 0.264
        imped_y_caracteristica = 0.295
        
    if tensao_alimentador == '2':
        vn = 23
        vn_sub = '23'
        tensao_alimentador = '23'
        bitola = '150'
        tensao_cabo = '25'
        imped_x_caracteristica = 0.210
        imped_y_caracteristica = 0.285
        
    if tensao_alimentador == '3':
        vn = 34.5
        vn_sub = '34,5'
        tensao_alimentador = '34.5'
        bitola = '150'
        tensao_cabo = '35'
        imped_x_caracteristica = 0.210
        imped_y_caracteristica = 0.285
        
    net = pp.create_empty_network()
    
    # Lançamento das barras
    
    df_nomes = df_trafos.drop(['alim','coord_x','coord_y','coord','proprietario','tap'], axis=1)
    df_fusiveis_nome = df_fusiveis.drop(['alim','coord_x','coord_y','coord'], axis=1)
    df_religadores_nome = df_religadores.drop(['alim','coord_x','coord_y','coord'], axis=1)
    df_chaves = df_chaves.drop(['alim','coord_x','coord_y','coord'], axis=1)
    
    df_nomes = pd.concat([df_nomes, df_fusiveis_nome])
    df_nomes = pd.concat([df_nomes, df_religadores_nome])
    df_nomes = pd.concat([df_nomes, df_chaves])
    
    df_nomes.dropna(subset={'bus'}, inplace=True)
    df_nomes.drop_duplicates(subset='bus', keep='first', inplace=True)
    
    df_buses = pd.merge(df_nomes, df_buses, how='right', left_on=['bus'], right_on = ['bus'])
    df_buses.apply(lambda bus: pp.create_bus(net, name=bus['fu'],
                                              index=bus['bus'],
                                              vn_kv=vn,
                                              geodata=bus['coord']), axis=1)
    
    # Lançamento das Linhas
    df_lines.apply(lambda line: pp.create_line_from_parameters(net,
                                                                from_bus=line['bus_start'],
                                                                to_bus=line['bus_end'],
                                                                length_km=float(line['comprimento'])/1000,
                                                                r_ohm_per_km=float(line['r1']),
                                                                x_ohm_per_km=float(line['z1']),
                                                                c_nf_per_km=0,
                                                                max_i_ka=100,
                                                                geodata=(line['coord_start'], line['coord_end'])), axis=1)
    
    df_trafos = df_trafos.dropna(subset=['bus'])
    
    # Lançamento das Cargas
    for i in df_trafos.index:
        if float(df_trafos['pot_dem'][i]) > 0:
            pp.create_load_from_cosphi(net, name=df_trafos['fu'][i],
                                        bus=df_trafos['bus'][i],
                                        sn_mva=float(df_trafos['pot_dem'][i])/1000,
                                        cos_phi=float(0.92),
                                        mode='underexcited')   
    
    # Lançamento dos Geradores
    for i in df_trafos.index:
        if float(df_trafos['pot_ger'][i]) > 0:
            pp.create_sgen(net,
                            bus=df_trafos['bus'][i],
                            p_mw=float(df_trafos['pot_ger'][i])/(1000*1000),
                            sn_mva=float(df_trafos['pot_ger'][i])/(1000*1000),
                            k=1.2)
       
    # Lançamento dos Bancos de Capacitores
    df_bancos = df_bancos.dropna(subset=['bus'])
    for i in df_bancos.index:
        pp.create_shunt_as_capacitor(net, name=df_trafos['fu'][i],
                        bus=df_bancos['bus'][i],
                        q_mvar=float(df_bancos['pot_nom'][i])/1000,
                        loss_factor=0)
    
    # # Lançamento do Regulador de Tensão
    # pandapower.create_transformer_from_parameters(net, 
    #                                               hv_bus, lv_bus, 
    #                                               sn_mva=100, vn_hv_kv=23100, vn_lv_kv=23100, vkr_percent, vk_percent, pfe_kw, 
    #                                               i0_percent, shift_degree=0, tap_side=None, tap_neutral=nan, tap_max=nan, tap_min=nan, tap_step_percent=nan, 
    #                                               tap_step_degree=nan, tap_pos=nan, tap_phase_shifter=False, in_service=True, name=None, vector_group=None, index=None, 
    #                                               max_loading_percent=nan, parallel=1, df=1.0, vk0_percent=nan, vkr0_percent=nan, mag0_percent=nan, mag0_rx=nan, si0_hv_partial=nan, 
    #                                               pt_percent=None, oltc=None, tap_dependent_impedance=None, vk_percent_characteristic=None, vkr_percent_characteristic=None, xn_ohm=None, **kwargs)
    
    sub = alim[0]
    sub = sub[:-2]
    
    dados_curto = pd.read_excel(r"C:\Users\joaof\OneDrive\Área de Trabalho\TCC\programa atualizado\Dados Subestações\ICC.xlsx", skiprows=[0])
    dados_curto = dados_curto[dados_curto['NOME'].str.contains(sub)]
    dados_curto = dados_curto[dados_curto['NOME'].str.contains(vn_sub)]
    dados_curto.reset_index(inplace=True)
    
    # Lançamento da Slack-Bus
    pp.create_ext_grid(net, subestacao,
                        s_sc_max_mva=dados_curto['CC 3F (MVA)'][0],
                        s_sc_min_mva=dados_curto['CC FF (MVA)'][0],
                        rx_max=dados_curto['X/R'][0],
                        rx_min=dados_curto['X/R.2'][0])
    
    # pp.create_ext_grid(net, subestacao,
    #                    s_sc_max_mva=dados_curto['CC 3F (MVA)'][0], rx_max=dados_curto['X/R'][0])
    
    # pp.create_ext_grid(net, subestacao,
    #                     s_sc_max_mva=860, rx_max=4)
    
    # pp.create_ext_grid(net, subestacao,
    #                     s_sc_max_mva=20*220*1.73, rx_max=0.1)
    
    ponto_de_conexao = df_buses[df_buses['fu'] == equip_conectado]
    ponto_conexao = ponto_de_conexao.iloc[0]['bus']
    ponto_conexao = int(ponto_conexao)
    
    # Roda o Fluxo de Potência
    start_time = time.time()
    
    pp.runpp(net, algorithm='nr')
    valor_inicial = net.res_bus.iloc[ponto_conexao][0]
    
    # Lançamento do Gerador em Análise
    pp.create_sgen(net, ponto_conexao, 
                    p_mw = potencia_conexao, 
                    sn_mva = potencia_conexao, 
                    k=1.2)
    
    pp.runpp(net, algorithm='nr')
    
    
    valor_final = net.res_bus.iloc[ponto_conexao][0]
    delta_v = (valor_final - valor_inicial)*100
    
    print("--- %s seconds ---" % (time.time() - start_time))
    print('Variação de Tensão Inicial: '+str(delta_v))
    print('Tensão inicial: '+str(valor_inicial))
    print('Tensão final: '+str(valor_final))
    
    # Caminho mais curto entre a SE e a Conexão
    # Calcula a distância
    import pandapower.topology as top
    import networkx as nx
    
    mg = top.create_nxgraph(net)
    path = nx.shortest_path(mg, subestacao, ponto_conexao)
    barras = net.bus.loc[path]
    barras.reset_index(inplace=True)
    df_barras = pd.merge(barras, df_buses, how='inner', left_on=['index'], right_on = ['bus'])
    df_reguladores.dropna(subset={'fu'}, inplace=True)
    
    for i in range(len(df_reguladores)):
        df_barras['name'] = np.where((df_barras["bus"] == df_reguladores['bus'][i]), df_reguladores['fu'][i], df_barras['name'])
    df_reguladores = df_reguladores.drop(columns={'coord_x','coord_y','rotacao','coord','id'})
    reguladores_troca = pd.merge(df_reguladores, df_barras, how='inner', left_on=['fu'], right_on = ['name'])
    
    
    linhas = net.line.loc[top.elements_on_path(mg, path, "line")]
    df_linhas = pd.merge(linhas, df_lines, how='inner', left_on=['from_bus','to_bus'], right_on = ['bus_start','bus_end'])
    
    # Recondutoramento Monofásico para Trifásico
    numero_linhas = 0
    recondutoramento_mono = 0
    equip_final_recon_mono = ''
    prox_barra_recond_mono = ''
    if '3 - T' in df_linhas.values or '2 - S' in df_linhas.values or '1 - R' in df_linhas.values:
        # Caso seja a fase T
        if '3 - T' in df_linhas.values:
            numero_linhas = (df_linhas['nfases'] == '3 - T').sum()
            lista_recond_mono = df_linhas[df_linhas['nfases'] == '3 - T']
            recondutoramento_mono = df_linhas.loc[df_linhas['nfases'] == '3 - T', 'length_km'].sum()
        # Caso seja a fase S
        if '2 - S' in df_linhas.values:
            numero_linhas = (df_linhas['nfases'] == '2 - S').sum()
            lista_recond_mono = df_linhas[df_linhas['nfases'] == '2 - S']
            recondutoramento_mono = df_linhas.loc[df_linhas['nfases'] == '2 - S', 'length_km'].sum()
        # Caso seja a fase R
        if '1 - R' in df_linhas.values:
            numero_linhas = (df_linhas['nfases'] == '1 - R').sum()
            lista_recond_mono = df_linhas[df_linhas['nfases'] == '1 - R']
            recondutoramento_mono = df_linhas.loc[df_linhas['nfases'] == '1 - R', 'length_km'].sum()
        for i in range(len(lista_recond_mono)):
            net.line['r_ohm_per_km'] = np.where(((net.line["from_bus"] == lista_recond_mono['from_bus'].iloc[-i]) & (net.line["to_bus"] == lista_recond_mono['to_bus'].iloc[-i])), imped_x_caracteristica, net.line['r_ohm_per_km'])
            net.line['x_ohm_per_km'] = np.where(((net.line["from_bus"] == lista_recond_mono['from_bus'].iloc[-i]) & (net.line["to_bus"] == lista_recond_mono['to_bus'].iloc[-i])), imped_y_caracteristica, net.line['x_ohm_per_km'])
        prox_barra_recond_mono = df_barras[:-numero_linhas]
        prox_barra_recond_mono.dropna(subset={'name'}, inplace=True)
        prox_barra_recond_mono = prox_barra_recond_mono['name'].iloc[-1]
        if any(df_trafos['fu'] == prox_barra_recond_mono):
                # barra_regulador é trafo
                equip_final_recon_mono='Transformador'
        if any(df_fusiveis['fu'] == prox_barra_recond_mono):
                # barra_regulador é fusivel
                equip_final_recon_mono='Fusível'
        if any(df_religadores['fu'] == prox_barra_recond_mono):
                # barra_regulador é religador
                equip_final_recon_mono='Regulador'
        if any(df_chaves['fu'] == prox_barra_recond_mono):
                # barra_regulador é religador
                equip_final_recon_mono='Chave Seccionadora'
        
    else:
        recondutoramento_mono = 0
        
    caminho_recondutorado_mono = df_linhas.iloc[-i:]
    recondutorado_mono = linhas[-i:].index
    
    # Recondutoramento pela Variação de Tensão
    # df_linhas = df_linhas[:-numero_linhas]
    i = 1
    y = 0
    maximo = len(df_linhas)
    delta_recondutorado = 0
    if delta_v > 4:
        try:
            while delta_v > 4 or maximo < y:
                net.line['r_ohm_per_km'] = np.where(((net.line["from_bus"] == df_linhas['from_bus'].iloc[-i]) & (net.line["to_bus"] == df_linhas['to_bus'].iloc[-i])), imped_x_caracteristica, net.line['r_ohm_per_km'])
                net.line['x_ohm_per_km'] = np.where(((net.line["from_bus"] == df_linhas['from_bus'].iloc[-i]) & (net.line["to_bus"] == df_linhas['to_bus'].iloc[-i])), imped_y_caracteristica, net.line['x_ohm_per_km'])
            
                net.sgen['in_service'] = np.where(net.sgen["bus"] == ponto_conexao, False, net.sgen['in_service'])
                pp.runpp(net, algorithm='nr')
                valor_inicial = net.res_bus.iloc[ponto_conexao][0]
            
                net.sgen['in_service'] = np.where(net.sgen["bus"] == ponto_conexao, True, net.sgen['in_service'])
                pp.runpp(net, algorithm='nr')
            
                valor_final = net.res_bus.iloc[ponto_conexao][0]
                delta_v = (valor_final - valor_inicial)*100
                
                print(delta_v)
                i=i+1
                y=y+1
        except:
            print('SO para o planejamento')
            sys.exit()
    if y == 0:
        delta_recondutorado = 0
    if y != 0:
        delta_recondutorado = df_linhas['length_km'].iloc[-y:].sum()
    
    recondutorament_total = delta_recondutorado
    recondutoramento_normal = delta_recondutorado-recondutoramento_mono
    caminho_recondutorado = df_linhas.iloc[-y:]
    recondutorado = linhas[-y:].index
    
    # Recondutoramento Trifásico - Local e Equipamento
    equip_comeco_recond =''
    equip_final_recond = ''
    recondutorado_trifasico_inicial =''
    recondutorado_trifasico_final=''
    
    if (recondutoramento_normal > 0) & (recondutoramento_mono == 0):
        local_comeco_recond = str(alim[1])
        if any(df_trafos['fu'] == alim[1]):
                # barra_regulador é trafo
                equip_comeco_recond='Transformador'
        if any(df_fusiveis['fu'] == alim[1]):
                # barra_regulador é fusivel
                equip_comeco_recond='Fusível'
        if any(df_religadores['fu'] == alim[1]):
                # barra_regulador é religador
                equip_comeco_recond='Regulador'
        if any(df_chaves['fu'] == alim[1]):
                # barra_regulador é religador
                equip_final_recon_mono='Chave Seccionadora'
    
    if (recondutoramento_normal > 0) & (recondutoramento_mono > 0):
        recondutorado_trifasico_inicial = barras[y:]
        recondutorado_trifasico_inicial = pd.merge(recondutorado_trifasico_inicial, df_buses, how='inner', left_on=['index'], right_on = ['bus'])
        recondutorado_trifasico_inicial.dropna(subset={'name'}, inplace=True)
        recondutorado_trifasico_inicial = recondutorado_trifasico_inicial['name'].tail(1)
        recondutorado_trifasico_inicial = recondutorado_trifasico_inicial.reset_index()
        recondutorado_trifasico_inicial = recondutorado_trifasico_inicial['name'][0]
        if any(df_trafos['fu'] == recondutorado_trifasico_inicial):
                # barra_regulador é trafo
                equip_comeco_recond='Transformador'
        if any(df_fusiveis['fu'] == recondutorado_trifasico_inicial):
                # barra_regulador é fusivel
                equip_comeco_recond='Fusível'
        if any(df_religadores['fu'] == recondutorado_trifasico_inicial):
                # barra_regulador é religador
                equip_comeco_recond='Regulador'
        if any(df_chaves['fu'] == recondutorado_trifasico_inicial):
                # barra_regulador é religador
                equip_final_recon_mono='Chave Seccionadora'            
    
        
        recondutorado_trifasico_final = barras[:-y]
        recondutorado_trifasico_final = pd.merge(recondutorado_trifasico_final, df_buses, how='inner', left_on=['index'], right_on = ['bus'])
        recondutorado_trifasico_final.dropna(subset={'name'}, inplace=True)
        recondutorado_trifasico_final = recondutorado_trifasico_final['name'].tail(1)
        recondutorado_trifasico_final = recondutorado_trifasico_final.reset_index()
        recondutorado_trifasico_final = recondutorado_trifasico_final['name'][0]
        if any(df_trafos['fu'] == recondutorado_trifasico_final):
                # barra_regulador é trafo
                equip_final_recond='Transformador'
        if any(df_fusiveis['fu'] == recondutorado_trifasico_final):
                # barra_regulador é fusivel
                equip_final_recond='Fusível'
        if any(df_religadores['fu'] == recondutorado_trifasico_final):
                # barra_regulador é religador
                equip_final_recond='Regulador'
        if any(df_chaves['fu'] == recondutorado_trifasico_final):
                # barra_regulador é religador
                equip_final_recon_mono='Chave Seccionadora'
                
    def subs_equipamentos(df_barras):
        df_barras['coord_prox'] = df_barras['coord'].shift(1)
        df_barras['coord_ant'] = df_barras['coord'].shift(-1)
        df_barras.dropna(subset={'rotacao'}, inplace=True)
        df_barras.dropna(subset={'coord_ant'}, inplace=True)
        
        df_barras['coord'] = df_barras['coord'].astype(str)
        df_barras['coord'] = df_barras['coord'].str.replace('(','', regex=True)
        df_barras['coord'] = df_barras['coord'].str.replace(')','', regex=True)
        
        df_barras['coord_prox'] = df_barras['coord_prox'].astype(str)
        df_barras['coord_prox'] = df_barras['coord_prox'].str.replace('(','', regex=True)
        df_barras['coord_prox'] = df_barras['coord_prox'].str.replace(')','', regex=True)
        
        df_barras['coord_ant'] = df_barras['coord_ant'].astype(str)
        df_barras['coord_ant'] = df_barras['coord_ant'].str.replace('(','', regex=True)
        df_barras['coord_ant'] = df_barras['coord_ant'].str.replace(')','', regex=True)
        
        # Create two lists for the loop results to be placed
        lat_init = []
        lon_init = []
        for row in df_barras['coord']:
            try:
                lat_init.append(row.split(',')[0])
                lon_init.append(row.split(',')[1])
            except:
                lat_init.append(np.NaN)
                lon_init.append(np.NaN)
        
        df_barras['coord_x_inicio'] = lat_init
        df_barras['coord_y_inicio'] = lon_init
        
        lat_final = []
        lon_final = []
        
        for row in df_barras['coord_prox']:
            try:
                lat_final.append(row.split(',')[0])
                lon_final.append(row.split(',')[1])
            except:
                lat_final.append(np.NaN)
                lon_final.append(np.NaN)
        
        df_barras['coord_x_final'] = lat_final
        df_barras['coord_y_final'] = lon_final
        
        lat_final = []
        lon_final = []
        
        for row in df_barras['coord_ant']:
            try:
                lat_final.append(row.split(',')[0])
                lon_final.append(row.split(',')[1])
            except:
                lat_final.append(np.NaN)
                lon_final.append(np.NaN)
        
        df_barras['coord_x_anterior'] = lat_final
        df_barras['coord_y_anterior'] = lon_final
        
        def UTM2LatLong(df_barras, coluna_x, coluna_y, csv: bool):
            df_barras[[coluna_x, coluna_y]] = df_barras[[coluna_x, coluna_y]].applymap(lambda x: float(str(x).replace('(', '')))
            df_barras[[coluna_x, coluna_y]] = df_barras[[coluna_x, coluna_y]].applymap(lambda x: float(str(x).replace(')', '')))
        
            df_barras[[coluna_x, coluna_y]] = df_barras[[coluna_x, coluna_y]].applymap(lambda x: float(str(x).replace(',', '.')))
            df_barras[coluna_y] = df_barras[coluna_y] - 10000000
            
            #Cálculos
            df_barras['Latitude_init'] = 0
            df_barras['Longitude_init'] = 0
            myProj = Proj(proj='utm',zone=22,ellps='WGS84', preserve_units=False)
            df_barras['Longitude_init'], df_barras['Latitude_init'] = myProj(df_barras[coluna_x].values, df_barras[coluna_y].values, inverse = True)        
        
        coluna_x = 'coord_x_inicio'
        coluna_y = 'coord_y_inicio'
        UTM2LatLong(df_barras, coluna_x, coluna_y, csv = True)
        
        def UTM2LatLong(df_barras, coluna_x, coluna_y, csv: bool):
            df_barras[[coluna_x, coluna_y]] = df_barras[[coluna_x, coluna_y]].applymap(lambda x: float(str(x).replace(',', '.')))
            df_barras[coluna_y] = df_barras[coluna_y] - 10000000
            
            #Cálculos
            df_barras['Latitude_final'] = 0
            df_barras['Longitude_final'] = 0
            myProj = Proj(proj='utm',zone=22,ellps='WGS84', preserve_units=False)
            df_barras['Longitude_final'], df_barras['Latitude_final'] = myProj(df_barras[coluna_x].values, df_barras[coluna_y].values, inverse = True)        
            
        coluna_x = 'coord_x_final'
        coluna_y = 'coord_y_final'
        UTM2LatLong(df_barras, coluna_x, coluna_y, csv = True)
        
        def UTM2LatLong(df_barras, coluna_x, coluna_y, csv: bool):
            df_barras[[coluna_x, coluna_y]] = df_barras[[coluna_x, coluna_y]].applymap(lambda x: float(str(x).replace('(', '')))
            df_barras[[coluna_x, coluna_y]] = df_barras[[coluna_x, coluna_y]].applymap(lambda x: float(str(x).replace(')', '')))
        
            df_barras[[coluna_x, coluna_y]] = df_barras[[coluna_x, coluna_y]].applymap(lambda x: float(str(x).replace(',', '.')))
            df_barras[coluna_y] = df_barras[coluna_y] - 10000000
            
            #Cálculos
            df_barras['Latitude_anterior'] = 0
            df_barras['Longitude_anterior'] = 0
            myProj = Proj(proj='utm',zone=22,ellps='WGS84', preserve_units=False)
            df_barras['Longitude_anterior'], df_barras['Latitude_anterior'] = myProj(df_barras[coluna_x].values, df_barras[coluna_y].values, inverse = True)        
        
        coluna_x = 'coord_x_anterior'
        coluna_y = 'coord_y_anterior'
        UTM2LatLong(df_barras, coluna_x, coluna_y, csv = True)
        df_barras.drop(labels={'coord','coord_prox','coord_x_anterior','coord_y_anterior','coord_x_inicio','coord_y_inicio','coord_x_final','coord_y_final'}, inplace=True, axis=1)
        
        
        for row in df_barras:
            long1= df_barras['Longitude_init']
            long2=df_barras['Longitude_final']
            lat1=df_barras['Latitude_init']
            lat2=df_barras['Latitude_final']
            geodesic = pyproj.Geod(ellps='clrk66')
            df_barras['az12_frente'], df_barras['az21_frente'], df_barras['distance'] = geodesic.inv(long1, lat1, long2, lat2)
            
            long1= df_barras['Longitude_init']
            long2=df_barras['Longitude_anterior']
            lat1=df_barras['Latitude_init']
            lat2=df_barras['Latitude_anterior']
            geodesic = pyproj.Geod(ellps='clrk66')
            df_barras['az12_atras'], df_barras['az21_atras'], df_barras['distance'] = geodesic.inv(long1, lat1, long2, lat2)
            
        df_barras.drop(labels={'distance','Latitude_init','Longitude_init','Longitude_final','Latitude_final','Longitude_anterior','Longitude_anterior','Latitude_init','Latitude_anterior'}, inplace=True, axis=1)
        rotacao = []
        i=0
        df_barras.reset_index(inplace=True)
        df_barras['rotacao'] = df_barras['rotacao'].astype(float)
        for i in range(len(df_barras)):
            if df_barras['rotacao'][i] >= 0 and df_barras['rotacao'][i] < 90:
                a = (90 - df_barras['az12_frente'][i])
                rotacao.append(a)
            
            if df_barras['rotacao'][i] >= 90 and df_barras['rotacao'][i] < 180:
                a = (90 - df_barras['az12_frente'][i])    
                rotacao.append(a)
        
            if df_barras['rotacao'][i] >= 180 and df_barras['rotacao'][i] < 270:
                a = (90 - df_barras['az12_frente'][i])    
                rotacao.append(a)
        
            if df_barras['rotacao'][i] >= 270 and df_barras['rotacao'][i] < 360:
                a = (-(df_barras['az12_frente'][i]-450))     
                rotacao.append(a)
        df_barras['Rot_calculado_1']=rotacao
        rotacao2 = []
        i=0
        for i in range(len(df_barras)):
            if df_barras['rotacao'][i] >= 0 and df_barras['rotacao'][i] < 90:
                a = (90 - df_barras['az12_atras'][i])
                rotacao2.append(a)
            
            if df_barras['rotacao'][i] >= 90 and df_barras['rotacao'][i] < 180:
                a = (90 - df_barras['az12_atras'][i])    
                rotacao2.append(a)
        
            if df_barras['rotacao'][i] >= 180 and df_barras['rotacao'][i] < 270:
                a = (90 - df_barras['az12_atras'][i])    
                rotacao2.append(a)
        
            if df_barras['rotacao'][i] >= 270 and df_barras['rotacao'][i] < 360:
                a = (-(df_barras['az12_atras'][i]-450))     
                rotacao2.append(a)
        df_barras['Rot_calculado_2']=rotacao2
        equipamentos_no_caminho = df_barras.loc[((df_barras['rotacao'] >= df_barras['Rot_calculado_1']-2) & (df_barras['rotacao'] <= df_barras['Rot_calculado_1']+2)) | ((df_barras['rotacao'] >= df_barras['Rot_calculado_2']-2) & (df_barras['rotacao'] <= df_barras['Rot_calculado_2']+2))]
        equipamentos=equipamentos_no_caminho['name']
        if len(equipamentos) != 0:
            fusiveis_troca = pd.merge(df_fusiveis, equipamentos, how='inner', left_on=['fu'], right_on = ['name'])
            religadores_troca = pd.merge(df_religadores, equipamentos, how='inner', left_on=['fu'], right_on = ['name'])
        else:
            fusiveis_troca = []
            religadores_troca = []
        
        return df_barras, equipamentos_no_caminho, fusiveis_troca, religadores_troca;
    
    df_barras, equipamentos_no_caminho, fusiveis_troca, religadores_troca = subs_equipamentos(df_barras)
    # df_reguladores.dropna(subset=['bus'], inplace=True)
    
    def curto_circuito(potencia_conexao, vn):
        for line in net.sgen:
            net.sgen['in_service'] = False
        
        import pandapower.shortcircuit as sc
        sc.calc_sc(net, case="max", ip=True, ith=True, branch_results=True)
        curto = net.res_bus_sc
        curto_gd = curto.iloc[ponto_conexao]
        
        pot_usina = float(potencia_conexao)*1000
        # vn = 13.8
        
        Pot_curto = curto_gd['ikss_ka']*1000
        delta_curto = pot_usina/(Pot_curto*vn)
        print('Corrente de Curto-Circuito: ')
        print(delta_curto*100)
    curto_circuito(potencia_conexao, vn)
    
    def nivel_tensao():
        import plotly.express as px
        import plotly.io as pio
        pio.renderers.default='browser'
        # pio.renderers.default='svg'
        resultado_tensao = net.res_bus
        resultado_tensao = resultado_tensao.reset_index()
        df_tensao_bar = pd.merge(barras, resultado_tensao, how='inner', left_on=['index'], right_on = ['index'])
        df_tensao_bar['color'] = 0
        for i in range(len(df_tensao_bar)):
            if df_tensao_bar['vm_pu'][i] < 0.95:
                df_tensao_bar['color'][i] = 'red'
            if df_tensao_bar['vm_pu'][i] >= 0.95:
                df_tensao_bar['color'][i] = 'blue'
    
        fig = px.bar(df_tensao_bar, x=df_tensao_bar.index, y='vm_pu',
                 hover_data=['vm_pu'], color=df_tensao_bar['color'],
                 labels={'index':'Barras'}, height=700)
        fig.show()
    nivel_tensao()    
    
    def plot_tensao():
        from pandapower.plotting.plotly import pf_res_plotly
        pf_res_plotly(net)
        
    plot_tensao()
    
    def plot_recondutoramento(ponto_de_conexao):
        import pandapower.plotting.plotly as pplotly
        
        gd = pplotly.create_bus_trace(net, ponto_de_conexao, size=15, color="#6db9a2", patch_type='square', trace_name='Unidade Geradora')
        lc = pplotly.create_line_trace(net,net.line.index, color='black')
        lcl = pplotly.create_line_trace(net, recondutorado, color="red", width=2,
                                        infofunc=Series(index=net.line.index, 
                                                        data=net.line.name[recondutorado] + '<br>' + net.line.length_km[recondutorado].astype(str) + ' km'))
        ext_grid_trace = pplotly.create_bus_trace(net, buses=net.ext_grid.bus,
                                          color='#fefe8a', size=10, trace_name='Subestação',
                                          patch_type='square')
        pplotly.draw_traces(lc + lcl + gd + ext_grid_trace)
    plot_recondutoramento(ponto_de_conexao)
    
    def plotagem(fusiveis_troca, religadores_troca):
        from pandapower.plotting.plotly.mapbox_plot import set_mapbox_token
        import pandapower.plotting.plotly as pplotly
            
        df_equip_religador = pd.merge(barras, df_religadores, how='inner', left_on=['index'], right_on = ['bus'])
        religadores_troca.drop(['alim','id','fu','coord_x','coord_y','rotacao','coord','bus'], axis=1,inplace=True)
        df_equip_religador = pd.merge(religadores_troca, df_equip_religador, how='inner', left_on=['name'], right_on = ['name'])
        df_equip_religador=df_equip_religador.set_index('index')
        religadores = df_equip_religador.index
        
        # Encontra os fusíveis no caminho
        df_equip_fusivel = pd.merge(barras, df_fusiveis, how='inner', left_on=['index'], right_on = ['bus'])
        fusiveis_troca.drop(['alim','id','fu','coord_x','coord_y','rotacao','coord','bus'], axis=1,inplace=True)
        df_equip_fusivel = pd.merge(fusiveis_troca, df_equip_fusivel, how='inner', left_on=['name'], right_on = ['name'])
        df_equip_fusivel=df_equip_fusivel.set_index('index')
        fusiveis = df_equip_fusivel.index 
        
        ponto_de_conexao = df_buses[df_buses['fu'] == equip_conectado]
        ponto_de_conexao=ponto_de_conexao.index
        
        ponto_subestacao = df_buses[df_buses['bus'] == subestacao]
        ponto_subestacao = ponto_subestacao.index
        
        caminho_eletrico = linhas.index
        
        rec = pplotly.create_line_trace(net, recondutorado_mono, trace_name='Recondutorado', color="#8B0000", width=3,
                                            infofunc=Series(index=net.line.index, 
                                                            data=net.line.name[recondutorado_mono] + '<br>' + net.line.length_km[recondutorado_mono].astype(str) + ' km'))
        gd = pplotly.create_bus_trace(net, ponto_de_conexao, size=15, color="#6db9a2", patch_type='square', trace_name='Unidade Geradora')
        rl = pplotly.create_bus_trace(net, religadores, size=10, color="#a8476f", patch_type='square', trace_name='Religador')
        fu = pplotly.create_bus_trace(net, fusiveis, size=10, color="#21512b", trace_name='Fusível')
    
        ext_grid_trace = pplotly.create_bus_trace(net, buses=net.ext_grid.bus,
                                          color='#fefe8a', size=10, trace_name='Subestação',
                                          patch_type='square')
        bc = pplotly.create_bus_trace(net, net.bus.index, size=1, color="gray",
                                infofunc=Series(index=net.bus.index,
                                                data=net.bus.name + '<br>' + net.bus.vn_kv.astype(str) + ' kV'))
        
        lc = pplotly.create_line_trace(net, net.line.index, trace_name='Ramais', color='gray')
        lcl = pplotly.create_line_trace(net, caminho_eletrico, trace_name='Caminho Elétrico', color="black", width=3)
        
        pplotly.draw_traces(lc + lcl + rec + rl + fu + gd + ext_grid_trace)
        pplotly.draw_traces(lc + gd + ext_grid_trace)
    
    try:
        plotagem(fusiveis_troca, religadores_troca)
    except:
        print('hi')
    
    #########################################
    #########################################
    df_lista_rel_subs = pd.read_excel("C:\\Users\\joaof\\OneDrive\\Área de Trabalho\\TCC\programa atualizado\\Lista Religadores\\Religadores Subestação.xlsx")
    df_lista_rel_alim = pd.read_excel("C:\\Users\\joaof\\OneDrive\\Área de Trabalho\\TCC\programa atualizado\\Lista Religadores\\Religadores Alimentador.xlsx")
    df_alim_por_reg = pd.read_excel(r"C:\Users\joaof\OneDrive\Área de Trabalho\TCC\programa atualizado\Lista Religadores\Lista Regionais.xlsx")
    
    df_lista_rel_alim = pd.merge(df_lista_rel_alim, df_alim_por_reg, how='left', left_on=['REG'], right_on = ['REG'])
    
    resultado_tensao = net.res_bus
    resultado_tensao = resultado_tensao.reset_index()
    df_tensao_bar = pd.merge(barras, resultado_tensao, how='inner', left_on=['index'], right_on = ['index'])
    df_tensao_bar.dropna(subset={'name'}, inplace=True)
    
    # Instalação de Regulador de Tensão e local da sua instalação
    barra_regulador = []
    num_equip_local_reg = 0
    equip_local_reg = ''
    if (len(df_reguladores) == 0) & len(df_tensao_bar.loc[(df_tensao_bar['vm_pu'] <= 0.95)]) != 0:
        bar_regulador = df_tensao_bar.loc[(df_tensao_bar['vm_pu'] <= 0.95)]
        bar_regulador = bar_regulador.head(1)
        bar_regulador.reset_index(inplace=True)
        barra_regulador=[]
        barra_regulador.append(bar_regulador['name'][0])
        if any(df_trafos['fu'] == bar_regulador['name'][0]):
                # barra_regulador é trafo
                num_equip_local_reg = bar_regulador['name'][0]
                equip_local_reg='Transformador'
        if any(df_fusiveis['fu'] == bar_regulador['name'][0]):
                # barra_regulador é fusivel
                num_equip_local_reg = bar_regulador['name'][0]
                equip_local_reg='Fusível'
        if any(df_religadores['fu'] == bar_regulador['name'][0]):
                # barra_regulador é religador
                num_equip_local_reg = bar_regulador['name'][0]
                equip_local_reg='Regulador'
        if any(df_chaves['fu'] == bar_regulador['name'][0]):
                # barra_regulador é Chave Secc
                equip_final_recon_mono='Chave Seccionadora'
                
    # Troca do fusível igual ao do Trafo por religador
    try:
        fusiveis_para_troca_faca = fusiveis_troca['name'].tolist()
    except:
        fusiveis_para_troca_faca = []
        
    fusiveis_para_troca_reli = []
    if len(df_fusiveis.loc[(df_fusiveis['fu'] == alim[1])]) == len(df_trafos.loc[(df_trafos['fu'] == alim[1])]):
        #troca o fusível por religador
        fus_troca = df_fusiveis.loc[(df_fusiveis['fu'] == str(alim[1]))]
        fusiveis_para_troca_reli = fus_troca['fu'].tolist()
        
    # Troca do Religador da Subestação
    religador_substacao = df_lista_rel_subs.loc[(df_lista_rel_subs['ID'] == alim[0])]['Sensor de Tensão no 2 lado do Religador ?']
    religador_substacao = religador_substacao.reset_index()
    if religador_substacao['Sensor de Tensão no 2 lado do Religador ?'][0] == 'NÃO':
        relig_subs = 'Sim'
    else:
        relig_subs = 'Não'
        
    # Troca do Religador do Alimentador
    lista_religadores_troca = []
    fora_da_lista_DVAS = []
    i=0
    df_lista_rel_alim.set_index('N_EQP', inplace=True)
    for i in range(len(religadores_troca)):
        try:
            if df_lista_rel_alim.loc[(df_lista_rel_alim['Alimentador'] == (alim[0]))]['Possuem Sensor nos 2 lados do RL ?'][int(religadores_troca['name'][i])] == 'NÃO':
                lista_religadores_troca.append(religadores_troca['name'][i])
        except:
            fora_da_lista_DVAS.append(religadores_troca['name'][i])
            
    # Lista dos Religadores que necessitam de troca:
    lista_religadores_troca.extend(fora_da_lista_DVAS)
    
    #####################################
    # Construção de rede
    construcao_rede = 'Sim'
    
    # Recondutoramento Mono
    if recondutoramento_mono <= 0:
        rec_mono_tri = 'Não'
    else:
        rec_mono_tri = 'Sim'
        
    # Recondutoramento Tri
    if recondutoramento_normal <= 0:
        rec_tri = 'Não'
    else:
        rec_tri = 'Sim'
    
    # Troca de Religadores de Subs
    relig_subs  
            
    # Troca de Religadores de alimentador
    if lista_religadores_troca:
        tem_relig = 'Sim'
    else:
        tem_relig = 'Não'  
    
    # Troca de Regulador no alimentador
    reguladores_troca = reguladores_troca['fu_x'].tolist()
    if reguladores_troca:
        tem_troca_reg = 'Sim'
    else:
        tem_troca_reg = 'Não'
    
    # Religador na Entrada
    if potencia_conexao >= 0.3:
        relig_entrada = 'Sim'
    else:
        relig_entrada = 'Não'        
    
    # Instalação de Regulador
    if barra_regulador:
        tem_regu = 'Sim'
    else:
        tem_regu = 'Não'  
    
    # Fusível por Reli
    if fusiveis_para_troca_reli:
        tem_fus_por_reli = 'Sim'
    else:
        tem_fus_por_reli = 'Não'  
            
    # Fusível por Faca
    if fusiveis_para_troca_faca:
        tem_fus_por_faca = 'Sim'
    else:
        tem_fus_por_faca = 'Não'  
    
    #### Estrutura
    # caminhoSO = r'K:\DPGT_DVGT\1. GERAÇÃO DISTRIBUÍDA\01. Acesso ao Sistema Elétrico\01. Processo de análise Minigeração\01. Informação de Acesso\Automatização\João\programa fluxo de potência\teste SO'
    numero_so = SO_automatica
    caminhoSO = Pasta_SO(numero_so, df_alim_por_reg, alim)
    
    # Construção
    if any(df_trafos['fu'] == str(alim[1])):
            # barra_regulador é trafo
            const_rede='Transformador'
    if any(df_fusiveis['fu'] == str(alim[1])):
            # barra_regulador é fusivel
            const_rede='Fusível'
    if any(df_religadores['fu'] == str(alim[1])):
            # barra_regulador é religador
            const_rede='Regulador'
    if any(df_chaves['fu'] == str(alim[1])):
            # barra_regulador é Chave Secc
            equip_final_recon_mono='Chave Seccionadora'
    contrucao_rede = str(round(alim[2]*1.1,0))
    equi_contrucao_rede = const_rede
    num_equip_construcao_rede = str(alim[1])
    
    # Recondutoramento Mono
    recond_mono = str(round(recondutoramento_mono*1.1,2))
    if any(df_trafos['fu'] == str(alim[1])):
            # barra_regulador é trafo
            equip_final_recond_mono='Transformador'
    if any(df_fusiveis['fu'] == str(alim[1])):
            # barra_regulador é fusivel
            equip_final_recond_mono='Fusível'
    if any(df_religadores['fu'] == str(alim[1])):
            # barra_regulador é religador
            equip_final_recond_mono='Regulador'
    if any(df_chaves['fu'] == str(alim[1])):
            # barra_regulador é Chave Secc
            equip_final_recon_mono='Chave Seccionadora'
    tipo_ponto_a_mono = equip_final_recond_mono
    ponto_a_mono = str(alim[1])
    tipo_ponto_b_mono = equip_final_recon_mono
    ponto_b_mono = prox_barra_recond_mono
    
    # Recondutoramento Tri
    recond = str(round(recondutoramento_normal*1.1, 1))
    tipo_ponto_a = equip_comeco_recond
    ponto_a = recondutorado_trifasico_inicial
    tipo_ponto_b = equip_final_recond
    ponto_b = recondutorado_trifasico_final
    
    # Religadores
    religador_se = ''
    subestacao = sigla
    n_reli_troca = len(lista_religadores_troca)
    religadores = ', '.join(lista_religadores_troca)
    
    # Regulador de Tensão
    kv_banco = str(vn)
    corrente_banco = '200 A'
    tipo_ponto_banco = equip_local_reg
    num_ponto_banco = num_equip_local_reg
    n_regu_troca = len(reguladores_troca)
    regulador = ', '.join(reguladores_troca)
    
    # Fusíveis
    fusiveis_religadores = ', '.join(fusiveis_para_troca_reli)
    n_fus_reli = len(fusiveis_para_troca_reli)
    fusiveis_faca = ', '.join(fusiveis_para_troca_faca)
    n_fus_faca = len(fusiveis_para_troca_faca)
    alimentador = alimentador
    
    try:
        word(potencia_conexao, num_so, n_reli_troca, n_regu_troca, n_fus_faca,n_fus_reli, SO_automatica, tem_fus_por_faca, tem_fus_por_reli, tem_regu, relig_entrada, tem_troca_reg, tem_relig, relig_subs, rec_tri, rec_mono_tri, construcao_rede, caminhoSO, numero_so, nome_so, contrucao_rede, bitola, tensao_cabo, equi_contrucao_rede, num_equip_construcao_rede, recond_mono, tipo_ponto_a_mono, ponto_a_mono,tipo_ponto_b_mono,ponto_b_mono,recond,tipo_ponto_a,ponto_a,tipo_ponto_b,ponto_b,religador_se,subestacao,religadores,kv_banco,corrente_banco,tipo_ponto_banco,num_ponto_banco,regulador,fusiveis_religadores,fusiveis_faca,alimentador, tensao_alimentador, UTMX, UTMY)
    
    except:
        pasta_solicitacao = r'C:\Users\joaof\OneDrive\Área de Trabalho\TCC\programa automático\Solicitações Automáticas'
        pasta_sol = pasta_solicitacao + "\\" + sigla + "\\" + alimentador
        
        try:
            os.makedirs(pasta_sol)
        except:
            print('Essa Pasta Já Existe.')
        else:
            print('A Pasta Foi Criada.')
            
        word2(potencia_conexao, pasta_sol, num_so, n_reli_troca, n_regu_troca, n_fus_faca,n_fus_reli, SO_automatica, tem_fus_por_faca, tem_fus_por_reli, tem_regu, relig_entrada, tem_troca_reg, tem_relig, relig_subs, rec_tri, rec_mono_tri, construcao_rede, caminhoSO, numero_so, nome_so, contrucao_rede, bitola, tensao_cabo, equi_contrucao_rede, num_equip_construcao_rede, recond_mono, tipo_ponto_a_mono, ponto_a_mono,tipo_ponto_b_mono,ponto_b_mono,recond,tipo_ponto_a,ponto_a,tipo_ponto_b,ponto_b,religador_se,subestacao,religadores,kv_banco,corrente_banco,tipo_ponto_banco,num_ponto_banco,regulador,fusiveis_religadores,fusiveis_faca,alimentador, tensao_alimentador)
    
