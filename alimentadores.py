import pandas as pd
import numpy as np
from os.path import isfile, join
from os import listdir
from copy import deepcopy


# UTMX = str(642111.97)
# UTMY = str(7002505.73)

def ponto_analise(UTMX, UTMY):
    ###########################################
    #  UnificarDFs: soma todos os DFs         #
    #  de um dicionário, ignorando os valores #
    #  do  filtro                             #
    ###########################################
    def UnificarDFs(dicionario, nova_coluna :str, filtro = []):
        chaves = list(dicionario.keys())
        for chave in chaves:
            dicionario[chave][nova_coluna] = chave
            if chave in filtro:
                del dicionario[chave]
        df = pd.concat(dicionario, sort = False)
        df.reset_index(drop = True, inplace = True)
        return df
    
    ###########################################
    #   RenomearColuna: Renomeia as colunas   #
    #   de um DataFrame para os valores       #
    #   com base em seu nome inicial          #
    ###########################################
    def RenomearColuna(df):
        dici = {
                'Conectado?': 'Em Serviço',
                'Fase Ligada': 'Fase'}
        
        colunas = df.columns.to_list()
        for i, coluna in enumerate(colunas):
            if coluna in list(dici.keys()):
                colunas[i] = dici[coluna]
        df.columns = colunas
        return df
    ###########################################
    #   FiltroDist: Calcula a distância       #
    #   Entre o ponto fornecido e os equips   #
    #   da Celesc                             #
    ###########################################
    def FiltroDist(df,utm_x, utm_y, passos: int, filtro_inicial: int, pontos: int, curvas = 1.4):
        #Cálculo das distâncias máximas e mínimas
        x_max = filtro_inicial + utm_x
        x_min = - filtro_inicial + utm_x
        y_max = filtro_inicial + utm_y
        y_min = - filtro_inicial + utm_y
    
        #Aquisição das colunas com as coordenadas
        colunas = df.columns
        col_coord = [s for s in colunas if s.__contains__('Coordenada')]
        
        #Backup do DF
        aux = deepcopy(df)
        
        #Eixo X
        indexNames = aux[aux[col_coord[0]] > x_max].index
        aux.drop(indexNames, inplace = True)
        indexNames = aux[aux[col_coord[0]] < x_min].index
        aux.drop(indexNames, inplace = True)
        
        #Eixo Y
        indexNames = aux[aux[col_coord[1]] > y_max].index
        aux.drop(indexNames, inplace = True)
        indexNames = aux[aux[col_coord[1]] < y_min].index
        aux.drop(indexNames, inplace = True)
        
        #Verificação de quantos equipamentos foram encontrados
        itens = len(aux.index)
        alimentadores = aux['Código do Alimentador'].unique()
        print(alimentadores)
        itens = [len(aux[aux['Código do Alimentador'] == alimentador]) for alimentador in alimentadores]
        
        print(itens)
        
        #Verifica se o lopp tem de ser feito novamente para aumentar o range
        if itens == []:
            print('\n\n\n\LOPPOU')
            print(filtro_inicial)
            if filtro_inicial >= 5000:
                return
            return FiltroDist(df = df, utm_x = utm_x, utm_y = utm_y, passos = passos,filtro_inicial =  filtro_inicial + passos, pontos = pontos, curvas = curvas)
            print(itens)
    
        if min(itens[:2]) > pontos and len(itens) > 1:
            df = deepcopy(aux)
            
            #Cálculo da distância
            df['Distância'] = np.linalg.norm(df[col_coord].sub(np.array([utm_x, utm_y])), axis = 1)
            df['Distância'] = df['Distância']*curvas
            
            return df
        else:
            return FiltroDist(df = df, utm_x = utm_x, utm_y = utm_y, passos = passos,filtro_inicial =  filtro_inicial + passos, pontos = pontos, curvas = curvas)
    ###########################################
    #   ListaEquip: Coverte dataframes em     #
    #   listas, para ser utilizado na função  # 
    #   genesis                               #
    ###########################################  
    def ListaEquip(df, alimentador, numero = 20):#Converte Dataframe em listas
        aux = df[df['Código do Alimentador'] == alimentador]
        tipo, ponto, dist = aux['Tipo de Equipamento'].tolist(), aux['Número Operacional'].tolist(), aux['Distância'].tolist()
        if len(tipo) < numero:
            tamanho = numero - len(tipo)
            vetor = ['' for i in range(tamanho)]
            tipo.extend(vetor)
            ponto.extend(vetor)
            dist.extend(vetor)
        return tipo[:numero], ponto[:numero], dist[:numero]
    
    ###########################################
    # Acha os 2 alimentadores mais perto com base 
    #nos equipamentos
    ###########################################
    def Genesis(utm_x,utm_y):
        mypath = join('C:/Users/joaof/OneDrive/Área de Trabalho/TCC/programa atualizado', 'Relatorios_Genesis')#Pasta com os relatórios
        
        #Variáveis iniciais
        utm_x = float(utm_x.replace(',', '.'))
        utm_y = float(utm_y.replace(',', '.'))
        curvas = 1.4#Constante de aproximação de uma linha reta para um trecho curvo (aproximadamente raiz de 2) 
        filtro_inicial = 10000 #Distância máxima nos eixos X e Y a serem considerados no estudo
        passo = 100
        df_query = {}
        dicionario = {}
        dicionario['Usina'] = {}
        dicionario['Alimentador'] = {}
        pontos = 5
        onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
        equipamentos = [f.replace('.csv', '') for f in onlyfiles]
        tensoes = []
        
        #Cálculo das distâncias e seleção dos resultados
        for equipamento in equipamentos:
            #Criação do dataframe e correção dos dados
            arquivo = join('C:/Users/joaof/OneDrive/Área de Trabalho/TCC/programa atualizado', 'Relatorios_Genesis', f'{equipamento}.csv')
            df_query[equipamento] = pd.read_csv(arquivo, sep = ';', encoding = "ISO-8859-1", index_col = False)
            df_query[equipamento].replace(' ', '', inplace = True)
            df_query[equipamento].replace('', np.nan, inplace = True)
            df_query[equipamento].dropna(inplace = True)
            
            #Renomeando as colunas para agrupar o resultado final
            colunas = df_query[equipamento].columns.to_list()
            df_query[equipamento].columns = [x.split('.')[-1] for x in colunas]
            colunas = df_query[equipamento].columns.to_list()
            df_query[equipamento] = RenomearColuna(df_query[equipamento])
            
            #Inicialização do DataFrame do Alimentador
            if equipamento == 'Alimentador':
                df_alimentadores = deepcopy(df_query['Alimentador'])
                df_alimentadores.set_index('Código do Alimentador', inplace = True)
        
        #Unifica os Dataframes
        df_equip = UnificarDFs(df_query, 'Tipo de Equipamento', ['Alimentador'])
        
        #Converter as colunas para float
        colunas = df_equip.columns
        col_coord = [s for s in colunas if s.__contains__('Coordenada')]
        for coord in col_coord:
            df_equip[coord] = df_equip[coord].astype(float)
        
        #Filtro das distâncias
        df_equip = FiltroDist(df_equip, utm_x, utm_y, passo, filtro_inicial, pontos, curvas)
        df_equip.sort_values('Distância', inplace = True)
        df_equip.reset_index(drop = True, inplace = True)
        df_equip.drop_duplicates(subset = ['Código do Alimentador', 'Número Operacional', 'Tipo de Equipamento'], inplace = True)
        
        #Aquisição dos alimentadores próximos
        alimentadores = df_equip['Código do Alimentador'].unique().tolist()
    
        for i, alimentador in enumerate(alimentadores):
            ali = df_equip.loc[0, 'Código do Alimentador']
            reg = df_alimentadores.loc[ali, 'Nome da Regional']
            tensao = df_alimentadores.loc[ali, 'Tensão Nominal do Alimentador'][4:]
            tipo, ponto, dist = ListaEquip(df_equip, ali, numero = pontos)
            tensoes.append(tensao)
        
            #Gerando os dicionários
            dici = {
                    f'Alimentador {i+1}': ali,
                    f'Tipo Ali {i+1}': tipo,
                    f'Ponto Ali {i+1}': ponto,
                    f'Distância Ali {i+1}': dist}
            dicionario['Usina'].update(dici)
            
            dici = {
                   f'Alimentador {i+1}': ali,
                   f'Regional {i+1}': reg,
                   f'Tensão Nominal {i+1}': tensao}
            dicionario['Alimentador'].update(dici)
            
            #Removendo os dados já utilizados
            df_equip = df_equip[df_equip['Código do Alimentador'] != ali]
            df_equip.reset_index(drop = True, inplace = True)
            
            #Parando no segundo alimentador
            if i == 1:
                continue
        
        return dicionario, alimentadores, tensoes, reg
    
    dicionario, alimentadores, tensoes, reg = Genesis(UTMX,UTMY)
    
    alimentador = alimentadores[0]
    equipamento = dicionario['Usina']['Ponto Ali 1'][0]
    distancia = dicionario['Usina']['Distância Ali 1'][0]
    
    return (alimentador, equipamento, distancia, reg)
