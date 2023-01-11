import PySimpleGUI as sg
from programa_completo import programa_IA

def PromptUpdate(string):
    window['-OUTPUT-'].update(string) 
    window.refresh()
    
# def front():
sg.theme('DarkBlue17')
opt_fnt = ["","Resposta em Mais de 30 Dias","Micro na Aba Minigeração","Faltam Dados"]

layout = [
          [sg.Text("Nome da Solicitação: ",pad=(9,1)),sg.In(size=(22,1), enable_events=False,key="Nome_SO",default_text='Usina Fotovoltaica CTC')],
          [sg.Text("Número da Solicitação:",pad=(5 ,1)),sg.In(size=(15,1), enable_events=False,key="SO",default_text='603720')],
          [sg.Text("Alimentador Específico:",pad=(5 ,1)),sg.In(size=(22,1), enable_events=False,key="alimentadorespec",default_text='alimentador mais próximo')],
          [sg.Text("Dados da Consulta de Acesso: ",size=(35,1), key='-OUTPUT-')],
          [sg.Text("UTM X:",pad=((7,1))),sg.In(size=(12,10),key="utm_x",default_text='745059.16'), sg.Text("UTM Y:",pad=((7,1))),sg.In(size=(12,10),key="utm_y",default_text='6944633.11')],
          [sg.Text("Potência (MW):",pad=((2,1))),sg.In(size=(29,1),key="potencia",default_text='1')],
          [sg.Text("",size=(4,2)),sg.Button((' 01. Emitir a Informação '), size=(35,1))],
          [sg.Text("",size=(15,2)), sg.Button(('Sair'), size=(10,1))],
]

window = sg.Window('Programa Info. de Acesso', layout)

# def acesso_PEP(num_so):
#     UTMX = str(values['utm_x'])
#     UTMY = str(values['utm_y'])
#     potencia = float(values['potencia'])
#     return UTMX, UTMY, potencia

while True:
    event, values = window.read()
    # if event == ' 01. Emitir a Informação ' and values['SO'] == "":
    #     window['-OUTPUT-'].update('Não foi inserido número de Solicitação!')
    
    if event == ' 01. Emitir a Informação ' and values['SO'] != "" :
#        from AbreSO import *
        SO_automatica = int(values['SO'])
        UTMX = str(values['utm_x'])
        UTMY = str(values['utm_y'])
        potencia_conexao = float(values['potencia'])
        nome_so = str(values["Nome_SO"])
        alimentadorespec = str(values["alimentadorespec"]).upper()
        alimentadorespec = alimentadorespec.replace('-','')
        # try:
        programa_IA(nome_so, SO_automatica, UTMX, UTMY, potencia_conexao,alimentadorespec)
        # except:
        #     msg = 'Erro ao Abrir a SO'
        #     PromptUpdate(msg)
    if event == 'Sair':    # Window close button event
        break
    if event == sg.WIN_CLOSED:    # Window close button event
        break
window.close()

# front()