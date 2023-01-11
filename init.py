import PySimpleGUI as sg


def init():
    sg.theme('DarkBlue17')
    
    layout = [[sg.Text("Solicitação",pad=((4,1))),sg.In(size=(20,1),key="Solicitação")],
              [sg.Text("UTM_X",pad=((4,1))),sg.In(size=(15,10),key="utm_x")],
              [sg.Text("UTM_Y",pad=((4,1))),sg.In(size=(15,10),key="utm_y")],
              [sg.Text("Potência",pad=((4,1))),sg.In(size=(15,1),key="potencia")],
              [sg.Submit('Pesquisar'), sg.Cancel('Cancelar')]]
    
    window = sg.Window('ID Alimentador', layout)
    
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Cancelar"):
            break
        elif event == 'Pesquisar':
            SO = str(values["Solicitação"])
            utm_x = str(values['utm_x'])
            utm_y = str(values['utm_y'])
            potencia = str(values["potencia"])
    
    # Finish up by removing from the screen
    window.close()
    return utm_x, utm_y, potencia, SO


