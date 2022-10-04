
from integracao_planilhas import integracao_matriz
import pandas as pd
from datetime import datetime
import PySimpleGUI as sg
import threading


def salvar_carteira(window):
    cp = sg.cprint
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    status = 'Copiando dados do Modulo'
    cp(status)

    planilha = integracao_matriz()
    lista = []
    data = datetime.now().strftime('%d%m%Y_%H%M')
    file = r'./CARTEIRA/CARTEIRA_RETROGLASS_'+data+'.xlsx'
    cp(file)
    for row in planilha:
        if int(row[8])>0:
            lista.append([row[0], row[2], row[3], row[8], row[5]])
    #print(lista)

    df = pd.DataFrame(lista, columns= ['NOME','OC', 'CODIGO', 'SALDO', 'VALOR'])
    df['SALDO'] = pd.to_numeric(df['SALDO'], errors='coerce')
    df['VALOR'] = df['VALOR'].str.replace(',', '.')
    df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')
#print(df)
    df.to_excel(file, index=False)        
    cp('Carteira Salva........')


if __name__ == '__main__':
    size_btn = (50,1)
    layout = [[sg.Push(), sg.Button('Executar', button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()]]
    layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(55,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
    layout += [[sg.Push(), sg.Cancel('Sair', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],
    window = sg.Window('Limpeza de Sistema', layout, finalize=True, keep_on_top=False,)
               
    while True:
                
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event.startswith('Executar'):
            threading.Thread(target=salvar_carteira, args=(window,), daemon=True).start()  