
from integracao_planilhas import integracao_pedidos
from dig_itens_pedido_ACO import digitar_itens_tab_pedidos
import pandas as pd
from datetime import datetime
import PySimpleGUI as sg
import threading



def dig_pedidos_google(window):
    cp = sg.cprint
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    status = 'Copiando dados do Modulo'
    cp(status)

    planilha = integracao_pedidos()
    lista = []
    data = datetime.now().strftime('%d%m%Y_%H%M')
    file = r'./arquivos/pedidos.xlsx'
    cp(file)
    for row in planilha:
        lista.append([row[0], row[1], row[2]])
   

    df = pd.DataFrame(lista, columns= ['CODIGO','QTD', 'VALOR'])
    df['QTD'] = pd.to_numeric(df['QTD'], errors='coerce')
    df['VALOR'] = df['VALOR'].str.replace(',', '.')
    df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')

    df.to_excel(file, index=False)        
    cp('PEDIDO SALVO........')
    
    cp('Iniciando conexao com ACO')
    #*********************alteração temporaria para teste*******************************
    threading.Thread(target=digitar_itens_tab_pedidos, args=(window, file,), daemon=True).start()
    

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
            threading.Thread(target=dig_pedidos_google, args=(window,), daemon=True).start()  