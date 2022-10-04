from integracao_planilhas import integracao_separacao
import pandas as pd
import numpy as np
from datetime import datetime
from dig_pedido_bling_seleniun import digitar_pedidos_online
import threading
import PySimpleGUI as sg


def salvar_pedidos(window):
    cp = sg.cprint
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    status = 'Inciando processo de digitação de pedidos'
    cp(status)
    cp('Baixando dados do Modulo......')


    planilha = integracao_separacao()
    lista = []
    data = datetime.now().strftime('%d%m%Y_%H%M')
    file = './PEDIDOS/PEDIDOS_RETROGLASS_'+data+'.xlsx'
    for row in planilha:
        if row[11]=='0':
            lista.append([row[7], row[3], row[4], row[5], row[8], row[11]])

    

    df = pd.DataFrame(lista, columns= ['CLIENTE','CODIGO', 'SEPARACAO', 'PRECO VENDA', 'PED', 'STATUS'])
    df['SEPARACAO'] = pd.to_numeric(df['SEPARACAO'], errors='coerce')
    df['PRECO VENDA'] = df['PRECO VENDA'].str.replace(',', '.')
    df['PRECO VENDA'] = pd.to_numeric(df['PRECO VENDA'], errors='coerce')

    df.to_excel(file, index=False)
    cp('Dados baixados....')
#modulo responsavel em digitar os pedidos no bling
    # digita = input('Digitar pedidos? S/N').upper()
    # if digita == 'S':
    #     digitar_pedidos_online(file)
    # else:
    #     exit

    # return file
    cp('Iniciando conexao com Sistema Bling.....')
    #*********************alteração temporaria para teste*******************************
    threading.Thread(target=digitar_pedidos_online, args=(window, file,), daemon=True).start()
    
    #digitar_pedidos_online(file)
    cp('Pedidos Finalizados.........')
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
        elif event == 'Executar':
            salvar_pedidos(window)    