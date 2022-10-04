import tabula, PyPDF2
import pandas as pd
from datetime import datetime
from tkinter import Tk, filedialog as dlg
import numpy as np
import os
#from login_bling import start
import PySimpleGUI as sg
import threading


def converter_estoque_tabula(window, caminho):


    cp = sg.cprint
    status=''
    window.write_event_value('-CONV-', (threading.current_thread().name, status))
    cp('Inciando processo de converção....')
    cp(caminho)

    data = datetime.now().strftime('%d%m%Y_%H%M')
    cp(f'Data: {data}')
    ca = caminho.split('/')
    sub = len(ca)
    file = ca[sub-1]
    dire = ca[0]
    cp(f'Arquivo selecionado: {file}')
    
    
    for di in range(1, sub-1):
        dire = dire +'/'+ca[di]

    temp = dire +'/'+'temp.csv'
    tabula.convert_into(caminho, temp, output_format='csv', pages='all')
  

    #identificando o numero de paginas
    ultima_pagina = PyPDF2.PdfFileReader(caminho).getNumPages()
    cp(f'Numero de Paginas dentro do arquivo: {ultima_pagina}')

    #ini = 1
    #eend = 0
    #numero_de_paginas = 1
    cp('Em processamento, converterndo arquivo pdf')
  
    df = pd.read_csv(temp, names=['CODIGO','DESCRICAO','PGM','PRODUZIDO','SALDO PGM','ESTOQUE','CARTEIRA','SEPARACAO','DISPONIVEL','NECESSIDADE'], encoding='latin-1')    
    
    df2 = df.sort_values(by=['CODIGO'])
    cp('Limpando dados coluna DESCRICAO')
    df2['DESCRICAO'].replace('', np.nan, inplace=True)
    df2.dropna(subset=['DESCRICAO'], inplace=True)
    cp('Limpando dados coluna CODIGO')
    df2['CODIGO'].replace('', np.nan, inplace=True)
    df2['CODIGO'].replace('0', np.nan, inplace=True)
    df2.dropna(subset=['CODIGO'], inplace=True)
    filtro = df2['CODIGO'] != 'Classe'
    df2 = df2[filtro]
    filtro = df2['CODIGO'] != 'PRODUTO'
    df2 = df2[filtro]
    cp('Excluindo Dados Desnecessarios')
    df2 = df2.drop(columns=['DESCRICAO', 'PGM', 'PRODUZIDO','SALDO PGM','ESTOQUE','CARTEIRA','SEPARACAO', 'NECESSIDADE'])
    cp('Convertendo em numero')
    df2['DISPONIVEL'] = pd.to_numeric(df2['DISPONIVEL'], errors='coerce', downcast='integer')
    dire = dire +'/'+'ESTOQUE_'+data+'.xlsx'
    cp(f'Caminnho para salvar arquivo: {dire}')
    #CASO QUEIRA SALVAR EM UMA TABELA EXCEL JA EXISTENTE, USAR AS LINHAS COMENTADAS, AJUSTANDO A VARIAVEL DIRE PARA A PLANILHA CORRETA
    #table = pd.ExcelWriter(dire)
    cp('Convertendo dados para arquivo Excel')
    df2.to_excel(dire, sheet_name='ESTOQUE_FABRICA',index=False )
    cp('Processo Finalizado....')
    os.remove(temp)



if __name__ == '__main__':
    cp = sg.cprint
    size_btn = (50,1)
    layout = [[sg.Push(), sg.Text('', key='status'),sg.Push()], [sg.Push(), sg.Input(size=(50,1)), sg.FileBrowse(), sg.Push()]]
    layout += [[sg.Push(), sg.Button('Executar', button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()]]
    layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(55,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
    layout += [[sg.Push(), sg.Cancel('Sair', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],
    window = sg.Window('Teste...', layout, finalize=True, keep_on_top=False,)
               
    while True:
                
        event, values = window.read()
        #cp(event, values)


        if event == sg.WIN_CLOSED or event == 'Sair':
            break

        elif event.startswith('Executar'):
            cp(values[0])
            threading.Thread(target=converter_estoque_tabula, args=(window, values[0]), daemon=True).start()

   