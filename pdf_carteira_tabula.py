import tabula
import pandas as pd
import numpy as np
from tkinter import filedialog
from datetime import datetime
import os
import PySimpleGUI as sg
import threading


def converter_carteira(window, caminho):

   cp = sg.cprint
   status=''
   window.write_event_value('-CONV-', (threading.current_thread().name, status))
   cp('Inciando processo de converção....')
   cp(caminho)

   data = datetime.now().strftime('%d%m%Y_%H%M')
   ca = caminho.split('/')
   sub = len(ca)
   file = ca[sub-1]
   dire = ca[0]
   print(f'Arquivo selecionado: {file}')
   
   for di in range(1, sub-1):
      dire = dire +'/'+ca[di]

   temp = dire +'/'+'temp.csv'
   tabula.convert_into(caminho, temp, output_format='csv', pages='all', lattice=True)
   print('Convertendo arquivo PDF para CSV')
   print('Renomeando Colunas')

   df = pd.read_csv(temp, names=['IT','PRODUTO','DESCRICAO','PRECO LIQ','QTD','SALDO','VALOR TOTAL SALDO',
                                          'QTD SEP','VALOR TOTAL SEP'], encoding='latin-1')
   #retirar as linhas que nao contem R, o na, é para as linhas NaN, por ser um valor nulo, pode gerar erro                                       
   print('Limpando dados inconsistentes')
   df2 = df[df['PRODUTO'].str.contains('R', na=False)]  
   df2.dropna()                                     
   df2 = df.sort_values(by=['IT'])
   df2 = df2.drop(columns=['IT', 'DESCRICAO', 'PRECO LIQ', 'QTD','VALOR TOTAL SALDO', 'VALOR TOTAL SEP'])
   df2['SALDO'] = df2['SALDO'].replace(',','.', regex=True)
   df2['QTD SEP'] = df2['QTD SEP'].replace(',','.', regex=True)
   df2['SALDO'] = pd.to_numeric(df2['SALDO'], errors='coerce')
   df2['QTD SEP'] = pd.to_numeric(df2['QTD SEP'], errors='coerce')
   
   #somar produtos
   print('Somando saldos....')
   saldo_prod =  df2[['PRODUTO','SALDO','QTD SEP']].groupby('PRODUTO').sum()
   dire = dire +'/'+'CARTEIRA_'+data+'.xlsx'
   table = pd.ExcelWriter(dire)
   
   saldo_prod.to_excel(table, index=True )
   #saldo_prod.to_excel(table, index=True )
   table.save()
   #excluir arquivo csv
   os.remove(temp)
   print('Processo Finalizado')
   

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
            threading.Thread(target=converter_carteira, args=(window, values[0]), daemon=True).start()








