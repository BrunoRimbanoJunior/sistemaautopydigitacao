import shutil
import os
import threading
import PySimpleGUI as sg
from time import sleep

cp = sg.cprint


def salvar_arq_estoque_pasta_local(window):
    
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    source = r'F:\DONWLOAD'
    destination = './ESTOQUE'
    files = os.listdir(source)
    files_carteira = os.listdir(destination)
    status = 'Executando Processos Internos....'
    cp(status)
    
    for file in files:
        arq = file.split('.')
        sub = len(arq)
        extensao = arq[sub-1]
        if extensao =='csv':
            shutil.move(f'{source}/{file}', f'{destination}/estoque.csv')
            cp(f'Arquivo de estoque salvo: {file}')
        
            
def excluir(diretorio):
    
    files = os.listdir(diretorio)
    for file in files:
         os.remove(f'{diretorio}/{file}')
    
def limpar_sistema(window):


    tarefa=''
    window.write_event_value('-THREAD-', (threading.current_thread().name, tarefa))
    cp('Iniciando processo de limpeza')
    sleep(1)
    pasta_pedidos = './PEDIDOS'
    pasta_estoque = './ESTOQUE'
    pasta_carteira = './CARTEIRA'  
    excluir(pasta_pedidos)
    tarefa = 'Excluindo arquivos da pasta Pedidos'
    
    cp('Excluindo arquivos da pasta Pedidos')
    sleep(1)
    excluir(pasta_carteira)
    #cp('Excluindo arquivos da pasta Carteira')
    tarefa = 'Excluindo arquivos da pasta Carteira'
    cp(tarefa)
    sleep(1)
    excluir(pasta_estoque)
    tarefa = 'Excluindo arquivos da pasta Estoque'
    cp(tarefa)
    sleep(1)
    cp('Processo Finalizado......')


if __name__ == '__main__':
    cp = sg.cprint
    size_btn = (50,1)
    layout = [[sg.Push(), sg.Button('Executar', button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()]]
    layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(55,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
    layout += [[sg.Push(), sg.Cancel('Sair', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],
    window = sg.Window('Teste...', layout, finalize=True, keep_on_top=False,)
               
    while True:
                
        event, values = window.read()
        #cp(event, values)


        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event == 'Executar':
            salvar_arq_estoque_pasta_local(window) 

