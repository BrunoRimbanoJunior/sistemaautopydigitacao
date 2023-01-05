import PySimpleGUI as sg
from dig_itens_pedido import digitar_itens_bling
from dig_itens_pedido_ACO import digitar_pedidos_ACO
from salvar_carteira import salvar_carteira
from separacao_modulo import fazer_separacao

from manipulacao_de_arquivos_locais import limpar_sistema
from login_bling import auto_estoque
from time import sleep
import threading
from pdf_estoque_tabula import converter_estoque_tabula
from pdf_carteira_tabula import converter_carteira
from dig_pedidos_dados_google import dig_pedidos_google
from ajuste_estoque_aco import ajuste_estoque
from dig_pedido_bling_seleniun import digitar_pedidos_online


#usando para depuracao, mostra os prints em uma janela
#print = sg.Print

sg.theme('Dark Grey 13')
#sg.theme('System Default 1')

progress = 'Aguardando...'
mult_line = ''
size_btn = (35,1)
cp = sg.cprint
use_custom_titlebar = False





##------------------------------------LAYOUT----------------------------------------##
list_btn_r = ['AJUSTE DE ESTOQUE',
            'SINCRONIZAR ESTOQUES',
            'SEPARAR ITENS DISPONIVEIS',
            'SALVAR CARTEIRA',
            'LANÇAR PEDIDOS NO BLING',
            'LIMPAR SISTEMA',
            'CONVERTER ESTOQUE',
            'CONVERTER CARTEIRA',
            'DIGITAR PEDIDOS TAB PEDIDOS',
            'DIGITAR PEDIDOS TAB SEPARACAO']            


#def constr_btn(btn_text):
#    return [sg.Push(), sg.Button(btn_text, button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],            

layout_l = [[sg.Push(), sg.Text('', key='status'),sg.Push()], [sg.Push(), sg.Input(size=(37,1)), sg.FileBrowse(), sg.Push()]]

layout_l += [[sg.Push(), sg.Button(list_btn_r[0], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()], 
            [sg.Push(), sg.Button(list_btn_r[9], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],
            [sg.Push(), sg.Button(list_btn_r[8], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],
            [sg.Push(), sg.Button(list_btn_r[1], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()], 
            [sg.Push(), sg.Button(list_btn_r[2], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],
            [sg.Push(), sg.Button(list_btn_r[3], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],
            [sg.Push(), sg.Button(list_btn_r[4], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()], 
            [sg.Push(), sg.Button(list_btn_r[5], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()], 
            [sg.Push(), sg.Button(list_btn_r[6], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],
            [sg.Push(), sg.Button(list_btn_r[7], button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],
            
            [sg.Push(), sg.Cancel('SAIR', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()],] 
            
layout_r = [[sg.Push(), sg.Multiline(key='-PRINT-', size=(80,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
              
layout = [[sg.Column(layout_l), sg.Column(layout_r)]]


window = sg.Window('Modulo para controle de carteira', layout,finalize=True,keep_on_top=False)

#window['-PRINT-'+sg.WRITE_ONLY_KEY].print(progress, end='', text_color='red', background_color='yellow')


status = ''
cp('Iniciando sistemas....')
threading.Thread(target=limpar_sistema, args=(window,), daemon=True).start()
cp('Selecione uma operacao.....')


#funcao de teste
def str_out(window):
    str = 'Teste de Funcao 1'

    window.write_event_value('-THREAD-', (threading.current_thread().name, str))      # Data sent is a tuple of thread name and counter
    cp(str, c='white on green')
    sleep(1)
    window.write_event_value('-THREAD-', (threading.current_thread().name, str))      # Data sent is a tuple of thread name and counter
    cp(str, c='white on green')



while True:
      
    window['status'].update(status)
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == 'SAIR':
        break
    
    #-------------------------------Funcoes Proprias-------------------------------------#  
    

    #--------------------DIGITACAO DE PEDIDOS SISTEMA BLIN VIA PYAU-----------------------#     
    elif event == 'DIGITAR PEDIDOS TAB SEPARACAO':
        status = ''
        threading.Thread(target=digitar_pedidos_ACO, args=(window,), daemon=True).start()

    #------------------------DIGITACAO DE PEDIDOS SISTEMA ACO -----------------------------#     
    #    
    elif event == 'DIGITAR PEDIDOS TAB PEDIDOS':
        status = ''
        threading.Thread(target=dig_pedidos_google, args=(window,), daemon=True).start()

    #----------------------------SALVAR CARTEIRA EM ARQUIVO EXCEL--------------------------#

    elif event.startswith('SALVAR CARTEIRA'):
        status = ''
        threading.Thread(target=salvar_carteira, args=(window,), daemon=True).start()
      
     #----------------------SEPARACAO INTES DISPONIVEIS EM ESTOQUE------------------------#

    elif event.startswith('SEPARAR ITENS DISPONIVEIS'):
        threading.Thread(target=fazer_separacao, args=(window,), daemon=True).start()
    
     #---------------------------DIGITAR PEDIDOS SISTEMA BLING----------------------------#

    elif event.startswith('LANÇAR PEDIDOS NO BLING'):
        # threading.Thread(target=salvar_pedidos, args=(window,), daemon=True).start()
        #window.perform_long_operation(salvar_pedidos(window), '-FUNCTION COMPLETED-')
        threading.Thread(target=digitar_pedidos_online, args=(window,), daemon=True).start()

    #-----------------------LIMPAR PASTAS COM ARQUIVOS TEMPORARIOS------------------------#

    elif event.startswith('LIMPAR SISTEMA'):
        threading.Thread(target=limpar_sistema, args=(window,), daemon=True).start()
       
     #--------------------COPIAR DADOS DE ESTOQUE DO BLING PARA O MODULO------------------------#

    elif event.startswith('SINCRONIZAR ESTOQUES'):
        threading.Thread(target=auto_estoque, args=(window,), daemon=True).start()

    #--------------------AJUSTE DE ESTOQUE VIA MANUTENCAO ACO------------------------#

    elif event.startswith('AJUSTE DE ESTOQUE'):
        if values[0]!='':
            cp('Digitador ACO trabalhando')
            
            try:
                cp('Digitando itens.....')
                threading.Thread(target=ajuste_estoque, args=(window, values[0]), daemon=True).start()
                values[0]='' 
                cp('Ajustes finalizados')
            except:    
                cp('Erro na digitação')
        else:
            status = 'Selecione um arquivo excel'
            window['status'].update(status)
            cp('Selecione um arquivo...')    
    #------------------------DIGITACAO DE PEDIDOS SISTEMA ACO -----------------------------#     
    #    
    elif event == 'LANÇAR PEDIDO NO ACO VIA EXCEL':
        if values[0]!='':
            cp('Digitador ACO trabalhando')
            
            try:
                cp('Digitando pedido.....')
                threading.Thread(target=digitar_pedidos_ACO, args=(window, values[0]), daemon=True).start()
                values[0]='' 
                cp('Pedido finalizado')
            except:    
                cp('Erro na digitação')
        else:
            status = 'Selecione um arquivo excel'
            window['status'].update(status)
            cp('Selecione um arquivo...')    
     
     #--------------------------BOTAO PARA TESTE TEMPORARIO--------------------------------------#
         
    elif event.startswith('CONVERTER ESTOQUE'):
        if values[0]!='':
            cp('Conversor Selecionado')
            
            try:
                cp('Inciando processo')
                threading.Thread(target=converter_estoque_tabula, args=(window, values[0]), daemon=True).start()
               
                
                values[0]='' 
                cp('Conversao Finalizada')
            except:    
                cp('Erro na importcao')
        else:
            status = 'Selecione um arquivo pdf'
            window['status'].update(status)
            cp('Selecione um arquivo...')

    elif event.startswith('CONVERTER CARTEIRA'):
        if values[0]!='':
            cp('Conversor Selecionado')
            
            try:
                cp('Inciando processo')
                threading.Thread(target=converter_carteira, args=(window, values[0]), daemon=True).start()
               
                
                values[0]='' 
                cp('Conversao Finalizada')
            except:    
                cp('Erro na importcao')
        else:
            status = 'Selecione um arquivo pdf'
            window['status'].update(status)
            cp('Selecione um arquivo...')        
        
    elif event == '-FUNCTION COMPLETED-':
        cp('Processo finalizado......')

    values[0]=''         
    progress = ''
   

window.close(); del window