import PySimpleGUI as sg
from dig_itens_pedido import digitar_itens_bling
from dig_itens_pedido_ACO import digitar_itens
from salvar_carteira import salvar_carteira
from separacao_modulo import fazer_separacao
from salvar_pedidos import salvar_pedidos
from manipulacao_de_arquivos_locais import limpar_sistema
from login_bling import auto_estoque
from time import sleep
import threading
from pdf_estoque_tabula import converter_estoque_tabula
from pdf_carteira_tabula import converter_carteira


#usando para depuracao, mostra os prints em uma janela
#print = sg.Print

sg.theme('Dark Grey 13')
#sg.theme('System Default 1')

progress = 'Aguardando...'
mult_line = ''
size_btn = (45,1)
cp = sg.cprint



def str_out(window):
    str = 'Teste de Funcao 1'

    window.write_event_value('-THREAD-', (threading.current_thread().name, str))      # Data sent is a tuple of thread name and counter
    cp(str, c='white on green')
    sleep(1)
    window.write_event_value('-THREAD-', (threading.current_thread().name, str))      # Data sent is a tuple of thread name and counter
    cp(str, c='white on green')


def constr_btn(btn_text):
    return [sg.Push(), sg.Button(btn_text, button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()],


##------------------------------------LAYOUT----------------------------------------##
list_btn = ['LANÇAR PEDIDO NO BLING VIA EXCEL',
            'LANÇAR PEDIDO NO ACO VIA EXCEL',
            'SINCRONIZAR ESTOQUES',
            'SEPARAR ITENS DISPONIVEIS',
            'SALVAR CARTEIRA',
            'LANÇAR PEDIDOS NO BLING',
            'LIMPAR SISTEMA',
            'CONVERTER ESTOQUE',
            'CONVERTER CARTEIRA',]

layout = [[sg.Push(), sg.Text('', key='status'),sg.Push()], [sg.Push(), sg.Input(size=(50,1)), sg.FileBrowse(), sg.Push()]]
layout += [[constr_btn(t) for t in(list_btn)],]
layout += [[sg.Push(), sg.Cancel('SAIR', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],                           
              

#layout_l = [[sg.Text('', key='status')]]    
layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(100,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]

window = sg.Window('Modulo para controle de carteira', layout,finalize=True,keep_on_top=False)

#window['-PRINT-'+sg.WRITE_ONLY_KEY].print(progress, end='', text_color='red', background_color='yellow')


status = ''
cp('Iniciando sistemas....')
cp('Selecione uma operacao.....')



while True:
      
    window['status'].update(status)
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == 'SAIR':
        break
    
    #-------------------------------Funcoes Proprias-------------------------------------#  

    #--------------------DIGITACAO DE PEDIDOS SISTEMA BLIN VIA PYAU-----------------------#     
    elif event == 'LANÇAR PEDIDO NO BLING VIA EXCEL':
        if values[0]!='':
            
            threading.Thread(target=str_out, args=(window,), daemon=True).start()
            cp('Abrindo arquivo excel....')

            try:
                cp('Digitando.....')
                digitar_itens_bling(values[0])
                cp('Limpando Arquivos Temporarios')
                values[0]=''
                cp('Pedido Finalizado...')
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
                threading.Thread(target=digitar_itens, args=(window, values[0]), daemon=True).start()
                #digitar_itens(values[0])
                
                values[0]='' 
                cp('Pedido finalizado')
            except:    
                cp('Erro na digitação')
        else:
            status = 'Selecione um arquivo excel'
            window['status'].update(status)
            cp('Selecione um arquivo...')

    #----------------------------SALVAR CARTEIRA EM ARQUIVO EXCEL--------------------------#

    elif event.startswith('SALVAR CARTEIRA'):
        status = ''
        threading.Thread(target=salvar_carteira, args=(window,), daemon=True).start()
      
     #----------------------SEPARACAO INTES DISPONIVEIS EM ESTOQUE------------------------#

    elif event.startswith('SEPARAR ITENS DISPONIVEIS'):
        threading.Thread(target=fazer_separacao, args=(window,), daemon=True).start()
    
     #---------------------------DIGITAR PEDIDOS SISTEMA BLING----------------------------#

    elif event.startswith('LANÇAR PEDIDOS NO BLING'):
        threading.Thread(target=salvar_pedidos, args=(window,), daemon=True).start()
        #window.perform_long_operation(salvar_pedidos(window), '-FUNCTION COMPLETED-')

    #-----------------------LIMPAR PASTAS COM ARQUIVOS TEMPORARIOS------------------------#

    elif event.startswith('LIMPAR SISTEMA'):
        threading.Thread(target=limpar_sistema, args=(window,), daemon=True).start()
       
     #--------------------COPIAR DADOS DE ESTOQUE DO BLING PARA O MODULO------------------------#

    elif event.startswith('SINCRONIZAR ESTOQUES'):
        threading.Thread(target=auto_estoque, args=(window,), daemon=True).start()
     
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