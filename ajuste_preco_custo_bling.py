import pyautogui
import pandas as pd
from time import sleep
import threading
import PySimpleGUI as sg
from login_bling import start



def digitar_itens(window, file):
    
    cp = sg.cprint
    status=''
    lista_produtos = []
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    cp('Inciando processo de digitacao....')
    cp(file)
    try:
        df = pd.read_excel(file, names=['CODIGO', 'VALOR'])
        lista_produtos = df.values.tolist()
    except:
        status = 'Erro na leitura do excel'
        

    contador = 0
    

    start.iniciar()
    sleep(5)
    start.cad_produtos()
    cp("Acessando cadastro de produtos...")
      
    status = f'CODIGO\tVALOR'
    cp(status)
    

    for itens in lista_produtos:
      
        codigo = itens[0]
        valor = str(itens[1])
                        
        cp(f'{contador + 1:>2} \t{codigo:<14} \t\t{valor:>6}')
    
        start.alterar_preco_custo_bling(codigo, valor)
        contador+=1
                
        cp('Processo Finalizado')  

if __name__ == '__main__':
  
    cp = sg.cprint
    size_btn = (50,1)
    layout = [[sg.Push(), sg.Button('Executar', button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()]]
    layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(70,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
    layout += [[sg.Push(), sg.Cancel('Sair', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],
    window = sg.Window('Teste...', layout, finalize=True, keep_on_top=False,)
               
    while True:
                
        event, values = window.read()
        #cp(event, values)
        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event.startswith('Executar'):
            file = './arquivos/tab_preco_iventario.xlsx'
            threading.Thread(target=digitar_itens, args=(window, file), daemon=True).start()
    
