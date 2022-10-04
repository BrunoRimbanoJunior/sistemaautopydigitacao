
import pyautogui
import pandas as pd
from time import sleep
import threading
import PySimpleGUI as sg



def digitar_itens(window, file):
    img_alerta_preco = './IMAGES/alerta_preco.png'
    img_loc_sys_aco = './IMAGES/LOC_SYS_ACO.png'
    img_alerta_qtd = './IMAGES/alerta_qtd.png'
    cp = sg.cprint
    status=''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    cp('Inciando processo de digitacao....')
    cp(file)
    try:
        df = pd.read_excel(file, names=['CODIGO', 'QTD', 'PRECO'])
        lista_pedido = []
        lista_pedido = df.values.tolist()
    except:
        status = 'Erro na leitura do excel'
        

    contador = 0
    total = 0

    loc_bt_inserir = pyautogui.locateOnScreen(img_loc_sys_aco, confidence=0.9)
    #cp(f'Localização do Sistema ACO: {str(loc_bt_inserir)}')
    
    if loc_bt_inserir == None:

        status = 'Sistema Aco não localizado'
        cp(status)
        exit
    else:
        pyautogui.click(loc_bt_inserir)
        
        status = f'ID\tCODIGO\t\tQTD\t VALOR\t\t TOTAL'
        cp(status)

        for itens in lista_pedido:
            if contador >= 60:
                pyautogui.alert('Limite de itens maior que 60')
                contador = 0
            else:
                qtd = ''
                window.write_event_value('-THREAD-', (threading.current_thread().name, qtd))
                
                qtd = itens[1]
                valor = itens[2]
                
                subtotal = qtd*valor
                total = total + subtotal
                try:
                    codigo = int(itens[0])
                except:
                    codigo = str(itens[0])   

                cp(f'{contador + 1:>2} \t{codigo:<14} \t\t{str(itens[1]):>3} \t{float(valor):>6.2f}')
            
                #print(f'{contador + 1} \t{itens[0]} \t{str(itens[1])} \t{(itens[2]):>6.2f}\t\t{subtotal:>6.2f}')
                #pyautogui.write(str(codigo))
                pyautogui.write(str(codigo))
                sleep(0.3)
                pyautogui.press('enter')
                sleep(0.2)
                #digita quantidade do item
                pyautogui.write(str(itens[1]))
                sleep(0.2)
                pyautogui.press('enter')
                sleep(0.3)
                alerta_qtd = pyautogui.locateOnScreen(img_alerta_qtd, confidence=0.8)
                sleep(0.5)
                if alerta_qtd != None:
                    pyautogui.press('enter')
                sleep(1)
                pyautogui.write(str(valor))
                sleep(0.2)
                pyautogui.press('enter')
                sleep(0.5)
                alertas = pyautogui.locateOnScreen(img_alerta_preco, confidence=0.8)
                sleep(1)
               
                if alertas!=None:
                    pyautogui.press('enter')
                    
                sleep(0.5)
                pyautogui.press('insert')
                sleep(0.5)
                contador += 1
        pyautogui.press('esc')    

    print(f'Valor total do pedido:    \t\t    {total:7.2f}')       


if __name__ == '__main__':
    img_loc_sys_aco = './IMAGES/LOC_SYS_ACO.png'
    loc_bt_inserir = pyautogui.locateOnScreen(img_loc_sys_aco, confidence=0.9)
    print(f'Localização do Sistema ACO: {str(loc_bt_inserir)}')
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

        elif event.startswith('Executar'):
            if loc_bt_inserir == None:
                print('Sistema Aco não localizado')
                exit
            else:
                pyautogui.click(loc_bt_inserir)
                file = './arquivos/transferencia.xlsx'
                threading.Thread(target=digitar_itens, args=(window, file), daemon=True).start()
        
