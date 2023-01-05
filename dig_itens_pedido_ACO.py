
import pyautogui
import pandas as pd
from time import sleep
import threading
import PySimpleGUI as sg
from salvar_pedidos import salvar_pedidos



def digitar_pedidos_ACO(window):
    
    file = salvar_pedidos(window)
    cp = sg.cprint
    


    status = ''
    window.write_event_value(
        '-THREAD-', (threading.current_thread().name, status))

    try:
        df = pd.read_excel(
            file, names=['CLIENTE', 'CODIGO', 'SEPARACAO', 'PRECO_VENDA', 'OC', 'STATUS', 'COD_CLIENTE', 'COD_VENDEDOR'])
        lista_pedidos = df.values.tolist()
        lista_pedidos2 = df.values.tolist()
        cp('IMPORTAÇÃO EXCEL COM SUCESSO\n')
    except:
        cp('********Erro na leitura do excel******')

    list_cliente = []
    list_cod_cliente = []
    list_oc = []
    list_vendedor = []

   

    for cliente in lista_pedidos:
        if str(cliente[0]) in list_cliente:
            pass
        else:
            # da lista principal, cria uma lista com o nome dos clientes
            list_cliente.append(str(cliente[0]))
            list_cod_cliente.append(str(cliente[6]))
            list_vendedor.append(str(cliente[7]))
          

        if str(cliente[4]) in list_oc:
            pass
        else:
            # da lista pincipal cria uma lista com o numero dos pedidos
            list_oc.append(str(cliente[4]))

    oc_ant = ''
    oc = ''
    cont = 1

    # passar por todos os clientes da lista
    for cliente in list_cliente:
        # novo cliente novo pedido
        pyautogui.alert('NOVO CLIENTE: '+ str(cliente))
        id = list_cliente.index(cliente)
        cp(f'CLIENTE: {cliente}   COD: {list_cod_cliente[id]}\n')
        cp(f'ID\tCODIGO        \t\tQTD\t\tVALOR  \t\t  ORDEM DE COMPRA')
        sleep(1)
        # funcao para novo cliente que recebe o codigo do cliente
        novo_pedido(window, str(list_cod_cliente[id]), list_vendedor[id], list_oc[id])

        cont = 1
        total = 0
        subtotal = 0

        # passar pela lista principal inteira
        for item in lista_pedidos2:
            # se o cliente for igual ao registro cliente...
            if str(item[0]) == cliente:
                # grava o numero da oc atual
                oc = str(item[4])
                # compara o numero das ordens, se for igual ou for o primeiro item imprime
                if oc == oc_ant or cont == 1:
                    
                    sleep(0.2)
                    qtd = float(item[2])
                    valor = float(item[3])
                    subtotal = qtd * valor
                    total = total + subtotal
                    #funcao para digitar itens, parametros CODIGO, QTD, VALOR
                    codigo = item[1]
                    
                    cp(f'{str(cont)}\t{str(codigo):14} \t\t{str(qtd):>3} \t\t{str(valor):>7} \t\t{str(item[4]):10}')
                    
                    digitar_itens(window, codigo, qtd, valor)
                    sleep(1)
                    
                    oc_ant = str(item[4])
                    #print(f'Caso Igual: {oc} - {oc_ant}')
                    cont += 1
                    if cont >= 60:
                        finalizar_pedido()
                        novo_pedido(window, str(list_cod_cliente[id]), list_vendedor[id], oc)
                        cp(f'Valor total.......................{total:5.2f}')
                        cp('Limite de itens atendido......')
                        cont = 1

                else:
                    cp(f'Valor total.......................{total:5.2f}')
                    
                    sleep(2)
                    #finalizar pedido
                    finalizar_pedido(window)
                    sleep(2)
                
                    cp('\n')
                    cp('*************NOVA ORDEM DE COMPRA***************')
                    total = 0
                    subtotal = 0
                    cont = 1
                    # login_bling.start.iniciar()
                    # sleep(5)
                    novo_pedido(window, str(list_cod_cliente[id]), list_vendedor[id], oc)
                    cp(f'ID\tCODIGO        \t\tQTD\t\tVALOR  \t\t  ORDEM DE COMPRA')
                    codigo = str(item[1])
                    qtd = float(item[2])
                    valor = float(item[3])
                    subtotal = qtd * valor
                    total = total + subtotal
                    digitar_itens(window, codigo, qtd, valor)
                    sleep(1)
                    cp(f'{str(cont)}\t{str(item[1]):14} \t\t{str(item[2]):>3} \t\t{str(item[3]):>7} \t\t{str(item[4]):10}')
                    oc_ant = str(item[4])
                    cont += 1

        cp(f'Valor total.......................{total:5.2f}')
        cp('\n')
        finalizar_pedido(window)
    cp('Pedidos finalizados..........')    
    
####**************************************************************###############################
def digitar_itens_tab_pedidos(window, file):
    img_alerta_preco = './IMAGES/alerta_preco.png'
    img_loc_sys_aco = './IMAGES/LOC_SYS_ACO.png'
    img_alerta_qtd = './IMAGES/alerta_qtd.png'
    cp = sg.cprint
    status=''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    #fazer importacao dos dados da plan google
    
    
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

def digitar_itens(window, codigo, qtd, valor):
    cp = sg.cprint
    
    img_alerta_preco = './IMAGES/alerta_preco.png'
    img_loc_sys_aco = './IMAGES/LOC_SYS_ACO.png'
    img_alerta_qtd = './IMAGES/alerta_qtd.png'
    img_alerta_desc = './IMAGES/alerta_desc.png'
    

    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    # fazer importacao dos dados da plan google
    sleep(0.3)
    loc_bt_inserir = pyautogui.locateOnScreen(img_loc_sys_aco, confidence=0.9)

    if loc_bt_inserir == None:

        status = 'Sistema Aco não localizado'
        cp(status)
        pyautogui.alert('SISTEMA NÃO LOCALIZADO')

    else:
        pyautogui.click(loc_bt_inserir)

        try:
            codigo = int(codigo)
        except:
            codigo = str(codigo)

        # digita codigo do item
        pyautogui.write(str(codigo))
        sleep(0.3)
        pyautogui.press('enter')
        sleep(0.2)
        # digita quantidade do item
        pyautogui.write(str(qtd))
        sleep(0.2)
        pyautogui.press('enter')
        sleep(0.3)
        alerta_qtd = pyautogui.locateOnScreen(img_alerta_qtd, confidence=0.8)
        sleep(0.5)
        if alerta_qtd != None:
            pyautogui.press('enter')
        sleep(0.5)
        # digita valor
        pyautogui.write(str(valor))
        sleep(0.2)
        pyautogui.press('enter')
        try:
            alerta_desc = pyautogui.locateOnScreen(img_alerta_desc, confidence=0.8)
            if alerta_desc != None:
                pyautogui.press('enter')
        except:
            pass  
        sleep(0.5)
        alertas = pyautogui.locateOnScreen(img_alerta_preco, confidence=0.8)
        sleep(0.5)

        if alertas !=None:
            pyautogui.press('enter')

        sleep(0.5)
        pyautogui.press('insert')
        sleep(0.5)
   
def novo_pedido(window, cliente, vendedor, ordem_de_compra):

    #img_tela = '.IMAGES/LOC_SYS_ACO'
    img_alert_vendedor = '.IMAGES/alerta_vendedor.png'
    img_ja_existe_pedido = './IMAGES/ja_existe_pedido.png'


    cp = sg.cprint
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    sleep(1)
    pyautogui.click(1961,419)
    sleep(0.2)
    pyautogui.press('enter', presses=2, interval=0.2)
    pyautogui.write(cliente)
    pyautogui.press('enter')
    
    sleep(1)

    pyautogui.keyDown('ctrlleft')
    sleep(0.2)
    pyautogui.press('enter', presses=1)
    pyautogui.keyUp('ctrlleft')
    sleep(0.2)
    pyautogui.press('enter')
    pyautogui.write(vendedor)
    pyautogui.press('enter')
    sleep(0.3)
    loc_alert_vendedor = pyautogui.locateAllOnScreen(img_alert_vendedor, confidence=0.8)
    
    if loc_alert_vendedor != None:
        pyautogui.press('S')
        sleep(0.2)
    pyautogui.press('enter', presses=4, interval=0.2)
    OC = 'OC ' + ordem_de_compra
    pyautogui.write(OC)
    pyautogui.press('enter', presses=4)
    pyautogui.write(ordem_de_compra)
    pyautogui.press('enter')

def finalizar_pedido(window):


    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    pyautogui.press('esc')
    pyautogui.press('p')
    sleep(0.2)
    pyautogui.press('enter')
    sleep(2)


if __name__ == '__main__':

    
    cp = sg.cprint
    size_btn = (50, 1)
    layout = [[sg.Push(), sg.Button('Executar', button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()]]
    layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(55, 20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
    layout += [[sg.Push(), sg.Cancel('Sair', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],
    window = sg.Window('Teste...', layout, finalize=True, keep_on_top=False,)

    while True:

        event, values = window.read()
        #cp(event, values)

        if event == sg.WIN_CLOSED or event == 'Sair':
            break

        elif event.startswith('Executar'):
            
        
            codigo = 'RJYC24LI'
            qtd = '10'
            valor = '10.15'
            #threading.Thread(target=novo_pedido, args=(window, cliente, vendedor, ordem_de_compra), daemon=True).start()
            threading.Thread(target=digitar_pedidos_ACO, args=(window,), daemon=True).start()