import pandas as pd
from time import sleep
from login_bling import start
import PySimpleGUI as sg
import threading



def digitar_pedidos_online(window, file):
    cp = sg.cprint
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    status = 'Inciando processo de digitação de pedidos'
    cp(status)


    start.iniciar()
    sleep(5)
    cp(f'Arquivo excel selecionado: {file}')
    lista_pedidos = []
    lista_pedidos2 = []

    try:
        df = pd.read_excel(file, names=['CLIENTE', 'CODIGO','SEPARACAO','PRECO', 'PED','STATUS'])
        lista_pedidos = df.values.tolist()
        lista_pedidos2 = df.values.tolist()
        cp('IMPORTAÇÃO EXCEL COM SUCESSO\n')
    except: 
        cp('********Erro na leitura do excel******')

    list_cliente = []
    list_oc = []

    for cliente in lista_pedidos:
        if str(cliente[0]) in list_cliente:
            pass
        else:
            #da lista principal, cria uma lista com o nome dos clientes
            list_cliente.append(str(cliente[0]))

        if str(cliente[4]) in list_oc:
            pass
        else:
            #da lista pincipal cria uma lista com o numero dos pedidos
            list_oc.append(str(cliente[4]))        



    oc_ant = ''
    oc = ''
    cont = 1


    #passar por todos os clientes da lista
    for cliente in list_cliente:
        #novo cliente novo pedido
        cp(f'CLIENTE: {cliente}\n')
        cp(f'ID   Codigo        \t\tQTD\t\tVALOR  \t\t  ORDEM DE COMPRA')
        sleep(5)
        start.novo_pedido(cliente)
        cont=1
        total = 0
        subtotal = 0

        #passar pela lista principal inteira
        for item in lista_pedidos2:
            #se o cliente for igual ao registro cliente...
            if str(item[0])== cliente:
                #grava o numero da oc atual
                oc = str(item[4])
                #compara o numero das ordens, se for igual ou for o primeiro item imprime
                if oc == oc_ant or cont == 1:
                #login_bling.start.iniciar()
                    sleep(2)
                    qtd = float(item[2])
                    valor = float(item[3])
                    subtotal = qtd * valor
                    total = total + subtotal
                    start.digitar_itens(cont-1, str(item[1]), str(item[2]), str(item[3]))
                    sleep(1)
                    cp(f'{str(cont)}\t{str(item[1]):14} \t\t{str(item[2]):>3} \t\t{str(item[3]):>7} \t\t{str(item[4]):10}')
                    oc_ant = str(item[4]) 
                    #print(f'Caso Igual: {oc} - {oc_ant}')
                    cont+=1
                else:
                    cp(f'Valor total.......................{total:5.2f}')
                    start.digitar_info_ped(str(oc_ant))
                    sleep(2)
                    start.salvar_pedido()
                    sleep(2)
                    #login_bling.start.fechar_browse()
                    #sleep(5)

                    cp('*************NOVA ORDEM DE COMPRA***************')
                    total = 0
                    subtotal = 0
                    cont=1
                    #login_bling.start.iniciar()
                    #sleep(5)
                    start.novo_pedido(cliente)
                    sleep(5)
                    qtd = float(item[2])
                    valor = float(item[3])
                    subtotal = qtd * valor
                    total = total + subtotal
                    start.digitar_itens(cont-1, str(item[1]), str(item[2]), str(item[3]))
                    sleep(1)
                    cp(f'{str(cont)}\t{str(item[1]):14} \t\t{str(item[2]):>3} \t\t{str(item[3]):>7} \t\t{str(item[4]):10}')
                    oc_ant = str(item[4]) 
                    cont+=1

        cp(f'Valor total.......................{total:5.2f}')
        cp('\n')
        start.digitar_info_ped(str(oc_ant))
        sleep(2)  
        start.salvar_pedido()
        sleep(2)
        #start.fechar_browse()
        #sleep(5)

   # start.digitar_itens()
    start.fechar_browse()
