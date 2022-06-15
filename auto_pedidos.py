import pyautogui
import pandas as pd
from time import sleep
import abrir_novo_pedido


file = 'pedidos_clientes.xlsx'



try:
    df = pd.read_excel(file, names=['CLIENTE', 'CODIGO','SEPARACAO','PRECO', 'PED'])
    lista_pedidos = []
    lista_pedidos = df.values.tolist()
    lista_pedidos2 = df.values.tolist()
    print('IMPORTAÇÃO EXCEL COM SUCESSO\n')
except: 
    print('********Erro na leitura do excel******')

list_cliente = []
list_oc = []
for itens in lista_pedidos:
    if str(itens[0]) in list_cliente:
        pass
    else:
        list_cliente.append(str(itens[0]))

    if str(itens[4]) in list_oc:
        pass
    else:
        list_oc.append(str(itens[4]))        



oc_ant = ''
oc = ''
cont = 1
for cliente in list_cliente:
    print(f'CLIENTE: {cliente}\n')
    abrir_novo_pedido.incluir_novo_pedido(cliente)
    sleep(4)
    #print(f'Codigo\t\t\tQTD\t\tVALOR\t\tORDEM DE COMPRA')
    for item in lista_pedidos2:
        
        if str(item[0])== cliente:
            oc = str(item[4])
            if oc == oc_ant or cont == 1:
                print(f'{str(item[1]):14} \t\t{str(item[2]):>3} \t\t{str(item[3]):>5} \t\t{str(item[4])}')
                oc_ant = str(item[4]) 
                #print(f'Caso Igual: {oc} - {oc_ant}')
                cont+=1
            else:    
                print('*********NOVA OC*********')
                print(f'Codigo\t\t\tQTD\t\tVALOR\t\tORDEM DE COMPRA')
                oc_ant = str(item[4]) 
                #print(f'Caso Diferente: {oc} - {oc_ant}')
                print(f'{str(item[1]):14} \t\t{str(item[2]):>3} \t\t{str(item[3]):>5} \t\t{str(item[4])}')  
                cont=+1
    print('\n')  



