from ast import Break
import pyautogui
import pandas as pd
from time import sleep


file = './arquivos/pedidos.xlsx'
img_alerta_preco = './IMAGES/alerta_preco.png'
img_loc_sys_aco = './IMAGES/LOC_SYS_ACO.png'
img_alerta_qtd = './IMAGES/alerta_qtd.png'


try:
    df = pd.read_excel(file, names=['CODIGO', 'QTD', 'PRECO'])
    lista_pedido = []
    lista_pedido = df.values.tolist()
except:
    print('Erro na leitura do excel')

def digitar_itens(lista):
    contador = 0
    total = 0
    print(f'ID\tCODIGO\t\tQTD\t VALOR\t\t TOTAL')
    for itens in lista:
        if contador >= 60:
            pyautogui.alert('Limite de itens maior que 60')
            contador = 0
        else:
            qtd = itens[1]
            valor = itens[2]
            subtotal = qtd*valor
            total = total + subtotal
            print(f'{contador + 1} \t{itens[0]} \t{str(itens[1])} \t{(itens[2]):>6.2f}\t\t{subtotal:>6.2f}')
            pyautogui.write(itens[0])
            sleep(0.3)
            pyautogui.press('enter')
            sleep(0.2)
            #digita quantidade do item
            pyautogui.write(str(itens[1]))
            sleep(0.2)
            pyautogui.press('enter')
            sleep(0.3)
            alerta_qtd = pyautogui.locateOnScreen(img_alerta_qtd, confidence=0.9)
            sleep(0.5)
            if alerta_qtd != None:
                pyautogui.press('enter')
                
            pyautogui.write(str(itens[2]))
            sleep(0.2)
            pyautogui.press('enter')
            sleep(0.5)
            alertas = pyautogui.locateOnScreen(img_alerta_preco, confidence=0.9)
            sleep(0.5)
               
            if alertas!=None:
                pyautogui.press('enter')
                    
            sleep(0.5)
            pyautogui.press('insert')
            sleep(0.5)
            contador += 1
               
    print(f'Valor total do pedido:    \t\t    {total:7.2f}')       


sleep(2)
loc_bt_inserir = pyautogui.locateOnScreen(img_loc_sys_aco, confidence=0.9)
print(f'Localização do Sistema ACO: {str(loc_bt_inserir)}')


if loc_bt_inserir == None:
    print('Sistema Aco não localizado')
    exit
else:
    pyautogui.click(loc_bt_inserir)
    digitar_itens(lista_pedido)
    pyautogui.press('esc')
