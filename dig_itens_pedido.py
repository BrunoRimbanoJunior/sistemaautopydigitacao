import pyautogui
import pandas as pd
from time import sleep



def digitar_itens_bling(file):
    try:
        df = pd.read_excel(file, names=['codigo', 'qtd', 'valor'])
        lista_pedido = []
        lista_pedido = df.values.tolist()
    except: 
        print('erro na leitura do excel')


    sleep(2)
    bt_inserir_produto = './IMAGES/inserir_produtos.png'
    loc_bt_inserir = pyautogui.locateOnScreen(bt_inserir_produto, confidence=0.9)
    #print(f'Localização do Sistema: {str(loc_bt_inserir)}')
    if loc_bt_inserir == None:
        print('Sistema Bling não localizado')
        exit
    else:
        pyautogui.click(loc_bt_inserir)
        digitacao(lista_pedido)
        pyautogui.press('esc')

def digitacao(lista):
    bt_inserir_produto = './IMAGES/inserir_produtos.png'
    sleep(5)
    contador = 0
    loc_bt_inserir = pyautogui.locateOnScreen(bt_inserir_produto, confidence=0.9)
    pyautogui.click(loc_bt_inserir)
    total = 0

    if loc_bt_inserir == None:
        print('Sistema de pedidos não localizao')
        exit
    else:    
        print(f'ID\tCODIGO\t\tQTD\t VALOR\t\t TOTAL')
        for itens in lista:
            
            
            codigo = itens[0]
            qtd = float(itens[1])
            valor = float(itens[2])
            subtotal = qtd * valor
            total = total + subtotal

            print(f'{contador + 1:<3}\t{codigo:<14} \t{qtd:>3} \t{valor:>6.2f}\t\t{subtotal:>6.2f} ')
            pyautogui.write(str(itens[0]))
            #print(f'{contador + 1}\t{itens[0]} \t{str(itens[1])} \t{str(itens[2])} ')
            sleep(0.5)
            pyautogui.press('enter')
            sleep(1.5)
            pyautogui.write(str(itens[1]))
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab') 

            pyautogui.write(str(itens[2]).replace('.',','))
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter') 
            sleep(1)
         
        
             
        sleep(2)
        
        contador+=1
        loc_bt_inserir = ''
        
    print(f'\nValor total do pedido.......................: {total:8.2f}\n')

if __name__ == '__main__':
    bt_inserir_produto = './IMAGES/inserir_produtos.png'
    loc_bt_inserir = pyautogui.locateOnScreen(bt_inserir_produto, confidence=0.9)
    if loc_bt_inserir == None:
        print('Sistema Aco não localizado')
        exit
    else:
        pyautogui.click(loc_bt_inserir)
        file = './arquivos/pedidos.xlsx'
        digitar_itens_bling(file)
        pyautogui.press('esc')
