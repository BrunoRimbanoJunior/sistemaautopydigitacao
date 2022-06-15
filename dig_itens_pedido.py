import pyautogui
import pandas as pd
from time import sleep


file = 'pedidos.xlsx'
bt_inserir_produto = './IMAGES/inserir_produtos.png'




try:
    df = pd.read_excel(file, names=['codigo', 'qtd', 'valor'])
    lista_pedido = []
    lista_pedido = df.values.tolist()
except: 
    print('erro na leitura do excel')

sleep(2)
contador = 0
loc_bt_inserir = pyautogui.locateOnScreen(bt_inserir_produto)
pyautogui.click(loc_bt_inserir)

for itens in lista_pedido:
    print(f'Item: {contador + 1}')
    sleep(1)
    pyautogui.write(str(itens[0]))
    print(f'Codigo: {itens[0]} \tQtd: {str(itens[1])} \t VALOR: {str(itens[2])} ')
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
    sleep(2)
    contador+=1
    loc_bt_inserir = ''
        

