import pyautogui
import pandas as pd
from time import sleep


file = 'COMPRAS.xlsx'





try:
    df = pd.read_excel(file, names=['CODIGO', 'VALOR'])
    lista_pedido = []
    lista_pedido = df.values.tolist()
except: 
    print('erro na leitura do excel')


contador = 0
sleep(5)

for itens in lista_pedido:
    print(f'Item: {contador + 1}')
    sleep(1)
    pyautogui.write(itens[0])
    pyautogui.press('enter')
    print(itens[0])
    sleep(0.5)
    pyautogui.write(str(itens[1]))
    pyautogui.press('enter')
    pyautogui.press('backspace', presses=14)
     
    sleep(0.5)
    contador+=1
    
        

