import pyautogui
import pandas as pd
from time import sleep


file = 'tab_preco_centerparts.xlsx'
bt_inserir_produto = './IMAGES/inserir_produto.png'
campo_inserir = './IMAGES/campo_inserir_codigo.png'
campo_valor = './IMAGES/campo_valor.png'
bt_adicionar = './IMAGES/btn_adicionar.png'



try:
    df = pd.read_excel(file, names=['codigo', 'preco'])
    lista_preco = []
    lista_preco = df.values.tolist()
except: 
    print('erro na leitura do excel')




for itens in lista_preco:
    loc_bt_inserir = pyautogui.locateOnScreen(bt_inserir_produto)
    
    if loc_bt_inserir!=None:
        pyautogui.click(loc_bt_inserir)
        sleep(2)
        pyautogui.write(itens[0])
        print(itens[0])
        sleep(1)
        #pyautogui.press('enter')
        pyautogui.moveTo(988,265, duration=0.5)
        pyautogui.click()
        loc_bt_inserir = pyautogui.locateOnScreen(campo_valor)
        pyautogui.click(campo_valor)
        #pyautogui.moveTo(845,465, duration=0.5)
        #pyautogui.click()
        #sleep(2)
        #preco_temp = str(itens[1])
        pyautogui.write(str(itens[1]).replace('.',','))
        #sleep(2)
        loc_btn_adicionar = pyautogui.locateOnScreen(bt_adicionar)
        pyautogui.click(loc_btn_adicionar)
        sleep(2)
        




#for itens in lista_preco:
#    print(f'codigo:{itens[0]}, preco: {itens[1]}')