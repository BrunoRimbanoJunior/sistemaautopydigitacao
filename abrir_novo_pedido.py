from time import sleep
import pyautogui


def incluir_novo_pedido(cliente):
    bt_inserir_pedido = './IMAGES/inserir_pedido.png'


    loc_bt_inserir = pyautogui.locateOnScreen(bt_inserir_pedido, confidence=0.9)
    if loc_bt_inserir!=None:
        pyautogui.click(loc_bt_inserir)
        sleep(2)
        pyautogui.write(cliente)
        pyautogui.press('enter')
    else:
        print('Botao inserir pedido não encontrado')

