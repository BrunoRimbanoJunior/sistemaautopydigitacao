import pyautogui
from time import sleep
import threading
import PySimpleGUI as sg




def novo_pedido(window, cliente, vendedor, ordem_de_compra):

    img_alert_vendedor = '.IMAGES/alerta_vendedor.png'
    cp = sg.cprint
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    #values = pyautogui.position()
    #cp(values)
    #definido uma posição fixa para o sistema
    pyautogui.click(1961,419)

    sleep(0.2)
    pyautogui.press('enter', presses=2, interval=0.2)
    pyautogui.write(cliente)
    sleep(0.3)
    pyautogui.press('enter', presses=2)
    pyautogui.write(vendedor)
    pyautogui.press('enter')
    sleep(0.3)
    loc_alert_vendedor = pyautogui.locateAllOnScreen(img_alert_vendedor, confidence=0.8)
    if loc_alert_vendedor != None:
        pyautogui.press('S')
        sleep(2)
    pyautogui.press('enter', presses=4, interval=0.2)
    OC = 'OC ' + ordem_de_compra
    pyautogui.write(OC)
    pyautogui.press('enter', presses=4)
    pyautogui.write(ordem_de_compra)
    pyautogui.press('enter')

    
    
   

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
            
            cliente = '04'
            vendedor = '04'
            ordem_de_compra = '123'
            threading.Thread(target=novo_pedido, args=(window, cliente, vendedor, ordem_de_compra), daemon=True).start()
                