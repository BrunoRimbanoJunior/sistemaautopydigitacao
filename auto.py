import pyautogui
from time import sleep

desc_capa = 'CAPA RETROV'
desc_eletr = 'RETROV EXT'
compl_eletr = ' FIXO'
imagem = './IMAGES/desc_eletrica.png'


while True:

    #loc_descr_eletrica = pyautogui.locateOnScreen('./IMAGES/desc_eletrica.png')
    #loc_capa = pyautogui.locateOnScreen('./IMAGES/desc_capa.png')
    loc_imagem = pyautogui.locateOnScreen(imagem)
    
    if loc_imagem!=None:
        print('Aplicacao encontrada')
        pyautogui.click(imagem)
        sleep(1)
        pyautogui.moveTo(121,397, duration=0.1)
        pyautogui.dragTo(40,397, duration=0.5)
        pyautogui.write(desc_eletr)
        sleep(1)
        #pyautogui.moveTo(585,397, duration=0.1)
        #pyautogui.click()
        #pyautogui.write(compl_eletr)
        #sleep(1)
        pyautogui.moveTo(1088,303, duration=0.5)
        pyautogui.click()
        sleep(2)

    else:
        break