from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep




def find_by_text(browser, tag, text):
    elementos = browser.find_elements(By.TAG_NAME,tag)

    for elemento in elementos:
        if elemento.text == text:
            return elemento

browser = Chrome()
browser.get('http://selenium.dunossauro.live/aula_04_b.html')

find_by_text(browser, 'div', 'um').click()
sleep(1)
find_by_text(browser, 'div', 'dois').click()
sleep(1)
find_by_text(browser, 'div', 'tres').click()
sleep(1)
find_by_text(browser, 'div', 'quatro').click()
sleep(1)
browser.back()
sleep(1)
browser.back()
sleep(1)
browser.back()
sleep(1)
browser.back()
sleep(1)

