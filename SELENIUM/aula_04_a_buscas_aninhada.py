from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By




def find_by_text(browser, tag, text):
    elementos = browser.find_elements(By.TAG_NAME,tag)

    for elemento in elementos:
        if elemento.text == text:
            return elemento

def fin_by_href(browser, tag, link):
    elementos = browser.find_elements(By.TAG_NAME, tag)

    for elemento in elementos:
        if link in elemento.get_attribute('href'):
            return elemento





browser = Chrome()
browser.get('http://selenium.dunossauro.live/aula_04_a.html')

'''
tag_lista = browser.find_elements(By.TAG_NAME, 'li')
print(f'QUantidade de itens com TAG li: {len(tag_lista)}')

lista_nao_ordena = browser.find_elements(By.TAG_NAME, 'ul')
print(f'QUantidade de itens com TAG ul: {len(lista_nao_ordena)}')
print('Lista 1 :')
#print(lista_nao_ordena[0].text)
lista = lista_nao_ordena[0].find_element(By.TAG_NAME, 'a').text
print(lista)
'''
elemento_ddg = find_by_text(browser, 'li', 'DuckDuckGo')
print(elemento_ddg.text)
link_ddg = fin_by_href(browser, 'a', 'ddg')
print(link_ddg.get_attribute('href'))

browser.quit()