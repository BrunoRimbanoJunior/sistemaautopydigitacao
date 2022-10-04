from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from manipulacao_de_arquivos_locais import salvar_arq_estoque_pasta_local
from integracao_planilhas import atualizar_estoque
import threading
import PySimpleGUI as sg



class Login:
    def __init__(self):
        
       

        self.browser = ''
        self.chrome_options = Options()
        self.chrome_options.add_argument('--disable-extensions')
        self.chrome_options.add_argument('--disable-gpu')
        self.chrome_options.add_argument('--no-sandbox')
        self.chrome_options.add_argument('--headless')
        self.prefs = {'download.default_directory' : r'.\ESTOQUE'}
        self.chrome_options.add_experimental_option('prefs', self.prefs)

        #dados de usuario
        self.user = 'juba.rimbano@gmail.com'
        self.password = '*Fb264e0d9efg'
        self.link1 = 'https://www.bling.com.br/login'
        self.link_pedidos = 'https://www.bling.com.br/b/vendas.php#list'
        self.xpath_user = '//*[@id="username"]'
        self.xpath_pass = '//*[@id="senha"]'
        self.xpath_btn_login = '//*[@id="login-buttons-site"]/button'

        
        #pedidos
        self.xpath_btn_novo_pedido = '//*[@id="btn-incluir"]/span'
        self.xpath_input_nome = '//*[@id="contato"]'
        self.xpath_input_cod_produto = '//*[@id="produto_descricao"]'
        self.xpath_input_qtd_produto = '//*[@id="quantidade_'
        self.xpath_input_valor_produto = '//*[@id="preco_unitario_'
        self.temp = '//*[@id="preco_unitario_0"]'
        self.xpath_novo_item = '//*[@id="add_new_item"]'
        self.xpath_num_oc = '//*[@id="numeroDaOrdemDeCompra"]'
        self.xpath_condicao = '//*[@id="pag_condicao"]'
        self.xpath_btn_salvar = '//*[@id="botaoSalvar"]'
        self.xpath_valor_total = '//*[@id="divTotal"]'


        #relatorios
        self.xpath_link_rel_estoque = 'https://www.bling.com.br/b/gerenciador.relatorio.php#view/82798'
        self.xpath_btn_relatorio = '//*[@id="exportRelatorioLnk"]'
        self.xpath_btn_salvar_relatorio = '/html/body/div[11]/div[3]/div/button[1]'
    
    def iniciar(self):
        self.browser = Chrome()
        self.browser.get(self.link1)
        print(f'Conectando ha: {self.link1}')
        sleep(4)
        self.browser.find_element(By.XPATH, self.xpath_user).send_keys(self.user)
        self.browser.find_element(By.XPATH, self.xpath_pass).send_keys(self.password)
        self.browser.find_element(By.XPATH, self.xpath_btn_login).click()
        sleep(4)
        self.browser.get(self.link_pedidos)

    def novo_pedido(self, nome):
        self.browser.find_element(By.XPATH, self.xpath_btn_novo_pedido).click()
        sleep(3)
        self.browser.find_element(By.XPATH, self.xpath_input_nome).send_keys(nome)
        self.browser.find_element(By.XPATH, self.xpath_input_nome).send_keys(Keys.ENTER)

    def digitar_itens(self, contador, codigo, qtd, valor):
        self.browser.find_element(By.XPATH, self.xpath_input_cod_produto).send_keys(codigo)
        self.browser.find_element(By.XPATH, self.xpath_input_cod_produto).send_keys(Keys.ENTER)
        sleep(1)
        xpath_qtd = str(self.xpath_input_qtd_produto)+str(contador)+'"]'
        self.browser.find_element(By.XPATH, xpath_qtd).send_keys(str(qtd))
        xpath_valor = str(self.xpath_input_valor_produto)+str(contador)+'"]'
        self.browser.find_element(By.XPATH, xpath_valor).click()
        sleep(0.5)
        self.browser.find_element(By.XPATH, xpath_valor).send_keys(str(valor).replace('.',','))
        sleep(0.3)
        self.browser.find_element(By.XPATH, self.xpath_novo_item).click()

    def digitar_info_ped(self, oc):
        self.browser.find_element(By.XPATH, self.xpath_num_oc).send_keys(str(oc))
  
    def salvar_pedido(self):
        self.browser.find_element(By.XPATH, self.xpath_btn_salvar).click()

    def validor_pedido(self):
        valor = self.browser.find_element(By.XPATH, self.xpath_valor_total).text
        print(valor)

    def salvar_estoque_bling(self, window):
        window.write_event_value('-THREAD-', (threading.current_thread().name))
        cp = sg.cprint
        cp('Abrindo navegador....')
        self.browser = Chrome()
        self.browser.get(self.link1)
        sleep(4)
        cp('Inserindo dados de usuario')
        self.browser.find_element(By.XPATH, self.xpath_user).send_keys(self.user)
        self.browser.find_element(By.XPATH, self.xpath_pass).send_keys(self.password)
        self.browser.find_element(By.XPATH, self.xpath_btn_login).click()
        sleep(4)
        cp('Acessando relatorios....')
        self.browser.get(self.xpath_link_rel_estoque)
        sleep(15)
        self.browser.find_element(By.XPATH, self.xpath_btn_relatorio).click()
        sleep(5)
        self.browser.find_element(By.XPATH, self.xpath_btn_salvar_relatorio).click()
        cp('Relatorio exportando....')
        sleep(2)
        cp('Salvando relatorio na pasta ESTOQUE...')
        salvar_arq_estoque_pasta_local(window)
        sleep(2)
        cp('Finalizando processo....')
        self.browser.close()

    def fechar_browse(self):
        self.browser.close()

start = Login()   

def auto_estoque(window):
    start.salvar_estoque_bling(window)
    sleep(2)
    atualizar_estoque()

if __name__ == '__main__':
    size_btn = (50,1)
    layout = [[sg.Push(), sg.Button('Executar', button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()]]
    layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(55,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
    layout += [[sg.Push(), sg.Cancel('Sair', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],
    window = sg.Window('Limpeza de Sistema', layout, finalize=True, keep_on_top=False,)
               
    while True:
                
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event.startswith('Executar'): 
            threading.Thread(target=start.salvar_estoque_bling, args=(window,), daemon=True).start()   
            #start.salvar_estoque_bling(window)
    
    