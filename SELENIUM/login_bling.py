from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep



class Login:
    def __init__(self):
        self.browser = Chrome()
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
        #//*[@id="quantidade_1"]
        self.xpath_input_valor_produto = '//*[@id="preco_unitario_'
        self.temp = '//*[@id="preco_unitario_0"]'
        self.xpath_novo_item = '//*[@id="add_new_item"]'
        self.xpath_num_oc = '//*[@id="numeroDaOrdemDeCompra"]'
        self.xpath_condicao = '//*[@id="pag_condicao"]'
        self.xpath_btn_salvar = '//*[@id="botaoSalvar"]'
        self.xpath_valor_total = '//*[@id="divTotal"]'


        #relatorios
        self.link_rel_estoque = 'https://www.bling.com.br/b/gerenciador.relatorio.php#view/82798'
   
    def iniciar(self):
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

    def fechar_browse(self):
        self.browser.close()

start = Login()   


start.iniciar()
sleep(5)
