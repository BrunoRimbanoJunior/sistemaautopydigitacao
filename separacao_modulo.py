from integracao_planilhas import integracao_estoque, integracao_matriz, inserir_dados
import threading
import PySimpleGUI as sg





def buscar_dados_estoque():
    
    planilha = integracao_estoque()
    lista_estoque = []
    
    for row in planilha:
        lista_estoque.append([row[0], row[3]])

    status = 'Carregando dados de estoque......'
    print(status)
    return lista_estoque

def salvar_carteira():
    

    planilha = integracao_matriz()
    carteira = []

    for row in planilha:
        if int(row[8])>0:
            carteira.append([row[0], row[1], row[2], row[3], row[8], row[5]])

    print('Carregando dados da carteira.....')
    return carteira

def busca_estoque(lista, elem):
    #print('Buscando dados do estoque pelo codigo do produto......')
    for posicao, valor in enumerate(lista):
        if valor[0] == elem:
            return valor[1]
           
    raise Exception('Codigo de produto na ta bela matriz nao existe da tabela saldo')      

def busca_index(lista, elem):
    #print('Buscando os valores de indice.....')
    for posicao, valor in enumerate(lista):
        if valor[0] == elem:
            return posicao
           
    raise Exception('Valor não encontrado')        

def menor_valor(valor1, valor2):
    if int(valor1)<int(valor2):
        return int(valor1)
    elif int(valor1)>int(valor2):
        return int(valor2)
    else:
        return valor2

def fazer_separacao(window):
    cp = sg.cprint
    status = ''
    window.write_event_value('-THREAD-', (threading.current_thread().name, status))
    status = 'Inciando processo de separacao no modulo....'
    cp(status)
    
    cp('Atualizando Estoques.........')
    estoque = buscar_dados_estoque()
    cp('Salvando Carteira............')
    carteira = salvar_carteira()
    list_sep = []


    for linha in carteira:
        #print(f'{linha[0]}, {linha[1]}, {linha[2]}, {linha[3]}, {linha[4]}, {linha[5]}')    
        #linha[0]= nome, linha[1] = id, linha[2] = oc, linha[3] = codigo, linha[4] = saldo, linha[5]= valor
        #print(f'Buscando item: {linha[3]}')
        qtd_est = int(busca_estoque(estoque, linha[3]))
        sep = int(menor_valor(linha[4], qtd_est))
        ind = busca_index(estoque, linha[3])
        #print(f'{linha[3]}', qtd_est, ind)
        if sep>0:
            print(f'{linha[0]}, {linha[1]}, {linha[2]}, {linha[3]}, {linha[4]}, {linha[5]}, Separacao: {sep}')
            estoque[ind][1] = int(estoque[ind][1])-sep
            list_sep.append([linha[0], linha[1], linha[2], linha[3], sep, linha[5]])
        #print(f'Saldo: {estoque[ind][1]}')

    cp('Inserindo dados na Tabela Separacao')        
    inserir_dados(list_sep)
    cp('Processo Finalizado')    
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
        elif event == 'Executar':
            fazer_separacao()   

   