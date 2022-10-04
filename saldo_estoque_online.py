from integracao_planilhas import integracao_estoque
import pandas as pd




def buscar_dados_estoque():

    planilha = integracao_estoque()
    lista = []
    file = './ESTOQUE/ESTOQUE_RETROGLASS.xlsx'
    for row in planilha:
        lista.append([row[0], row[3]])

    

    df = pd.DataFrame(lista, columns= ['CODIGO', 'SALDO'])
    df['SALDO'] = pd.to_numeric(df['SALDO'], errors='coerce')
    

    df.to_excel(file, index=False)


        
if __name__ == '__main__':
    buscar_dados_estoque()    

 