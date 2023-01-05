from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
import pandas as pd
import numpy as np
import PySimpleGUI as sg




SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def integracao_matriz():
    SAMPLE_SPREADSHEET_ID = '1NHsGMo8UpNEtZLg2WSBwV9xiKNLJeeMOF_0O7f8mUDU'
    SAMPLE_RANGE_NAME = 'MATRIZ!A9:R'
    creds = None
   
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
       
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        #SALVANDO OS DADOS LOCALMENTE                            
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return
        return values
    
    except HttpError as err:
        print(err)

def integracao_separacao():
    SAMPLE_SPREADSHEET_ID = '1NHsGMo8UpNEtZLg2WSBwV9xiKNLJeeMOF_0O7f8mUDU'
    SAMPLE_RANGE_NAME = 'SEPARACAO!A2:M'
    creds = None
   
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
       
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        #salvando os valores localmente                          
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        
        return values
    
    except HttpError as err:
        print(err)

def integracao_estoque():
    SAMPLE_SPREADSHEET_ID = '1NHsGMo8UpNEtZLg2WSBwV9xiKNLJeeMOF_0O7f8mUDU'
    SAMPLE_RANGE_NAME = 'SOMA_SEP!A2:D'
    creds = None
   
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
       
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        #salvando os valores localmente                          
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return
        return values
    
    except HttpError as err:
        print(err)

def inserir_dados(lista_sep):
    SAMPLE_SPREADSHEET_ID = '1NHsGMo8UpNEtZLg2WSBwV9xiKNLJeeMOF_0O7f8mUDU'
    SAMPLE_RANGE_NAME = 'SEPARACAO!A2:L'
    creds = None
   
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
       
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        values = integracao_separacao()
        #print(f'Quantidade de linhas: {len(values)+1}')
        new_range = 'SEPARACAO!A'+str(len(values)+2)
        range = len(values)+2
        #print(new_range)
        data = datetime.now().strftime('%d/%m/%Y')
        dados_sep = []
        #linha[0]= nome, linha[1] = id, linha[2] = oc, linha[3] = codigo, linha[4] = saldo, linha[5]= valor
        #seguencia 
        # 0 = DATA / 
        # 1 = FOMULA INDICE / 
        # 2 = ID CLIENTE / 
        # 3 = CODIGO / 
        # 4 = QTD / 
        # 5 = PRECO / 
        # 6 = TOTAL / 
        # 7 = NOME CLINETE / 
        # 8 = OC / 
        # 9 = SALDO CLIENTE / 
        # 10 = NOTA / 
        # 11 = STATUS 
        #-------------------------------------------------------------------------------------------------------
        #lista recebida de fora
        for row in lista_sep:
            formula1 = str('=CONCATENAR(C'+str(range)+';I'+str(range)+';D'+str(range)+')')
            formula2 = str('=e'+str(range)+'*f'+str(range))
            formula3 = str('=PROCV(B'+str(range)+';MATRIZ!$L$9:R;7;FALSO)')
            dados_sep.append([
                data,
             formula1,
             row[1],
             row[3],
             row[4],
             row[5],
             formula2,
             row[0],
             row[2],
             formula3,
             '',
             '0',
             ])
            range+=1
        
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=new_range, 
                                    valueInputOption='USER_ENTERED',
                                    body={'values': dados_sep}).execute()
        
    except HttpError as err:
        print(err)

def atualizar_estoque():
    cp = sg.cprint
        
    cp('Iniciando processo de Atualizaçao dos Estoque dentro do Modulo')



    SAMPLE_SPREADSHEET_ID = '1NHsGMo8UpNEtZLg2WSBwV9xiKNLJeeMOF_0O7f8mUDU'
    SAMPLE_RANGE_NAME = 'ESTOQUE!A2:B'
    creds = None
   
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
       
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
       
        planilha = integracao_estoque()
        sep_ant = []
        lista_vazia = []
        for row in planilha:
            sep_ant.append([row[0], row[1]])
        for i in range(0,5977):
            lista_vazia.append(['', ''])  
        
        #print(f'Quantidade de linhas: {len(planilha)+1}')
        new_range = 'ESTOQUE!A2'
                            
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range='ESTOQUE!A2:B5978', 
                                    valueInputOption='USER_ENTERED',
                                    body={'values': lista_vazia}).execute()
        
                     
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=new_range, 
                                    valueInputOption='USER_ENTERED',
                                    body={'values': sep_ant}).execute()
        
        try:
            df = pd.read_csv('./ESTOQUE/estoque.csv', sep=';', index_col=None)
        except:
            cp('Arquivo estoque.csv não localizado')   
             
        df[df.columns[0]].replace('', np.nan, inplace=True)
        df.dropna(subset=[df.columns[0]], inplace=True)
        estoque_online = df.values.tolist()
        list_estoque = []
        #print(estoque_online)
        for linha in estoque_online:
            list_estoque.append([linha[0], linha[2]])

        consulta = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute() 
        values = consulta.get('values', [])                               
          
        new_range = 'ESTOQUE!A'+str(len(values)+2)
        
        insercao = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=new_range, 
                                    valueInputOption='USER_ENTERED',
                                    body={'values': list_estoque}).execute()

        cp('Tabelas dentro do Modulo Atualizadas...')    
        
    except HttpError as err:
        print(err)
    
def integracao_pedidos():
    SAMPLE_SPREADSHEET_ID = '1NHsGMo8UpNEtZLg2WSBwV9xiKNLJeeMOF_0O7f8mUDU'
    SAMPLE_RANGE_NAME = 'PEDIDOS!A1:C'
    creds = None
   
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
       
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        #salvando os valores localmente                          
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return
        return values
    
    except HttpError as err:
        print(err)

if __name__ == '__main__':
    cp = sg.cprint
    size_btn = (50,1)
    layout = [[sg.Push(), sg.Button('Executar', button_color=('yellow', 'blue'), size=size_btn, font='bold'), sg.Push()]]
    layout += [[sg.Push(), sg.Multiline(key='-PRINT-', size=(55,20), autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True), sg.Push()]]
    layout += [[sg.Push(), sg.Cancel('Sair', button_color=('black', 'red'), size=size_btn, font='bold'), sg.Push()]],
    window = sg.Window('Teste...', layout, finalize=True, keep_on_top=False,)
               
    while True:
                
        event, values = window.read()
        #cp(event, values)


        if event == sg.WIN_CLOSED or event == 'Sair':
            break
        elif event.startswith('Executar'):
            integracao_separacao()