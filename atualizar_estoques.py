from integracao_planilhas import atualizar_estoque
from time import sleep
import threading
from manipulacao_de_arquivos_locais import limpar_sistema, salvar_arq_estoque_pasta_local
import PySimpleGUI as sg


cp = sg.cprint


def atualiza_estoque_modulo(window):
    tarefa=''
    window.write_event_value('-THREAD-', (threading.current_thread().name, tarefa))
    cp('Iniciando processo de limpeza')

