import threading
from time import sleep
import PySimpleGUI as sg


THREAD_EVENT = '-THREAD-'
cp = sg.cprint

def the_thread(window):

    i = 0
    while True:
        sleep(1)
        window.write_event_value('-THREAD-', (threading.current_thread().name, i)) 
        cp('This is cheating from the thread', c='white on green')
        i += 1

def frases(window):
    
    tarefa = ''
    tarefas = ['Abrindo Sistema', 'Operando', 'Finalizando Sistema']
    for tarefa in tarefas:
        sleep(1)
        window.write_event_value('-THREAD-', (threading.current_thread().name, tarefa))      # Data sent is a tuple of thread name and counter
        cp(tarefa, c='white on green')         

def teste_2(window):

    dado = ''
    dados = ['Teste de Funcao', 'Teste 2', 'Teste 3']
    
    for dado in dados:
        sleep(1)
        window.write_event_value('-THREAD-', (threading.current_thread().name, dado))
        cp(dado, c='red on white')
        


def main():
    

    layout = [  [sg.Text('Output Area - cprint\'s route to here', font='Any 15')],
                [sg.Multiline(size=(65,20), key='-ML-', autoscroll=True, reroute_stdout=True, write_only=True, reroute_cprint=True)],
                [sg.T('Input so you can see data in your dictionary')],
                [sg.Input(key='-IN-', size=(30,1))],
                [sg.B('Start A Thread'), sg.B('Teste'), sg.B('Dummy'), sg.Button('Exit')]  ]

    window = sg.Window('Window Title', layout)

    while True:             # Event Loop
        event, values = window.read()
        #cp(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        elif event.startswith('Start'):
            threading.Thread(target=frases, args=(window,), daemon=True).start()
        elif event.startswith('Teste'):
            threading.Thread(target=teste_2, args=(window,), daemon=True).start()
           
           
            
    window.close()


if __name__ == '__main__':
    main()