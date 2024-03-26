import PySimpleGUI as sg

class SistemaValidacao:
    def __init__(self, chave_valida):
        self.chave_valida = chave_valida

    def validar_chave(self, chave):
        if chave == self.chave_valida:
            return True
        else:
            return False

# Chave válida predefinida
chave_valida = "chave123"

# Instanciar o sistema de validação
sistema_validacao = SistemaValidacao(chave_valida)

# Layout da janela
layout = [
    [sg.Text('Digite a chave de acesso:')],
    [sg.InputText(key='-CHAVE-')],
    [sg.Button('Validar')]
]

# Criar a janela
window = sg.Window('Validação de Chave').Layout(layout)

# Loop de eventos
while True:
    event, values = window.Read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'Validar':
        chave_usuario = values['-CHAVE-']
        if sistema_validacao.validar_chave(chave_usuario):
            sg.popup("Chave válida. Acesso concedido.")
            # Coloque aqui o código que deseja executar após a validação bem-sucedida
        else:
            sg.popup("Chave inválida. Acesso negado.")

# Fechar a janela
window.close()
