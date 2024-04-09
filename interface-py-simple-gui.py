import PySimpleGUI as sg

# Layout
sg.theme('Dark Blue 3')

layout = [[sg.Text('Caminho da planilha: ', size=20), sg.InputText('', key='caminho'),
sg.FileBrowse("Procurar", file_types=(("Excel", "*.xlsx"),))],
[sg.Text('Linha inicial da planilha: ', size=20), sg.InputText("1", key='linhaini', size=4)],
[sg.Text('Linha final da planilha: ', size=20), sg.InputText("1", key='linhafim', size=4)],
[sg.Text('Meses a buscar: ', size=20), sg.DropDown(["1", "2", "3", "4", "5", "6"], "3", key='periodo')],
[sg.Text('Solicitar (1) / Baixar (2): ', size=20), sg.DropDown(["1", "2"], "1", key='finalidade')],
[sg.Text('Salvar arquivos: ', size=20), sg.InputText('', key='salvar'), sg.FolderBrowse('Procurar')],
[sg.Text('Data inicial (DD/MM/AAAA): ', size=20), sg.InputText('', size=10, key='dataini')],
[sg.Checkbox('Certificado Próprio', key='proprio')],
[sg.Text('1. Se selecionada a opção "Certificado Próprio" a planilha não precisa ser informada')],
[sg.Button('Iniciar'), sg.Button('Sair')]]

# Janela
janela = sg.Window('Download eSocial', layout)

while True:
    eventos, valores = janela.read()                

    if eventos == sg.WIN_CLOSED or eventos == 'Sair':
        break
    if eventos == 'Iniciar':
        janela.close()