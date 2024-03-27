import time
from datetime import datetime, timedelta
import PySimpleGUI as sg
import openpyxl
import undetected_chromedriver as uc
from auto_download_undetected_chromedriver import download_undetected_chromedriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from undetected_chromedriver import ChromeOptions
import os
import shutil
import requests
import psutil

""" def fechar_processos_chrome():
    for proc in psutil.process_iter():
        # Verifica se o processo pertence ao Chrome
        if "chrome" in proc.name():
            try:
                proc.terminate()  # Encerra o processo
            except psutil.AccessDenied:
                # Se houver permissões insuficientes para encerrar o processo
                print(f"Permissões insuficientes para encerrar o processo {proc.pid}") """

class SistemaValidacao:
    def __init__(self, chave_valida):
        self.chave_valida = chave_valida

    def validar_chave(self, chave):
        if chave == self.chave_valida:
            return True
        else:
            return False

# Chave válida predefinida
chave_valida = "S@nta799"

# Instanciar o sistema de validação
sistema_validacao = SistemaValidacao(chave_valida)

# Layout da janela de login
sg.theme('Dark Blue 3')
layout_login = [
    [sg.Text('Digite a chave de acesso:')],
    [sg.InputText(password_char='*', key='-CHAVE-')],
    [sg.Button('Validar')]
]

# Criar a janela de login
window_login = sg.Window('Login').Layout(layout_login)

# Loop de eventos da janela de login
while True:
    event, values = window_login.Read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'Validar':
        chave_usuario = values['-CHAVE-']
        if sistema_validacao.validar_chave(chave_usuario):
            sg.popup("Chave válida. Acesso concedido.")
            # Fechar a janela de login
            window_login.close()


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
            [sg.Button('Iniciar'), sg.Button('Sair')]]

            # Janela
            janela = sg.Window('Download eSocial', layout)

            while True:
                eventos, valores = janela.read()

                if eventos == sg.WIN_CLOSED or eventos == 'Sair':
                    break
                if eventos == 'Iniciar':

                    download_dir = valores['salvar'].rstrip().replace('/', '\\')

                    def criar_pasta(nome_empresa):
                        pasta_empresa = os.path.join(download_dir, nome_empresa)
                        if not os.path.exists(pasta_empresa):
                            os.makedirs(pasta_empresa)
                        return pasta_empresa

                    url = 'https://login.esocial.gov.br/login.aspx'
                    folder_path = "c:\\chromedriver"

                    workbook = openpyxl.load_workbook(valores['caminho'])
                    sheet_empresas = workbook.active
                    documento = sheet_empresas.cell(row=int(valores['linhaini']), column=3).value

                    # Condição para solicitar as datas quando o certificado é por procuração de PJ/PF
                    if valores['finalidade'] == '1' and valores['proprio'] is False:
                        chrome_options = ChromeOptions()
                        chromedriver_path = download_undetected_chromedriver(folder_path, undetected=True, arm=False,
                                                                            force_update=True)
                        driver = uc.Chrome(driver_executable_path=chromedriver_path, options=chrome_options, headless=False)            
                        driver.get(url)
                        driver.maximize_window()
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-acoes"]/div[2]/p/button'))
                        )
                        driver.find_element('xpath', '//*[@id="login-acoes"]/div[2]/p/button').click()

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-certificate"]'))
                        )
                        driver.find_element('xpath', '//*[@id="login-certificate"]').click()

                        print("Selecione o certificado para continuar")

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="sairAplicacao"]'))
                        )
                        
                        inscricao = driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/p[2]/span[1]').text.strip('-')
                        print(f'CNPJ do procurador: {inscricao}')
                        # Condição para identificar se a inscrição é um CNPJ ou CPF
                        if len(driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/p[2]/span[1]').text) < 18:
                            driver.find_element(By.XPATH, '//*[@id="geral"]/div').click()
                            WebDriverWait(driver, 120).until(
                                EC.presence_of_element_located((By.XPATH, '//*[@id="header"]/div[2]/a'))
                            )
                            driver.find_element('xpath', '//*[@id="header"]/div[2]/a').click()
                            pass
                        else:
                            driver.find_element(By.CLASS_NAME, 'alterar-perfil').click()

                        for linha in sheet_empresas.iter_rows(min_row=int(valores['linhaini']), max_row=int(valores['linhafim'])):
                            documento = len(str(linha[2].value))
                            print(f'CNPJ/CPF sendo buscado: {str(linha[2].value)}')
                            time.sleep(1)
                            # Condição que verifica se é CNPJ ou CPF na planilha
                            if documento >= 15:
                                cnpj = linha[2].value # CNPJ
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN + Keys.DOWN)
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                                driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(cnpj)
                                driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(Keys.LEFT_CONTROL + 'v')
                                time.sleep(0.5)
                                driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cnpj"]').click()
                                time.sleep(10)

                                # Condição se verifica se possui procuração para o CNPJ que está sendo procurado
                                mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'fade-alert').text[15:]
                                if mensagem_procuracao:
                                    print(f'Não possui procuração para o {cnpj}')
                                    linha_celula = linha[4]

                                    if hasattr(linha_celula, 'row'):
                                        linha_atual = linha_celula.row
                                        sheet_empresas[f'F{linha_atual}'] = 'Não possui procuração'
                                        workbook.save(valores['caminho'])
                                        print('Retornando as buscas')
                                        driver.refresh()
                                        pass
                                else:
                                    driver.find_element('xpath', '//*[@id="geral"]/div').click()
                                    WebDriverWait(driver, 120).until(
                                        EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                                    )
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(Keys.DOWN + Keys.ENTER)
                                    WebDriverWait(driver, 120).until(
                                        EC.presence_of_element_located((By.XPATH, '//*[@id="TipoPedido"]'))
                                    )

                                    # Verifica a data de abertura da empresa para pesquisar a partir dali se posterior a 01/01/2018

                                    cnpj_formatado = cnpj.replace('.', '').replace('/', '').replace('-', '')

                                    # Monta a URL com o CNPJ
                                    url = (f"https://www.receitaws.com.br/v1/cnpj/{cnpj_formatado}")

                                    abertura = None

                                    try:
                                        # Faz a requisição HTTP
                                        response = requests.get(url)
                                        response.raise_for_status()  # Lança uma exceção se a requisição falhar

                                        # Converte a resposta para JSON
                                        data = response.json()
                                        abertura = data.get('abertura')
                                        print(f"Data de abertura localizada: {abertura}")

                                    except requests.exceptions.RequestException as e:
                                        print("Erro na requisição", f"Erro: {e}")

                                    # Verifica se a data de abertura é válida antes de tentar converter
                                    start_date = datetime(2018, 1, 1)

                                    if abertura:
                                        try:
                                            # Converte a data de abertura para um objeto datetime
                                            abertura_dt = datetime.strptime(abertura, '%d/%m/%Y')
                                            print("Data de abertura original:", abertura_dt)

                                            # Substitui o dia pelo dia 01
                                            abertura_dt = abertura_dt.replace(day=1)
                                            print("Data de inicio utilizada:", abertura_dt)

                                            # Defina start_date para abertura_dt se for posterior a 2018-01-01, caso contrário, use 2018-01-01
                                            start_date = max(abertura_dt, datetime(2018, 1, 1))
                                            
                                        except ValueError:
                                            print("Erro ao converter a data de abertura.")
                                            start_date = datetime(2018, 1, 1)
                                    else:
                                        start_date = datetime(2018, 1, 1)
                                        
                                    data_corte = datetime.strptime(driver.find_element(By.CLASS_NAME, 'alert-info').text[58:68],
                                                                '%d/%m/%Y')

                                    if valores['dataini']:
                                        start_date = datetime.strptime(valores['dataini'], '%d/%m/%Y')
                                    else:
                                        start_date = datetime(2018, 1, 1)

                                    end_date = start_date + timedelta(days=30 * int(valores['periodo']))

                                    loop = True
                                    while loop:
                                        if end_date > data_corte:
                                            end_date = data_corte
                                            loop = False

                                        start_date_string = start_date.strftime("%d/%m/%Y")
                                        end_date_string = end_date.strftime("%d/%m/%Y")

                                        print("Inicio: ", start_date_string, "   Fim: ", end_date_string)

                                        driver.find_element('xpath', '//*[@id="TipoPedido"]/option[2]').click()

                                        driver.find_element('xpath', '//*[@id="DataInicial"]').send_keys(start_date_string)
                                        driver.find_element('xpath', '//*[@id="DataFinal"]').click()
                                        driver.find_element('xpath', '//*[@id="DataFinal"]').clear()
                                        driver.find_element('xpath', '//*[@id="DataFinal"]').send_keys(end_date_string)
                                        driver.find_element('xpath', '//*[@id="btnSalvar"]').click()
                                        pedido = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]
                                        print(pedido)
                                        if pedido == 'Solicitação enviada com sucesso.':
                                            WebDriverWait(driver, 120).until(
                                                EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                            )
                                            driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()

                                            start_date = end_date + timedelta(days=1)
                                            end_date = start_date + timedelta(days=30 * int(valores['periodo']))

                                        elif pedido == 'Pedido não foi aceito. Já existe um pedido do mesmo tipo.':
                                            driver.find_element(By.XPATH, '//*[@id="btnCancelarAlteracao"]').click()
                                            WebDriverWait(driver, 120).until(
                                                EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                            )
                                            driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()

                                            start_date = end_date + timedelta(days=1)
                                            end_date = start_date + timedelta(days=30 * int(valores['periodo']))

                                        elif pedido == 'O limite de solicitações foi alcançado. Somente é permitido 72 (doze) solicitações por dia.':
                                            time.sleep(2)
                                            driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                            break

                                        if loop == False:
                                            linha_celula = linha[4]

                                            if hasattr(linha_celula, 'row'):
                                                linha_atual = linha_celula.row
                                                sheet_empresas[f'F{linha_atual}'] = 'Solicitação realizada'
                                                sheet_empresas[f"G{linha_atual}"] = start_date_string + " - " + end_date_string
                                                workbook.save(valores['caminho'])

                                    driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()

                            else:
                                # Buscas por CPF
                                cnpj = linha[2].value
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN)
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                                driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(cnpj)
                                driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(Keys.LEFT_CONTROL + 'v')
                                time.sleep(0.5)
                                driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cpf"]').click()
                                time.sleep(10)

                                # Condição se verifica se possui procuração para o CPF que está sendo procurado
                                mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'alert-danger').text[15:]
                                if mensagem_procuracao:
                                    print(f'Não possui procuração para o {cnpj}')
                                    linha_celula = linha[4]

                                    if hasattr(linha_celula, 'row'):
                                        linha_atual = linha_celula.row
                                        sheet_empresas[f'F{linha_atual}'] = 'Não possui procuração'
                                        workbook.save(valores['caminho'])
                                        print('Retornando as buscas')
                                        driver.refresh()
                                        pass
                                else:
                                    driver.find_element('xpath', '//*[@id="geral"]/div').click()
                                    WebDriverWait(driver, 120).until(
                                            EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                                        )
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(Keys.DOWN + Keys.ENTER)
                                    WebDriverWait(driver, 120).until(
                                            EC.presence_of_element_located((By.XPATH, '//*[@id="TipoPedido"]'))
                                        )

                                    # Define as datas para busca
                                    data_corte = datetime.strptime(driver.find_element(By.CLASS_NAME, 'alert-info').text[58:68],
                                                            '%d/%m/%Y')
                                    start_date = datetime(2018, 1, 1)

                                    if valores['dataini']:
                                        start_date = datetime.strptime(valores['dataini'], '%d/%m/%Y')
                                    else:
                                        start_date = datetime(2018, 1, 1)

                                    end_date = start_date + timedelta(days=30 * int(valores['periodo']))

                                    loop = True
                                    while loop:
                                        if end_date > data_corte:
                                            end_date = data_corte
                                            loop = False

                                        start_date_string = start_date.strftime("%d/%m/%Y")
                                        end_date_string = end_date.strftime("%d/%m/%Y")

                                        print("Inicio: ", start_date_string, "   Fim: ", end_date_string)

                                        driver.find_element('xpath', '//*[@id="TipoPedido"]/option[2]').click()

                                        driver.find_element('xpath', '//*[@id="DataInicial"]').send_keys(start_date_string)
                                        driver.find_element('xpath', '//*[@id="DataFinal"]').click()
                                        driver.find_element('xpath', '//*[@id="DataFinal"]').clear()
                                        driver.find_element('xpath', '//*[@id="DataFinal"]').send_keys(end_date_string)
                                        driver.find_element('xpath', '//*[@id="btnSalvar"]').click()
                                        pedido = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]
                                        print(pedido)
                                        if pedido == 'Solicitação enviada com sucesso.':
                                            WebDriverWait(driver, 120).until(
                                                EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                            )
                                            driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()

                                            start_date = end_date + timedelta(days=1)
                                            end_date = start_date + timedelta(days=30 * int(valores['periodo']))

                                        elif pedido == 'Pedido não foi aceito. Já existe um pedido do mesmo tipo.':
                                            driver.find_element(By.XPATH, '//*[@id="btnCancelarAlteracao"]').click()
                                            WebDriverWait(driver, 120).until(
                                                EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                            )
                                            driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()

                                            start_date = end_date + timedelta(days=1)
                                            end_date = start_date + timedelta(days=30 * int(valores['periodo']))
                                            
                                        elif pedido == 'O limite de solicitações foi alcançado. Somente é permitido 72 (doze) solicitações por dia.':
                                            time.sleep(2)
                                            driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                            break
                                            
                                        if loop == False:
                                            linha_celula = linha[4]

                                            if hasattr(linha_celula, 'row'):
                                                linha_atual = linha_celula.row
                                                sheet_empresas[f'F{linha_atual}'] = 'Solicitação realizada'
                                                sheet_empresas[f"G{linha_atual}"] = start_date_string + " - " + end_date_string
                                                workbook.save(valores['caminho'])

                                    driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()

                        print('Buscas Finalizadas')
                        time.sleep(7)
                        driver.quit()

                    if valores['finalidade'] == '2' and valores['proprio'] is False:
                        chrome_options = ChromeOptions()
                        prefs = {'download.default_directory': download_dir}
                        chrome_options.add_experimental_option('prefs', prefs)
                        chromedriver_path = download_undetected_chromedriver(folder_path, undetected=True, arm=False,
                                                                            force_update=True)
                        driver = uc.Chrome(driver_executable_path=chromedriver_path, options=chrome_options,headless=False)            
                        driver.get(url)
                        driver.maximize_window()
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-acoes"]/div[2]/p/button'))
                        )
                        driver.find_element('xpath', '//*[@id="login-acoes"]/div[2]/p/button').click()

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-certificate"]'))
                        )
                        driver.find_element('xpath', '//*[@id="login-certificate"]').click()

                        print("Selecione o certificado para continuar")

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="sairAplicacao"]'))
                        )            
                        inscricao = driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/p[2]/span[1]').text.strip('-')
                        print(f'CNPJ do procurador: {inscricao}')
                        if len(driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/p[2]/span[1]').text) < 18:
                            driver.find_element(By.XPATH, '//*[@id="geral"]/div').click()
                            WebDriverWait(driver, 120).until(
                                EC.presence_of_element_located((By.XPATH, '//*[@id="header"]/div[2]/a'))
                            )
                            driver.find_element('xpath', '//*[@id="header"]/div[2]/a').click()
                            pass
                        else:
                            driver.find_element(By.CLASS_NAME, 'alterar-perfil').click()

                        for linha in sheet_empresas.iter_rows(min_row=int(valores['linhaini']), max_row=int(valores['linhafim'])):
                            pasta_empresa = criar_pasta(f'{linha[0].value} - {linha[1].value}')
                            documento = len(str(linha[2].value))
                            if documento >= 15:
                                cnpj = linha[2].value
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN + Keys.DOWN)
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                                driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(cnpj)
                                driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(Keys.LEFT_CONTROL + 'v')
                                time.sleep(0.5)
                                driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cnpj"]').click()
                                time.sleep(10)
                                # Condição se verifica se possui procuração para o CNPJ que está sendo procurado
                                mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'fade-alert').text[15:]
                                if mensagem_procuracao:
                                    print(f'Não possui procuração para o {cnpj}')
                                    linha_celula = linha[4]

                                    if hasattr(linha_celula, 'row'):
                                        linha_atual = linha_celula.row
                                        sheet_empresas[f'H{linha_atual}'] = 'Não possui procuração'
                                        workbook.save(valores['caminho'])
                                        print('Retornando as buscas')
                                        driver.refresh()
                                        pass
                                else:
                                    time.sleep(5)
                                    driver.find_element('xpath', '//*[@id="geral"]/div').click()
                                    WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                                )
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(
                                    Keys.DOWN + Keys.DOWN + Keys.ENTER)
                                    WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located(
                                        (By.XPATH, '//*[@id="conteudo-pagina"]/form/section/div/div[4]/input'))
                                )
                                    driver.find_element('xpath', '//*[@id="conteudo-pagina"]/form/section/div/div[4]/input').click()
                                    time.sleep(5)
                                    print(f'"Iniciando download da empresa CNPJ: {cnpj}')

                                    download_links = driver.find_elements('xpath',
                                                                            '//*[@id="DataTables_Table_0"]/tbody/tr/td/a')
                                    total_files = (len(download_links))
                                    print(f'Total de arquivos: {total_files}')
                                    soma_files = 0

                                    for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                                        link.click()
                                        time.sleep(3.5)
                                        soma_files += 1
                                        print(f'Baixando {soma_files}/{total_files} arquivos')

                                        arquivos_baixados = os.listdir(download_dir)
                                        for arquivo in arquivos_baixados:
                                            if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                                caminho_arquivo = os.path.join(download_dir, arquivo)
                                                shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))
                                    try:
                                        WebDriverWait(driver, 10).until(
                                                EC.presence_of_element_located(
                                                    (By.XPATH, '//*[@id="DataTables_Table_0_paginate"]/span/a[2]')))
                                        driver.find_element('xpath', '//*[@id="DataTables_Table_0_paginate"]/span/a[2]').click()
                                        for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                                            link.click()
                                            time.sleep(3.5)
                                            soma_files += 1
                                            print(f'Baixando {soma_files}/{total_files} arquivos')
                                            arquivos_baixados = os.listdir(download_dir)
                                            for arquivo in arquivos_baixados:
                                                if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                                    caminho_arquivo = os.path.join(download_dir, arquivo)
                                                    shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))

                                        loop = False

                                        if loop == False:
                                            linha_celula = linha[4]

                                            if hasattr(linha_celula, 'row'):
                                                linha_atual = linha_celula.row
                                                sheet_empresas[f'H{linha_atual}'] = str(soma_files) + "/" + str(total_files)
                                                workbook.save(valores['caminho'])

                                        print(f'"Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!"')
                                        time.sleep(2)
                                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                        time.sleep(2)
                                    except:
                                        print(f'"Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!"')
                                        time.sleep(2)
                                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                        time.sleep(2)
                                        continue

                            else:
                                cnpj = linha[2].value
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN)
                                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                                driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(cnpj)
                                driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(Keys.LEFT_CONTROL + 'v')
                                time.sleep(0.5)
                                driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cpf"]').click()
                                time.sleep(10)
                                # Condição se verifica se possui procuração para o CNPJ que está sendo procurado
                                mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'fade-alert').text[15:]
                                if mensagem_procuracao:
                                    print(f'Não possui procuração para o {cnpj}')
                                    linha_celula = linha[4]

                                    if hasattr(linha_celula, 'row'):
                                        linha_atual = linha_celula.row
                                        sheet_empresas[f'H{linha_atual}'] = 'Não possui procuração'
                                        workbook.save(valores['caminho'])
                                        print('Retornando as buscas')
                                        driver.refresh()
                                        pass
                                else:
                                    time.sleep(5)
                                    driver.find_element('xpath', '//*[@id="geral"]/div').click()
                                    WebDriverWait(driver, 120).until(
                                        EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                                    )
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                                    driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(
                                        Keys.DOWN + Keys.DOWN + Keys.ENTER)
                                    WebDriverWait(driver, 120).until(
                                        EC.presence_of_element_located(
                                            (By.XPATH, '//*[@id="conteudo-pagina"]/form/section/div/div[4]/input'))
                                    )
                                    driver.find_element('xpath', '//*[@id="conteudo-pagina"]/form/section/div/div[4]/input').click()
                                    time.sleep(5)
                                    print(f'"Iniciando download da empresa CNPJ: {cnpj}')

                                    download_links = driver.find_elements('xpath',
                                                                        '//*[@id="DataTables_Table_0"]/tbody/tr/td/a')
                                    total_files = (len(download_links))
                                    print(f'Total de arquivos: {total_files}')
                                    soma_files = 0

                                    for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                                        link.click()
                                        time.sleep(7)
                                        soma_files += 1
                                        print(f'Baixando {soma_files}/{total_files} arquivos')

                                        arquivos_baixados = os.listdir(download_dir)
                                        for arquivo in arquivos_baixados:
                                            if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                                caminho_arquivo = os.path.join(download_dir, arquivo)
                                                shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))
                                    try:
                                        WebDriverWait(driver, 10).until(
                                            EC.presence_of_element_located(
                                                (By.XPATH, '//*[@id="DataTables_Table_0_paginate"]/span/a[2]')))
                                        driver.find_element('xpath', '//*[@id="DataTables_Table_0_paginate"]/span/a[2]').click()
                                        download_links = driver.find_elements('xpath',
                                                                        '//*[@id="DataTables_Table_0"]/tbody/tr/td/a')
                                        total_files2 = (len(download_links))
                                        print(f'Total de arquivos: {total_files2}')                                        

                                        for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                                            link.click()
                                            time.sleep(7)
                                            soma_files += 1
                                            print(f'Baixando {soma_files}/{total_files + total_files2} arquivos')
                                            arquivos_baixados = os.listdir(download_dir)
                                            for arquivo in arquivos_baixados:
                                                if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                                    caminho_arquivo = os.path.join(download_dir, arquivo)
                                                    shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))

                                        loop = False

                                        if loop == False:
                                            linha_celula = linha[4]

                                            if hasattr(linha_celula, 'row'):
                                                linha_atual = linha_celula.row
                                                sheet_empresas[f'H{linha_atual}'] = str(soma_files) + "/" + str(total_files)
                                                workbook.save(valores['caminho'])

                                        print(f'"Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!"')
                                        time.sleep(2)
                                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                        time.sleep(2)
                                    except:
                                        print(f'"Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!"')
                                        time.sleep(2)
                                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                        time.sleep(2)
                                        continue

                        print("Baixa de arquivos finalizada com sucesso")
                        time.sleep(3)
                        driver.quit()

                    # Condição para realizar as buscas quando o certificado é da própria empresa
                    if valores['finalidade'] == '1' and valores['proprio'] is True:
                        chrome_options = ChromeOptions()
                        chromedriver_path = download_undetected_chromedriver(folder_path, undetected=True, arm=False,
                                                                            force_update=True)
                        driver = uc.Chrome(driver_executable_path=chromedriver_path, options=chrome_options, headless=False)            
                        driver.get(url)
                        driver.maximize_window()
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-acoes"]/div[2]/p/button'))
                        )
                        driver.find_element('xpath', '//*[@id="login-acoes"]/div[2]/p/button').click()

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-certificate"]'))
                        )
                        driver.find_element('xpath', '//*[@id="login-certificate"]').click()

                        print("Selecione o certificado para continuar")

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="sairAplicacao"]'))
                        )
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                            )
                        driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                        driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(Keys.DOWN + Keys.ENTER)
                        WebDriverWait(driver, 120).until(
                                EC.presence_of_element_located((By.XPATH, '//*[@id="TipoPedido"]'))
                            )
                        # Verifica a data de abertura da empresa para pesquisar a partir dali se posterior a 01/01/2018
                        cnpj = driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/p[2]/span[1]').text.strip('-').rstrip()
                        print(cnpj)

                        cnpj_formatado = cnpj.replace('.', '').replace('/', '').replace('-', '')

                        # Monta a URL com o CNPJ
                        url = (f"https://www.receitaws.com.br/v1/cnpj/{cnpj_formatado}")

                        abertura = None

                        try:
                            # Faz a requisição HTTP
                            response = requests.get(url)
                            response.raise_for_status()  # Lança uma exceção se a requisição falhar

                            # Converte a resposta para JSON
                            data = response.json()
                            abertura = data.get('abertura')
                            print(f"Data de abertura localizada: {abertura}")

                        except requests.exceptions.RequestException as e:
                            print("Erro na requisição", f"Erro: {e}")

                        # Verifica se a data de abertura é válida antes de tentar converter
                        start_date = datetime(2018, 1, 1)

                        if abertura:
                            try:
                                # Converte a data de abertura para um objeto datetime
                                abertura_dt = datetime.strptime(abertura, '%d/%m/%Y')
                                print("Data de abertura original:", abertura_dt)

                                # Substitui o dia pelo dia 01
                                abertura_dt = abertura_dt.replace(day=1)
                                print("Data de inicio utilizada:", abertura_dt)

                                # Defina start_date para abertura_dt se for posterior a 2018-01-01, caso contrário, use 2018-01-01
                                start_date = max(abertura_dt, datetime(2018, 1, 1))
                                
                            except ValueError:
                                print("Erro ao converter a data de abertura.")
                                start_date = datetime(2018, 1, 1)
                        else:
                            start_date = datetime(2018, 1, 1)                

                        # Define as datas para busca
                        data_corte = datetime.strptime(driver.find_element(By.CLASS_NAME, 'alert-info').text[58:68], '%d/%m/%Y')

                        if valores['dataini']:
                            start_date = datetime.strptime(valores['dataini'], '%d/%m/%Y')
                        else:
                            start_date = datetime(2018, 1, 1)

                        end_date = start_date + timedelta(days=30 * int(valores['periodo']))

                        loop = True
                        while loop:
                            if end_date > data_corte:
                                end_date = data_corte
                                loop = False

                            start_date_string = start_date.strftime("%d/%m/%Y")
                            end_date_string = end_date.strftime("%d/%m/%Y")

                            print("Inicio: ", start_date_string, "   Fim: ", end_date_string)

                            driver.find_element('xpath', '//*[@id="TipoPedido"]/option[2]').click()
                            driver.find_element('xpath', '//*[@id="DataInicial"]').send_keys(start_date_string)
                            driver.find_element('xpath', '//*[@id="DataFinal"]').click()
                            driver.find_element('xpath', '//*[@id="DataFinal"]').clear()
                            driver.find_element('xpath', '//*[@id="DataFinal"]').send_keys(end_date_string)
                            driver.find_element('xpath', '//*[@id="btnSalvar"]').click()
                            pedido = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]
                            print(pedido)
                            if pedido == 'Solicitação enviada com sucesso.':
                                WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                )
                                driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()

                                start_date = end_date + timedelta(days=1)
                                end_date = start_date + timedelta(days=30 * int(valores['periodo']))

                            elif pedido == 'Pedido não foi aceito. Já existe um pedido do mesmo tipo.':
                                driver.find_element(By.XPATH, '//*[@id="btnCancelarAlteracao"]').click()
                                WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                )
                                driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()

                                start_date = end_date + timedelta(days=1)
                                end_date = start_date + timedelta(days=30 * int(valores['periodo']))
                                            
                            elif pedido == 'O limite de solicitações foi alcançado. Somente é permitido 72 (doze) solicitações por dia.':
                                time.sleep(2)
                                driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                break

                        print('Buscas Finalizadas')
                        time.sleep(3)
                        driver.quit()

                    # Condição para fazer o download dos arquivos quando o certificado é próprio
                    if valores['finalidade'] == '2' and valores['proprio'] is True:
                        chrome_options = ChromeOptions()
                        prefs = {'download.default_directory': download_dir}
                        chrome_options.add_experimental_option('prefs', prefs)
                        chromedriver_path = download_undetected_chromedriver(folder_path, undetected=True, arm=False,
                                                                            force_update=True)
                        driver = uc.Chrome(driver_executable_path=chromedriver_path, options=chrome_options, headless=False)            
                        driver.get(url)
                        driver.maximize_window()
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-acoes"]/div[2]/p/button'))
                        )
                        driver.find_element('xpath', '//*[@id="login-acoes"]/div[2]/p/button').click()

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="login-certificate"]'))
                        )
                        driver.find_element('xpath', '//*[@id="login-certificate"]').click()

                        print("Selecione o certificado para continuar")

                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="sairAplicacao"]'))
                        )
                        for linha in sheet_empresas.iter_rows(min_row=int(valores['linhaini']), max_row=int(valores['linhafim'])):
                            pasta_empresa = criar_pasta(f'{linha[0].value} - {linha[1].value}')
                            WebDriverWait(driver, 120).until(
                                EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                            )
                            driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                            driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(
                                Keys.DOWN + Keys.DOWN + Keys.ENTER)
                            WebDriverWait(driver, 120).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, '//*[@id="conteudo-pagina"]/form/section/div/div[4]/input'))
                            )
                            driver.find_element('xpath', '//*[@id="conteudo-pagina"]/form/section/div/div[4]/input').click()
                            time.sleep(10)

                            download_links = driver.find_elements('xpath', '//*[@id="DataTables_Table_0"]/tbody/tr/td/a')
                            total_files = (len(download_links))
                            print(f'Total de arquivos: {total_files}')
                            soma_files = 0

                            for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                                link.click()
                                time.sleep(7)
                                soma_files += 1
                                print(f'Baixando {soma_files}/{total_files} arquivos')
                                arquivos_baixados = os.listdir(download_dir)
                                for arquivo in arquivos_baixados:
                                    if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                        caminho_arquivo = os.path.join(download_dir, arquivo)
                                        shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))

                            try:
                                WebDriverWait(driver, 3).until(
                                    EC.presence_of_element_located(
                                        (By.XPATH, '//*[@id="DataTables_Table_0_paginate"]/span/a[2]')))
                                driver.find_element('xpath', '//*[@id="DataTables_Table_0_paginate"]/span/a[2]').click()
                                download_links = driver.find_elements('xpath', '//*[@id="DataTables_Table_0"]/tbody/tr/td/a')
                                total_files2 = (len(download_links))
                                for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                                    link.click()
                                    time.sleep(7)
                                    soma_files += 1
                                    print(f'Baixando {soma_files}/{total_files + total_files2} arquivos')
                                    arquivos_baixados = os.listdir(download_dir)
                                    for arquivo in arquivos_baixados:
                                        if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                            caminho_arquivo = os.path.join(download_dir, arquivo)
                                            shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))

                            except:
                                time.sleep(2)
                                driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                time.sleep(2)
                                continue

                            time.sleep(2)
                            driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                            time.sleep(2)

                        print("Baixa de arquivos finalizada com sucesso")
                        time.sleep(5)

                    print("Programa Finalizado")
                    time.sleep(5)
                    driver.quit()                    