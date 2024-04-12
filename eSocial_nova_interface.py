import time
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import undetected_chromedriver as uc
from auto_download_undetected_chromedriver import download_undetected_chromedriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from undetected_chromedriver import ChromeOptions
import openpyxl
from openpyxl.utils.cell import get_column_letter
import os
import shutil
import requests

def solicitar_ou_baixar():
    # Função que será chamada quando o botão for clicado
    opcao = solicitar_baixar_var.get()
    if opcao == 1 and certificado_proprio_var.get() is False:        
        url = 'https://login.esocial.gov.br/login.aspx'
        folder_path = "c:\\chromedriver"
        chrome_options = ChromeOptions()
        chromedriver_path = download_undetected_chromedriver(folder_path, undetected=True, arm=False,
                                                                        force_update=True)
        driver = uc.Chrome(driver_executable_path=chromedriver_path, options=chrome_options, headless=False)            
        driver.get(url)
        driver.maximize_window()
        root.iconify()
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
        
        workbook = openpyxl.load_workbook(caminho_planilha_var.get())
        sheet_empresas = workbook.active        
        documento = sheet_empresas.cell(row=int(linha_ini.get()), column=3).value        
        empresas = list(sheet_empresas.iter_rows(min_row=int(linha_ini.get()), max_row=int(linha_fim.get())))
        
        if data_ini.get():
            data_inicial = datetime.strptime(data_ini.get(), '%d/%m/%Y')
        else:
            data_inicial = datetime(2018, 1, 1)

        # Dicionário para armazenar a última data final de cada empresa
        ultima_data_final_por_empresa = {}    

        data_corte = ''                   
        loop = True
        while loop:
            for linha in empresas:
                documento = len(str(linha[2].value))
                data_atual = data_inicial
                # Condição que verifica se é CNPJ ou CPF na planilha                                
                if documento >= 15:
                    cnpj = linha[2].value # CNPJ
                    driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                    driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN + Keys.DOWN)
                    driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                    driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(cnpj)
                    driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(Keys.LEFT_CONTROL + 'v')
                    WebDriverWait(driver, 120).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-verificar-procuracao-cnpj"]'))
                        )
                    driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cnpj"]').click()
                    mensagem_procuracao = ''
                    try:
                        WebDriverWait(driver, 15).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="geral"]/div'))
                        )
                    except:
                        mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]

                    # Condição se verifica se possui procuração para o CNPJ que está sendo procurado
                    print(mensagem_procuracao)
                    if mensagem_procuracao == 'O procurador não possui perfil com autorização de acesso à Web':
                        print(f'Não possui procuração para o {cnpj}')
                        linha_celula = linha[4]

                        if hasattr(linha_celula, 'row'):
                            linha_atual = linha_celula.row
                            sheet_empresas[f'F{linha_atual}'] = 'Não possui procuração'
                            workbook.save(caminho_planilha_var.get())
                            print('Retornando as buscas')
                            driver.refresh()
                            continue
                    else:
                        print(f'CNPJ/CPF sendo buscado: {str(linha[2].value)}')
                        driver.find_element('xpath', '//*[@id="geral"]/div').click()
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                        )
                        driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                        driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(Keys.DOWN + Keys.ENTER)
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="TipoPedido"]'))
                        )
                            
                        data_corte = datetime.strptime(driver.find_element(By.CLASS_NAME, 'alert-info').text[58:68],'%d/%m/%Y')    

                        # Obtém a última data final para a empresa
                        if linha not in ultima_data_final_por_empresa:
                            data_atual = data_inicial
                            ultima_data_final_por_empresa[linha] = data_atual
                        data_final = ultima_data_final_por_empresa[linha] + timedelta(days=30 * int(meses_buscar_var.get()))

                        # Fazer cinco requisições para a empresa atual
                        for i in range(5):
                            if data_final > data_corte:
                                data_final = data_corte
                                loop = False

                            data_inicial_str = ultima_data_final_por_empresa[linha].strftime('%d/%m/%Y')
                            data_final_str = data_final.strftime('%d/%m/%Y')

                            driver.find_element('xpath', '//*[@id="TipoPedido"]/option[2]').click()
                            driver.find_element('xpath', '//*[@id="DataInicial"]').send_keys(data_inicial_str)
                            driver.find_element('xpath', '//*[@id="DataFinal"]').click()
                            driver.find_element('xpath', '//*[@id="DataFinal"]').clear()
                            driver.find_element('xpath', '//*[@id="DataFinal"]').send_keys(data_final_str)
                            print(f'Data inicial: {data_inicial_str} - Data Final: {data_final_str}')
                            driver.find_element('xpath', '//*[@id="btnSalvar"]').click()
                            pedido = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]
                            print(pedido)
                            if pedido == 'Solicitação enviada com sucesso.':
                                WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                )
                                driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()
                                ultima_data_final_por_empresa[linha] = data_final + timedelta(days=1)
                                data_final = ultima_data_final_por_empresa[linha] + timedelta(days=30 * int(meses_buscar_var.get()))

                                linha_celula = linha[4]     
                                if hasattr(linha_celula, 'row'):
                                    linha_atual = linha_celula.row
                                    sheet_empresas[f'F{linha_atual}'] = 'Solicitação realizada'
                                    sheet_empresas[f"G{linha_atual}"] = "Última data solicitada: " + data_final_str
                                    workbook.save(caminho_planilha_var.get())

                            elif pedido == 'Pedido não foi aceito. Já existe um pedido do mesmo tipo.':
                                driver.find_element(By.XPATH, '//*[@id="btnCancelarAlteracao"]').click()
                                WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                )
                                driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()
                                ultima_data_final_por_empresa[linha] = data_final + timedelta(days=1)
                                data_final = ultima_data_final_por_empresa[linha] + timedelta(days=30 * int(meses_buscar_var.get()))                             

                            elif pedido == 'O limite de solicitações foi alcançado. Somente é permitido 72 (doze) solicitações por dia.':
                                time.sleep(2)
                                driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                break
                        
                    # Atualiza a data atual para a próxima iteração
                    data_atual = ultima_data_final_por_empresa[empresas[0]] + timedelta(days=1)
                    driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()                                

                else:
                    # Buscas por CPF
                    cnpj = linha[2].value
                    driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                    driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN)
                    driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                    driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(cnpj)
                    driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(Keys.LEFT_CONTROL + 'v')
                    WebDriverWait(driver, 120).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-verificar-procuracao-cpf"]'))
                        )
                    driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cpf"]').click()
                    mensagem_procuracao = ''
                    try:
                        WebDriverWait(driver, 15).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="geral"]/div'))
                        )
                    except:
                        mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]
                    
                    # Condição se verifica se possui procuração para o CNPJ que está sendo procurado
                    print(mensagem_procuracao)
                    if mensagem_procuracao == 'O procurador não possui perfil com autorização de acesso à Web':
                        print(f'Não possui procuração para o {cnpj}')
                        linha_celula = linha[4]

                        if hasattr(linha_celula, 'row'):
                            linha_atual = linha_celula.row
                            sheet_empresas[f'F{linha_atual}'] = 'Não possui procuração'
                            workbook.save(caminho_planilha_var.get())
                            print('Retornando as buscas')
                            driver.refresh()
                            continue
                    else:
                        print(f'CNPJ/CPF sendo buscado: {str(linha[2].value)}')
                        driver.find_element('xpath', '//*[@id="geral"]/div').click()
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="menuDownload"]'))
                        )
                        driver.find_element('xpath', '//*[@id="menuDownload"]').click()
                        driver.find_element('xpath', '//*[@id="menuDownload"]').send_keys(Keys.DOWN + Keys.ENTER)
                        WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="TipoPedido"]'))
                        )

                        data_corte = datetime.strptime(driver.find_element(By.CLASS_NAME, 'alert-info').text[58:68],'%d/%m/%Y')    

                        # Obtém a última data final para a empresa
                        if linha not in ultima_data_final_por_empresa:
                            data_atual = data_inicial
                            ultima_data_final_por_empresa[linha] = data_atual
                        data_final = ultima_data_final_por_empresa[linha] + timedelta(days=30)

                        # Fazer cinco requisições para a empresa atual
                        for i in range(5):
                            if data_final > data_corte:
                                data_final = data_corte
                                break

                            data_inicial_str = ultima_data_final_por_empresa[linha].strftime('%d/%m/%Y')
                            data_final_str = data_final.strftime('%d/%m/%Y')

                            driver.find_element('xpath', '//*[@id="TipoPedido"]/option[2]').click()
                            driver.find_element('xpath', '//*[@id="DataInicial"]').send_keys(data_inicial_str)
                            driver.find_element('xpath', '//*[@id="DataFinal"]').click()
                            driver.find_element('xpath', '//*[@id="DataFinal"]').clear()
                            driver.find_element('xpath', '//*[@id="DataFinal"]').send_keys(data_final_str)
                            print(f'Data inicial: {data_inicial_str} - Data Final: {data_final_str}')
                            driver.find_element('xpath', '//*[@id="btnSalvar"]').click()
                            pedido = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]
                            print(pedido)
                            if pedido == 'Solicitação enviada com sucesso.':
                                WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                )
                                driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()
                                ultima_data_final_por_empresa[linha] = data_final + timedelta(days=1)
                                data_final = ultima_data_final_por_empresa[linha] + timedelta(days=30 * int(meses_buscar_var.get()))

                                linha_celula = linha[4]
                                if hasattr(linha_celula, 'row'):
                                    linha_atual = linha_celula.row
                                    sheet_empresas[f'F{linha_atual}'] = 'Solicitação realizada'
                                    sheet_empresas[f"G{linha_atual}"] = "Última data solicitada: " + data_final_str
                                    workbook.save(caminho_planilha_var.get())

                            elif pedido == 'Pedido não foi aceito. Já existe um pedido do mesmo tipo.':
                                driver.find_element(By.XPATH, '//*[@id="btnCancelarAlteracao"]').click()
                                WebDriverWait(driver, 120).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                                )
                                driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()
                                ultima_data_final_por_empresa[linha] = data_final + timedelta(days=1)
                                data_final = ultima_data_final_por_empresa[linha] + timedelta(days=30 * int(meses_buscar_var.get()))

                            elif pedido == 'O limite de solicitações foi alcançado. Somente é permitido 72 (doze) solicitações por dia.':
                                time.sleep(2)
                                driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                                break
                        
                    # Atualiza a data atual para a próxima iteração
                    data_atual = ultima_data_final_por_empresa[empresas[0]] + timedelta(days=1)
                    driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
        
        print('Buscas Finalizadas')
        time.sleep(7)
        driver.quit()        

    elif opcao == 2 and certificado_proprio_var.get() is False:
        url = 'https://login.esocial.gov.br/login.aspx'
        folder_path = "c:\\chromedriver"

        download_dir = caminho_pasta_salvar_var.get().rstrip().replace('/', '\\')
        
        def criar_pasta(nome_empresa):
            pasta_empresa = os.path.join(download_dir, nome_empresa)
            if not os.path.exists(pasta_empresa):
                os.makedirs(pasta_empresa)
            return pasta_empresa        
        
        chrome_options = ChromeOptions()
        prefs = {'download.default_directory': download_dir}
        chrome_options.add_experimental_option('prefs', prefs)
        chromedriver_path = download_undetected_chromedriver(folder_path, undetected=True, arm=False,
                                                            force_update=True)
        driver = uc.Chrome(driver_executable_path=chromedriver_path, options=chrome_options,headless=False)            
        driver.get(url)
        driver.maximize_window()
        root.iconify()
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
        
        workbook = openpyxl.load_workbook(caminho_planilha_var.get())
        sheet_empresas = workbook.active        
        documento = sheet_empresas.cell(row=int(linha_ini.get()), column=3).value        
        empresas = list(sheet_empresas.iter_rows(min_row=int(linha_ini.get()), max_row=int(linha_fim.get())))            

        for linha in sheet_empresas.iter_rows(min_row=int(linha_ini.get()), max_row=int(linha_fim.get())):
            pasta_empresa = criar_pasta(f'{linha[0].value} - {linha[1].value.rstrip()}')
            documento = len(str(linha[2].value))
            if documento >= 15: #CNPJ
                cnpj = linha[2].value
                driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN + Keys.DOWN)
                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(cnpj)
                driver.find_element('xpath', '//*[@id="procuradorCnpj"]').send_keys(Keys.LEFT_CONTROL + 'v')
                WebDriverWait(driver, 120).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-verificar-procuracao-cnpj"]'))
                        )
                driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cnpj"]').click()
                mensagem_procuracao = ''
                try:
                    WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="geral"]/div'))
                    )
                except:
                    mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]

                # Condição se verifica se possui procuração para o CNPJ que está sendo procurado                                
                if mensagem_procuracao:
                    print(f'Não possui procuração para o {cnpj}')
                    linha_celula = linha[4]

                    if hasattr(linha_celula, 'row'):
                        linha_atual = linha_celula.row
                        sheet_empresas[f'H{linha_atual}'] = 'Não possui procuração'
                        workbook.save(caminho_planilha_var.get())
                        print('Retornando as buscas')
                        driver.refresh()
                        continue
                else:
                    WebDriverWait(driver, 120).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="geral"]/div'))
                    )
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
                    print(f'Iniciando download da empresa CNPJ: {cnpj}')                    

                    # Página 1
                    for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                        link.click()
                        time.sleep(7)
                        arquivos_baixados = os.listdir(download_dir)
                        for arquivo in arquivos_baixados:
                            if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                caminho_arquivo = os.path.join(download_dir, arquivo)
                                shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))
                    # Página 2
                    try:
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0_paginate"]/span/a[2]')))
                        driver.find_element('xpath', '//*[@id="DataTables_Table_0_paginate"]/span/a[2]').click()
                        for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                            link.click()
                            time.sleep(7)                            
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
                                sheet_empresas[f'H{linha_atual}'] = 'Baixados todos arquivos'
                                workbook.save(caminho_planilha_var.get())
                        
                    except:
                        print(f'Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!')
                        time.sleep(2)
                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                        time.sleep(2)
                        continue
                    
                    # Página 3
                    try:
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0_paginate"]/span/a[3]')))
                        driver.find_element('xpath', '//*[@id="DataTables_Table_0_paginate"]/span/a[3]').click()
                        for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                            link.click()
                            time.sleep(7)                            
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
                                sheet_empresas[f'H{linha_atual}'] = 'Baixados todos arquivos'
                                workbook.save(caminho_planilha_var.get())

                        print(f'Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!')
                        time.sleep(2)
                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                        time.sleep(2)

                    except:
                        print(f'Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!')
                        time.sleep(2)
                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                        time.sleep(2)
                        continue                

            else: #CPF
                cnpj = linha[2].value
                driver.find_element('xpath', '//*[@id="perfilAcesso"]').click()
                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.DOWN)
                driver.find_element('xpath', '//*[@id="perfilAcesso"]').send_keys(Keys.ENTER)
                driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(cnpj)
                driver.find_element('xpath', '//*[@id="procuradorCpf"]').send_keys(Keys.LEFT_CONTROL + 'v')
                WebDriverWait(driver, 120).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-verificar-procuracao-cpf"]'))
                        )
                driver.find_element('xpath', '//*[@id="btn-verificar-procuracao-cpf"]').click()
                mensagem_procuracao = ''
                try:
                    WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="geral"]/div'))
                    )
                except:
                    mensagem_procuracao = driver.find_element(By.CLASS_NAME, 'fade-alert').text[2:]

                # Condição se verifica se possui procuração para o CNPJ que está sendo procurado                                
                if mensagem_procuracao:
                    print(f'Não possui procuração para o {cnpj}')
                    linha_celula = linha[4]

                    if hasattr(linha_celula, 'row'):
                        linha_atual = linha_celula.row
                        sheet_empresas[f'H{linha_atual}'] = 'Não possui procuração'
                        workbook.save(caminho_planilha_var.get())
                        print('Retornando as buscas')
                        driver.refresh()
                        continue
                else:
                    WebDriverWait(driver, 120).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="geral"]/div'))
                    )
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
                    print(f'Iniciando download da empresa CNPJ: {cnpj}')                    

                    # Página 1
                    for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                        link.click()
                        time.sleep(7)
                        arquivos_baixados = os.listdir(download_dir)
                        for arquivo in arquivos_baixados:
                            if arquivo.endswith(".zip"):  # Verifica se o arquivo é um ZIP
                                caminho_arquivo = os.path.join(download_dir, arquivo)
                                shutil.move(caminho_arquivo, os.path.join(pasta_empresa, arquivo))
                    # Página 2
                    try:
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0_paginate"]/span/a[2]')))
                        driver.find_element('xpath', '//*[@id="DataTables_Table_0_paginate"]/span/a[2]').click()
                        for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                            link.click()
                            time.sleep(7)                            
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
                                sheet_empresas[f'H{linha_atual}'] = 'Baixados todos arquivos'
                                workbook.save(caminho_planilha_var.get())
                        
                    except:
                        print(f'Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!')
                        time.sleep(2)
                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                        time.sleep(2)
                        continue
                    
                    # Página 3
                    try:
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0_paginate"]/span/a[3]')))
                        driver.find_element('xpath', '//*[@id="DataTables_Table_0_paginate"]/span/a[3]').click()
                        for link in driver.find_elements(By.CLASS_NAME, 'icone-baixar'):
                            link.click()
                            time.sleep(7)                            
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
                                sheet_empresas[f'H{linha_atual}'] = 'Baixados todos arquivos'
                                workbook.save(caminho_planilha_var.get())

                        print(f'Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!')
                        time.sleep(2)
                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                        time.sleep(2)
                        
                    except:
                        print(f'Arquivos da empresa CNPJ: {cnpj} baixados com sucesso!')
                        time.sleep(2)
                        driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                        time.sleep(2)
                        continue

        print("Baixa de arquivos finalizada com sucesso")
        time.sleep(3)
        driver.quit()    

    elif opcao == 1 and certificado_proprio_var.get():
        url = 'https://login.esocial.gov.br/login.aspx'
        folder_path = "c:\\chromedriver"
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

        root.iconify()

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

        if data_ini.get():
            start_date = datetime.strptime(data_ini.get(), '%d/%m/%Y')
        else:
            start_date = datetime(2018, 1, 1)

        end_date = start_date + timedelta(days=30 * int(meses_buscar_var.get()))

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
                end_date = start_date + timedelta(days=30 * int(meses_buscar_var.get()))

            elif pedido == 'Pedido não foi aceito. Já existe um pedido do mesmo tipo.':
                driver.find_element(By.XPATH, '//*[@id="btnCancelarAlteracao"]').click()
                WebDriverWait(driver, 120).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/a'))
                )
                driver.find_element('xpath', '//*[@id="conteudo-pagina"]/div[1]/a').click()

                start_date = end_date + timedelta(days=1)
                end_date = start_date + timedelta(days=30 * int(meses_buscar_var.get()))
                            
            elif pedido == 'O limite de solicitações foi alcançado. Somente é permitido 72 (doze) solicitações por dia.':
                time.sleep(2)
                driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a').click()
                break

        print('Buscas Finalizadas')
        time.sleep(3)
        driver.quit()

    elif opcao == 2 and certificado_proprio_var.get():
        url = 'https://login.esocial.gov.br/login.aspx'
        folder_path = "c:\\chromedriver"
        download_dir = caminho_pasta_salvar_var.get().rstrip().replace('/', '\\')
        
        def criar_pasta(nome_empresa):
            pasta_empresa = os.path.join(download_dir, nome_empresa)
            if not os.path.exists(pasta_empresa):
                os.makedirs(pasta_empresa)
            return pasta_empresa
        
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

        root.iconify()

        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="sairAplicacao"]'))
        )

        workbook = openpyxl.load_workbook(caminho_planilha_var.get())
        sheet_empresas = workbook.active        
        documento = sheet_empresas.cell(row=int(linha_ini.get()), column=3).value        
        empresas = list(sheet_empresas.iter_rows(min_row=int(linha_ini.get()), max_row=int(linha_fim.get())))

        for linha in sheet_empresas.iter_rows(min_row=int(linha_ini.get()), max_row=int(linha_fim.get())):
            pasta_empresa = criar_pasta(f'{linha[0].value} - {linha[1].value.rstrip()}')
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

def meses_buscar():
    # Função que será chamada quando o botão for clicado
    meses_buscar_var.get()     

def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos XLSX", "*.xlsx")])
    if arquivo:
        caminho_planilha_var.set(arquivo)

def selecionar_pasta_salvar():
    pasta_salvar = filedialog.askdirectory()
    if pasta_salvar:
        caminho_pasta_salvar_var.set(pasta_salvar)

# Criando a janela principal
root = tk.Tk()
root.title("Download eSocial")
root.geometry("500x300")
root.resizable(False, False)

# Labels
labels = [
    "Caminho da planilha:",
    "Linha inicial da planilha:",
    "Linha final da planilha:",
    "Meses a buscar:", 
    "Solicitar (1) / Baixar (2):",   
    "Salvar arquivos:",
    "Data Inicial (DD/MM/AAAA):",
    "Certificado próprio:"
]

# Variáveis para armazenar valores
caminho_planilha_var = tk.StringVar()
caminho_pasta_salvar_var = tk.StringVar()
certificado_proprio_var = tk.BooleanVar()
linha_ini = tk.IntVar()
linha_fim = tk.IntVar()
data_ini = tk.StringVar()

# Posicionamento dos labels e entradas
for i, label_text in enumerate(labels):
    label = tk.Label(root, text=label_text)
    label.grid(row=i, column=0, padx=5, pady=5, sticky="w")
    
    if i == 0:
        entry = tk.Entry(root, textvariable=caminho_planilha_var, width=40)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        button = tk.Button(root, text="Selecionar", command=selecionar_arquivo)
        button.grid(row=i, column=2, padx=5, pady=5)
    elif i == 1:
        entry = tk.Entry(root, textvariable=linha_ini)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    elif i == 2:
        entry = tk.Entry(root, textvariable=linha_fim)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew") 
    elif i == 5:
        entry = tk.Entry(root, textvariable=caminho_pasta_salvar_var)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        button = tk.Button(root, text="Selecionar", command=selecionar_pasta_salvar)
        button.grid(row=i, column=2, padx=5, pady=5)
    elif i == 3:
        meses_buscar_var = tk.IntVar()
        dropdown = ttk.Combobox(root, values=[1, 2, 3, 4, 5, 6], textvariable=meses_buscar_var, state="readonly")
        dropdown.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    elif i == 4:
        solicitar_baixar_var = tk.IntVar()
        dropdown = ttk.Combobox(root, values=[1, 2], textvariable=solicitar_baixar_var, state="readonly")
        dropdown.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    elif i == 6:
        entry = tk.Entry(root, textvariable=data_ini)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")     
    elif i == 7:
        checkbutton = tk.Checkbutton(root, variable=certificado_proprio_var)
        checkbutton.grid(row=i, column=1, padx=5, pady=5, sticky="w")
    else:
        entry = tk.Entry(root)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")

# Botão
button = tk.Button(root, text="Iniciar", command=solicitar_ou_baixar, width=10)
button2 = tk.Button(root, text="Cancelar", command=root.quit, width=10)
button.grid(row=len(labels), column=0, pady=5, padx=5)
button2.grid(row=len(labels), column=1, pady=5, padx=5)

root.mainloop()