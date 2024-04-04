import requests
from datetime import datetime

def data_abertura(cnpj):
    cnpj = input('Digite o CNPJ: ')    
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
    data_inicial = datetime(2018, 1, 1)

    if abertura:
        try:
            # Converte a data de abertura para um objeto datetime
            abertura_dt = datetime.strptime(abertura, '%d/%m/%Y')
            print("Data de abertura original:", abertura_dt)

            # Substitui o dia pelo dia 01
            abertura_dt = abertura_dt.replace(day=1)
            print("Data de inicio utilizada:", abertura_dt)

            # Defina start_date para abertura_dt se for posterior a 2018-01-01, caso contrário, use 2018-01-01
            data_inicial = max(abertura_dt, datetime(2018, 1, 1))
            
        except ValueError:
            print("Erro ao converter a data de abertura.")
            data_inicial = datetime(2018, 1, 1)
    else:
        data_inicial = datetime(2018, 1, 1)
    
    return abertura

data_abertura(cnpj=0)    