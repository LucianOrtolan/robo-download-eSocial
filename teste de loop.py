from datetime import datetime, timedelta

# Lista de empresas
empresas = ["Empresa A", "Empresa B"]  # Adicione suas empresas aqui

# Data inicial
data_inicial = datetime(2018, 1, 1)

# Data de corte
data_corte = datetime(2024, 4, 4)

# Dicionário para armazenar a última data final de cada empresa
ultima_data_final_por_empresa = {}

# Loop principal para percorrer as datas
data_atual = data_inicial
while data_atual <= data_corte:
    for empresa in empresas:
        # Obtém a última data final para a empresa
        if empresa not in ultima_data_final_por_empresa:
            ultima_data_final_por_empresa[empresa] = data_atual
        data_final = ultima_data_final_por_empresa[empresa] + timedelta(days=30)

        # Fazer cinco requisições para a empresa atual
        for i in range(5):
            print(f"Fazendo requisição para {empresa} de {ultima_data_final_por_empresa[empresa]} até {data_final}")
            ultima_data_final_por_empresa[empresa] = data_final + timedelta(days=1)
            data_final = ultima_data_final_por_empresa[empresa] + timedelta(days=30)
            if data_final > data_corte:
                data_final = data_corte
                break

    # Atualiza a data atual para a próxima iteração
    data_atual = ultima_data_final_por_empresa[empresas[0]] + timedelta(days=1)
