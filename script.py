import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle

# Obtém o diretório atual onde o script está sendo executado
diretorio_atual = os.getcwd()

# Lista todos os arquivos no diretório atual
arquivos = os.listdir(diretorio_atual)

# Filtra arquivos com extensões .xls, .xlsx ou .csv
arquivos_validos = [arquivo for arquivo in arquivos if arquivo.endswith(('.xls', '.xlsx', '.csv'))]

# Verifica se há arquivos compatíveis no diretório
if arquivos_validos:
    # Pega o primeiro arquivo encontrado
    arquivo_original = arquivos_validos[0]
    
    # Caminho completo do arquivo original
    caminho_arquivo_original = os.path.join(diretorio_atual, arquivo_original)
    
    # Carrega o arquivo de acordo com a extensão
    if arquivo_original.endswith('.csv'):
        df = pd.read_csv(caminho_arquivo_original)
    else:  # Para .xls e .xlsx
        df = pd.read_excel(caminho_arquivo_original)

    # Cria o DataFrame modificado
    df_modificado = pd.DataFrame()

    # Preenche as novas colunas conforme solicitado
    df_modificado['ID'] = df['Matrícula']
    df_modificado['Nome próprio'] = df['Nome'].apply(lambda x: x.split()[0] if pd.notnull(x) else '')
    df_modificado['Apelido'] = df['Nome'].apply(lambda x: ' '.join(x.split()[1:]) if pd.notnull(x) and len(x.split()) > 1 else '')
    df_modificado['Departamento'] = 'All Departments/6 - CNAT_ALUNOS_2024'
    data_atual = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
    df_modificado['Hora de início do período de vigência'] = data_atual
    data_fim = (datetime.now() + timedelta(days=3650)).strftime('%Y/%m/%d %H:%M:%S')  # 10 anos depois
    df_modificado['Hora do fim do período de vigência'] = data_fim
    df_modificado['E-mail'] = df['E-mail Pessoal']
    df_modificado['Matricula'] = ''

    # Define o novo nome do arquivo com "(1) Modificado" adicionado e extensão .xlsx
    nome, _ = os.path.splitext(arquivo_original)
    novo_nome = f"{nome} Modificado.xlsx"
    caminho_novo_arquivo = os.path.join(diretorio_atual, novo_nome)

    # Salva o DataFrame modificado em um novo arquivo Excel .xlsx
    df_modificado.to_excel(caminho_novo_arquivo, index=False, engine='openpyxl')

    # Carrega o arquivo salvo para aplicar formatação
    workbook = load_workbook(caminho_novo_arquivo)
    sheet = workbook.active

    # Define um estilo de número para a coluna 'ID'
    num_style = NamedStyle(name="num_style")
    num_style.number_format = '0'

    # Define um estilo de data personalizada
    date_style = NamedStyle(name="date_style")
    date_style.number_format = 'yyyy/mm/dd hh:mm:ss'

    # Aplica a formatação de número à coluna 'ID'
    for cell in sheet['A']:  # Assumindo que 'ID' é a primeira coluna
        cell.style = num_style

    # Aplica a formatação de data às colunas de datas
    for col in ['E', 'F']:  # Colunas 'Hora de início' e 'Hora do fim'
        for cell in sheet[col]:
            cell.style = date_style

    # Salva o workbook com as formatações aplicadas
    workbook.save(caminho_novo_arquivo)

    print(f'Arquivo modificado salvo como {novo_nome}')
else:
    print('Nenhum arquivo .xls, .xlsx ou .csv encontrado no diretório atual.')
