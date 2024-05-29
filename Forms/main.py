from credenciais import login
from shareplum import Site
from shareplum.site import Version
from shareplum import Office365
import os
import time
import pandas as pd
from datetime import datetime

def upload_arquivo():
    authcookie = Office365('https://cityinc123.sharepoint.com', username=usuario, password=senha).GetCookies()
    site = Site('https://cityinc123.sharepoint.com/inovacao-sistemas/', version=Version.v365, authcookie=authcookie)

    folder = site.Folder(caminho_destino)

    with open(caminho_origem, 'rb') as f:
        file_content = f.read()
        folder.upload_file(file_content, nome_arquivo)


    data_info = datetime.now()
    data_info = data_info.strftime("%d-%m-%Y %H:%M:%S")
    print(f"Arquivo carregado para o SharePoint: {nome_arquivo} na Data: {data_info}")
    
def converte_hora():    
    # Converter a coluna para datetime
    df['Hora de conclusão'] = pd.to_datetime(df['Hora de conclusão'])

    # Verificar se a coluna já tem um fuso horário
    if df['Hora de conclusão'].dt.tz is None:
        # Localizar o fuso horário atual (exemplo: 'UTC')
        df['Hora de conclusão'] = df['Hora de conclusão'].dt.tz_localize('UTC')

    # Converter para um novo fuso horário (exemplo: 'America/Sao_Paulo')
    df['Hora de conclusão'] = df['Hora de conclusão'].dt.tz_convert('America/Sao_Paulo')
    df['Hora de conclusão'] = df['Hora de conclusão'].dt.tz_localize(None)
    
    return df
    # print(df)

def insere_qual_forms():
    armazena_df, nomes = pega_dfs()
    for i, nome in enumerate(nomes):
        df = armazena_df[i]
        if 'Forms' not in df.columns:
            df.insert(0, 'Forms', nome)
        else:
            df['Forms'] = nome
        # Atualizando o dataframe no dicionário
        armazena_df[i] = df

    # Consolidando todos os dataframes
    df_consolidado = armazena_df[:]
    return df_consolidado

def pega_dfs():
    armazena_df = []
    nomes = []
    caminho = "\\\\192.168.100.3\\dados city\\Inovação e Sistemas\\01-INOVACAO\\02-DESENVOLVIMENTO\\01-HOMOLOGACAO\\APIs\\Forms\\Download"
    nome_arquivos = os.listdir(caminho)
    for nome_arquivo in nome_arquivos:
        caminho_arquivo = os.path.join(caminho, nome_arquivo)
        if caminho_arquivo.endswith(('.xlsx', '.xls')):
            try:
                full_path = os.path.join(caminho, nome_arquivo)
                #  print(nome_arquivo)
            except Exception as e:
                print(f"Erro ao acessar arquivo: {full_path}")
                print("Erro:", e)
            download_caminho = full_path
            # Configurar a conexão com o site SharePoint
            authcookie = Office365('https://cityinc123.sharepoint.com', username=usuario, password=senha).GetCookies()
            site = Site('https://cityinc123.sharepoint.com/inovacao-sistemas/',version=Version.v365, authcookie=authcookie)

            pasta = site.Folder(doc)

            arquivo = pasta.get_file(nome_arquivo)

            # Escrever o conteúdo do arquivo em um arquivo local
            with open(download_caminho, 'wb') as f:
                f.write(arquivo)
            
            nome_arquivo = nome_arquivo.split("Book de Metas - ")
            nome_arquivo = nome_arquivo[1][:-5]
            # print(nome_arquivo)
            nomes.append(nome_arquivo)
            armazena_df.append(pd.read_excel(full_path))
    return armazena_df, nomes

while True:
    
    site_url = 'https://cityinc123.sharepoint.com/inovacao-sistemas'
    usuario = login['usuario']
    senha = login['senha']

    nome_arquivo = 'Final_Forms.xlsx'
    doc = 'Documentos Compartilhados/Forms - KM'
    caminho_destino = 'Documentos Compartilhados/Compartilhado'
    caminho_origem = "\\\\192.168.100.3\\dados city\\Inovação e Sistemas\\01-INOVACAO\\02-DESENVOLVIMENTO\\01-HOMOLOGACAO\\APIs\\Forms\\Final_Forms.xlsx"

    df_forms = insere_qual_forms()
    df = pd.concat(df_forms)
    df = converte_hora()
    df.to_excel(caminho_origem)

    upload_arquivo()
    
    time.sleep(24 * 60 * 60)
    # time.sleep(30)
    
