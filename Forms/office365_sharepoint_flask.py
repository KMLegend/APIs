from config import config
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
    armazena_df = pega_dfs()
    captacao_Financeira = armazena_df[0]
    contabilidade = armazena_df[1]
    gestao_e_processos = armazena_df[2]
    legalizacao = armazena_df[3]
    novos_negocios = armazena_df[4]
    obras_comerciais = armazena_df[5]
    orcamento = armazena_df[6]
    pep = armazena_df[7]
    personalizacao = armazena_df[8]
    planejamento_e_controle = armazena_df[9]
    projetos_comerciais = armazena_df[10]
    projetos_executivos = armazena_df[11]
    projetos = armazena_df[12]
    secretaria_de_vendas = armazena_df[13]
    seguranca_de_trabalho = armazena_df[14]
    sgq =  armazena_df[15]
    suprimentos = armazena_df[16]
    
    if 'Forms' not in captacao_Financeira.columns:
        captacao_Financeira.insert(0, 'Forms', 'captacao_Financeira')
    else:
        captacao_Financeira['Forms'] = 'captacao_Financeira'

    if 'Forms' not in contabilidade.columns:
        contabilidade.insert(0, 'Forms', 'contabilidade')
    else:
        contabilidade['Forms'] = 'contabilidade'

    if 'Forms' not in gestao_e_processos.columns:
        gestao_e_processos.insert(0, 'Forms', 'gestao_e_processos')
    else:
        gestao_e_processos['Forms'] = 'gestao_e_processos'

    if 'Forms' not in legalizacao.columns:
        legalizacao.insert(0, 'Forms', 'legalizacao')
    else:
        legalizacao['Forms'] = 'legalizacao'

    if 'Forms' not in novos_negocios.columns:
        novos_negocios.insert(0, 'Forms', 'novos_negocios')
    else:
        novos_negocios['Forms'] = 'novos_negocios'

    if 'Forms' not in obras_comerciais.columns:
        obras_comerciais.insert(0, 'Forms', 'obras_comerciais')
    else:
        obras_comerciais['Forms'] = 'obras_comerciais'

    if 'Forms' not in orcamento.columns:
        orcamento.insert(0, 'Forms', 'orcamento')
    else:
        orcamento['Forms'] = 'orcamento'

    if 'Forms' not in pep.columns:
        pep.insert(0, 'Forms', 'pep')
    else:
        pep['Forms'] = 'pep'
        
    if 'Forms' not in personalizacao.columns:
        personalizacao.insert(0, 'Forms', 'personalizacao')
    else:
        personalizacao['Forms'] = 'personalizacao'
        
    if 'Forms' not in planejamento_e_controle.columns:
        planejamento_e_controle.insert(0, 'Forms', 'planejamento_e_controle')
    else:
        planejamento_e_controle['Forms'] = 'planejamento_e_controle'

    if 'Forms' not in projetos_comerciais.columns:
        projetos_comerciais.insert(0, 'Forms', 'projetos_comerciais')
    else:
        projetos_comerciais['Forms'] = 'projetos_comerciais'

    if 'Forms' not in projetos_executivos.columns:
        projetos_executivos.insert(0, 'Forms', 'projetos_executivos')
    else:
        projetos_executivos['Forms'] = 'projetos_executivos'
        
    if 'Forms' not in projetos.columns:
        projetos.insert(0, 'Forms', 'projetos')
    else:
        projetos['Forms'] = 'projetos'
        
    if 'Forms' not in secretaria_de_vendas.columns:
        secretaria_de_vendas.insert(0, 'Forms', 'secretaria_de_vendas')
    else:
        secretaria_de_vendas['Forms'] = 'secretaria_de_vendas'

    if 'Forms' not in seguranca_de_trabalho.columns:
        seguranca_de_trabalho.insert(0, 'Forms', 'seguranca_de_trabalho')
    else:
        seguranca_de_trabalho['Forms'] = 'seguranca_de_trabalho'
        
    if 'Forms' not in sgq.columns:
        sgq.insert(0, 'Forms', 'sgq')
    else:
        sgq['Forms'] = 'sgq'

    if 'Forms' not in suprimentos.columns:
        suprimentos.insert(0, 'Forms', 'suprimentos')
    else:
        suprimentos['Forms'] = 'suprimentos'

    df_consolidado = [captacao_Financeira, contabilidade, gestao_e_processos, legalizacao, novos_negocios, obras_comerciais, orcamento, pep, personalizacao, planejamento_e_controle, projetos_comerciais, projetos_executivos, projetos, secretaria_de_vendas, seguranca_de_trabalho, sgq, suprimentos]
    return df_consolidado

def pega_dfs():
    armazena_df = []
    caminho = "C:\\Users\\kevin.maykel\\OneDrive - CITY INCORP LTDA\\Documentos\\Meus Projetos\\projeto-city\\sharepoint\\Python\\storage"
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
            armazena_df.append(pd.read_excel(full_path))
    return armazena_df

site_url = 'https://cityinc123.sharepoint.com/inovacao-sistemas'
usuario = config['sp_user']
senha = config['sp_password']

nome_arquivo = 'Final_Forms.xlsx'
doc = 'Documentos Compartilhados/Forms - KM'
caminho_destino = 'Documentos Compartilhados/Compartilhado'
caminho_origem = "C:\\Users\\kevin.maykel\\OneDrive - CITY INCORP LTDA\\Documentos\\Meus Projetos\\projeto-city\\sharepoint\\Python\\Final_Forms.xlsx"

df_forms = insere_qual_forms()
df = pd.concat(df_forms)
df = converte_hora()
df.to_excel(caminho_origem)

upload_arquivo()