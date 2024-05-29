
from credenciais import login
from shareplum import Site
from shareplum.site import Version
from shareplum import Office365
import os
import pandas as pd
import pyodbc

def identifica_lista():
    armazena_df, nomes = pega_dfs()

    # Iterando sobre os nomes e dataframes armazenados
    for i, nome in enumerate(nomes):
        df = armazena_df[i]
        if 'Lista' not in df.columns:
            df.insert(0, 'Lista', nome)
        else:
            df['Lista'] = nome
        # Atualizando o dataframe no dicionário
        armazena_df[i] = df

    # Consolidando todos os dataframes
    df_consolidado = armazena_df[:]
    return df_consolidado

def pega_dfs():
    armazena_df = []
    nomes = []

    # Configurar a conexão com o site SharePoint
    authcookie = Office365('https://cityinc123.sharepoint.com', username=usuario, password=senha).GetCookies()
    site = Site(f'https://cityinc123.sharepoint.com/{nome_site}/',version=Version.v365, authcookie=authcookie)
    listas = site.GetListCollection()
    # Extrair e imprimir os nomes das listas
    nomes_listas = [lista['Title'] for lista in listas]
    for nome_lista in nomes_listas:
        if nome_lista[0:15] == "Controle MKT - ":
            df = pd.DataFrame(site.List(nome_lista).GetListItems())
            nomes.append(nome_lista)
            armazena_df.append(df)
    return armazena_df, nomes


def insere_bd():
    dados_conexao = (
    "Driver={SQL Server};"
    "Server=srv-sql01;"
    "Database=BI;"
    "UID=;"
    "PWD=;"
    )

    conn = pyodbc.connect(dados_conexao)
    cur = conn.cursor()
    df_consolidado = identifica_lista()
    df = pd.concat(df_consolidado)
    
    # Selecionando as colunas necessárias
    df = df[['Lista', 'Numero JOB', 'DATA CONTRATAÇÃO', 'DATA VENCIMENTO', 'FORNECEDOR', 'NF', 'VALOR', 'DESCRIÇÃO', 'INSUMO', 'ID', 'Modificado', 'Criado', 'OBSERVAÇÕES', 'DATA DE REPROGRAMAÇÃO PGTO', 'FORMA DE PAGAMENTO', 'Status', 'Número Processo UAU', 'Criado por']]

    df.fillna('NULL', inplace=True)
    
    for index, row in df.iterrows():
        try:
            sql = f"INSERT INTO [dbo].[TB_SYS_API_SHAREPOINT_MKT] VALUES ('{row['Lista']}', '{row['Numero JOB']}', '{row['DATA CONTRATAÇÃO']}', '{row['DATA VENCIMENTO']}', '{row['FORNECEDOR'].replace("'", "")}', '{row['NF']}', '{row['VALOR']}', '{row['DESCRIÇÃO'].replace("'", "")}', '{row['INSUMO'].replace("'", "")}', '{row['ID']}', '{row['Modificado']}', '{row['Criado']}', '{row['OBSERVAÇÕES'].replace("'", "")}', '{row['DATA DE REPROGRAMAÇÃO PGTO']}', '{row['FORMA DE PAGAMENTO']}', '{row['Status'].replace("'", "")}', '{row['Número Processo UAU']}', '{row['Criado por']}')"
            
            cur.execute(sql)
            conn.commit()
            
        except Exception as e:
            print(f"Erro ao inserir a linha {index}: {e}")
            print(f"Dados problemáticos: {row}")
            conn.rollback()
    
    cur.close()
    conn.close()
    # df.to_excel(caminho_destino)

nome_site = 'Marketing'
site_url = f'https://cityinc123.sharepoint.com/{nome_site}'
usuario = login['usuario']
senha = login['senha']
# caminho_destino = "\\\\192.168.100.3\\dados city\\Inovação e Sistemas\\01-INOVACAO\\02-DESENVOLVIMENTO\\04-ARQUIVOS\\Controle MKT - Listas.xlsx"

insere_bd()