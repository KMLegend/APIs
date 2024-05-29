from credenciais import login
from shareplum import Site
from shareplum.site import Version
from shareplum import Office365
from bs4 import BeautifulSoup
import pandas as pd
import pyodbc

dados_conexao = (
    "Driver={SQL Server};"
    "Server=srv-sql01;"
    "Database=BI;"
    "UID=usruau;"
    "PWD=spfmed4$;"
)
conn = pyodbc.connect(dados_conexao)
cur = conn.cursor()

nome_site = 'Juridico'
site_url = f'https://cityinc123.sharepoint.com/{nome_site}'
usuario = login['usuario']
senha = login['senha']

# Login e conexão com o site SharePoint
authcookie = Office365('https://cityinc123.sharepoint.com', username=usuario, password=senha).GetCookies()
site = Site(f'https://cityinc123.sharepoint.com/{nome_site}/', version=Version.v365, authcookie=authcookie)

# Obter e listar os nomes das listas do site
listas = site.GetListCollection()
writer = pd.ExcelWriter("C:\\Users\\kevin.maykel\\OneDrive - CITY INCORP LTDA\\Documentos\\Meus Projetos\\projeto-city\\sharepoint\\Juridico\\final.xlsx", engine='xlsxwriter')

def limpa_html(html):
    if pd.isna(html):  # Verificar se o valor é NaN ou None
        return ''
    limpeza = BeautifulSoup(html, 'html.parser')
    return limpeza.get_text()

nome_lista = "Histórico SPES"
df = pd.DataFrame(site.List(nome_lista).GetListItems())

# df['Unidades Atribuídas'] = df['Unidades Atribuídas'].apply(limpa_html)

df = df[['Proprietário','Unidades Atribuídas']]


lista_valores = []
lista_nomes = []

for idx, row in df.iterrows():
    nomes = row['Proprietário']
    valores = row['Unidades Atribuídas'].split(';')
    for valor in valores:
        valor_pronto = valor.split('-')
        lista_valores.append(valor_pronto)
        lista_nomes.append(nomes)

df_valores = pd.DataFrame(lista_valores)
df_valores['Propietário'] = lista_nomes

df_valores.columns = ['Empreendimento', 'Unidade', 'Proprietário']
print(df_valores)

for index, row in df_valores.iterrows():
    try:
        sql = f"INSERT INTO [dbo].[TB_SYS_SHAREPOINT_JUR] values ('{row.Empreendimento}', '{row.Unidade}', '{row.Proprietário}')"
        # print(sql)
        cur.execute(sql)
        conn.commit()
    except Exception as e:
        ...

df_valores.to_excel(writer, sheet_name=nome_lista)





















































































































writer.close()