from config import config
from shareplum import Site
from shareplum.site import Version
from shareplum import Office365
import pandas as pd


# Detalhes de Autenticação
site_url = 'https://cityinc123.sharepoint.com/juridico/'
usuario = config['sp_user']
senha = config['sp_password']
nome_lista = 'Histórico SPES'

# Configurar a conexão com o site SharePoint
authcookie = Office365('https://cityinc123.sharepoint.com', username=usuario, password=senha).GetCookies()
site = Site('https://cityinc123.sharepoint.com/juridico/',version=Version.v365, authcookie=authcookie)


df = pd.DataFrame(site.List(nome_lista).GetListItems()) # Gera o Dataframe requisitando os dados da lista passada.

df.to_excel("C:\\Users\\kevin.maykel\\OneDrive - CITY INCORP LTDA\\Documentos\\Meus Projetos\\projeto-city\\sharepoint\\teste.xlsx")
print(df)
