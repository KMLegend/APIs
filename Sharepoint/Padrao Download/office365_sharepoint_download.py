from config import config
from shareplum import Site
from shareplum.site import Version
from shareplum import Office365
import os
import pandas as pd

site_url = 'https://cityinc123.sharepoint.com/inovacao-sistemas'
usuario = config['sp_user']
senha = config['sp_password']
nome_arquivo = ''
download_caminho = 'full_path'
doc = 'Documentos Compartilhados/Forms - KM'

# Configurar a conexão com o site SharePoint
authcookie = Office365('https://cityinc123.sharepoint.com', username=usuario, password=senha).GetCookies()
site = Site('https://cityinc123.sharepoint.com/inovacao-sistemas/',version=Version.v365, authcookie=authcookie)

pasta = site.Folder(doc)

arquivo = pasta.get_file(nome_arquivo)

# Escrever o conteúdo do arquivo em um arquivo local
with open(download_caminho, 'wb') as f:
    f.write(arquivo)