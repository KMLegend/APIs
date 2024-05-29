import requests
from credenciais import login

id_app = ''
credencial_segredo = ''

# ID do diretorio microsoft, dentro do microsoft entra
id_locatario_entra = ''

# Autenticação
token_url = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token'.format(id_locatario_entra)
token_data = {
    'grant_type': 'client_credentials',
    'client_id': id_app,
    'client_secret': credencial_segredo,
    'scope': 'https://graph.microsoft.com/.default'
}
token_r = requests.post(token_url, data=token_data)

if token_r.status_code!= 200:
    print("Erro na autenticação:", token_r.status_code, token_r.text)
    exit()

token = token_r.json().get('access_token')


# Obtenção dos dados do formulário
form_id = ''
# form_id = ''


form_url = f'https://forms.office.com/formapi/DownloadExcelFile.ashx?formid={form_id}&timezoneOffset=180&minResponseId=1&maxResponseId=1000'
headers = {
    'Authorization': f'Bearer {token}',
    'Content-Type': 'application/json'
}

response = requests.get(form_url, headers=headers)

# Processamento dos dadosd
if response.status_code == 200:
    data = response.json()
    print(data)
else:
    print("Erro ao obter dados do formulário:", response.status_code, response.text)