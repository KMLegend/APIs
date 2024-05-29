import requests
import pandas as pd

def extrair_dados_indicadores():
    # url_base = f"https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais"
    url_base = f"https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?$top=100&$format=json"
    resposta = requests.get(url_base)

    if resposta.status_code == 200:
        return resposta.json()
    else:
        print("Erro na requisição. Status code:", resposta.status_code, resposta.text)
        return None
    
    
dados_expectativas = extrair_dados_indicadores()



# print(dados_expectativas['value'])


df = pd.DataFrame(dados_expectativas['value'])


caminho_destino = "\\\\192.168.100.3/dados city\\Inovação e Sistemas\\01-INOVACAO\\02-DESENVOLVIMENTO\\01-HOMOLOGACAO\\APIs\\Indicadores-Financeiro\\sla.xlsx"

df.to_excel(caminho_destino)
print(df)