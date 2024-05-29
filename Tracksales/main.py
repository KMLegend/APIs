import pandas as pd
import requests
from datetime import date, datetime
from pytz import timezone
import time
import pyodbc

dados_conexao = (
    "Driver={SQL Server};"
    "Server=srv-sql01;"
    "Database=BI;"
    "UID=usruau;"
    "PWD=spfmed4$;"
)



while True:
    conn = pyodbc.connect(dados_conexao)
    cur = conn.cursor()
    def extrair_dados_tracksales_campanha(token_acesso):
        data = date.today()
        # print(data)
        url_base = f"https://api.tracksale.co/v2/campaign?start=2024-01-01&end={data}"
        headers = {"Authorization": "Bearer " + token_acesso}
        resposta = requests.get(url_base, headers=headers)

        if resposta.status_code == 200:
            return resposta.json()
        else:
            print("Erro na requisição. Status code:", resposta.status_code, resposta.text)
            return None

    def extrair_dados_tracksales_report(token_acesso):
        url_base = f"https://api.tracksale.co/v2/report/answer"
        headers = {"Authorization": "Bearer " + token_acesso}
        resposta = requests.get(url_base, headers=headers)

        if resposta.status_code == 200:
            return resposta.json()
        else:
            print("Erro na requisição. Status code:", resposta.status_code, resposta.text)
            return None

    
    # Exemplo de uso
    token_acesso = "9276526c26094e58e0f307b722bf4625"
    dados_campanha = extrair_dados_tracksales_campanha(token_acesso)
    dados_report = extrair_dados_tracksales_report(token_acesso)

    data_convertida_campanha = []
    data_convertida_report = []

    if dados_campanha is not None:
        df_campanha = pd.DataFrame(dados_campanha)
        df_report = pd.DataFrame(dados_report)
        
        for i in df_campanha["create_time"]:
            timestamp_campanha = i
            # print(timestamp)
            dt = datetime.fromtimestamp(timestamp_campanha, tz = timezone("America/Sao_Paulo"))
            dt = dt.strftime("%Y-%m-%d %H:%M:%S")
            data_convertida_campanha.append(dt)
            # print(dt)
            
        df_campanha["create_time"] = data_convertida_campanha
        
        for i in df_report['time']:
            timestamp_report = i
            # print(timestamp)
            dt = datetime.fromtimestamp(timestamp_report, tz = timezone("America/Sao_Paulo"))
            dt = dt.strftime("%Y-%m-%d %H:%M:%S")
            data_convertida_report.append(dt)
            
            # print(dt)
            
        df_report['time'] = data_convertida_report
        
        df_report = df_report[['time', 'type', 'name', 'email', 'identification', 'phone', 'alternative_email', 'alternative_phone', 'nps_answer', 'last_nps_answer', 'nps_comment', 'campaign_name', 'campaign_code', 'id', 'deadline', 'elapsed_time', 'dispatch_time', 'reminder_time', 'status', 'priority', 'assignee', 'lot_code', 'picture', 'tags', 'categories', 'justifications']]
        df_campanha = df_campanha[['name', 'code', 'description', 'detractors', 'passives', 'promoters', 'dispatches', 'comments', 'answers', 'main_channel', 'create_time']]
        
        df_campanha = df_campanha.drop_duplicates()
        
        # print(df_campanha)
        # print(df_report.dtypes)
        
        for index, row in df_campanha.iterrows():
            sla = row['name']            
            try:
                cur.execute(f"INSERT INTO [dbo].[TB_SYS_API_TRACKSALES_CLIENTE] values ({index},'{sla}',{row.code},'{row.description}',{row.detractors},{row.passives},{row.promoters},{row.dispatches},{row.comments},{row.answers},'{row.main_channel}','{row.create_time}')")
                conn.commit()
                # cur.execute(f"INSERT INTO [dbo].[TB_SYS_API_TRACKSALES_CLIENTE] values ({index},'{row['name']}',{row.code},'{row.description}',{row.detractors},{row.passives},{row.promoters},{row.dispatches},{row.comments},{row.answers},'{row.main_channel}','{row.create_time}')")
            except Exception as e:
                ...
                # print(e)
        
        for index_report, row_report in df_report.iterrows():            
            try:
                sql = f"INSERT INTO [dbo].[TB_SYS_API_TRACKSALES_REPORT] values ({index_report}, '{row_report.time}', '{row_report.type}', '{row_report.name}', '{row_report.email}', '{row_report.identification}', '{row_report.phone}', '{row_report.alternative_email}', '{row_report.alternative_phone}', {row_report.nps_answer}, '{row_report.last_nps_answer}', '{row_report.nps_comment}', '{row_report.campaign_name}', {row_report.campaign_code}, {row_report.id}, '{row_report.deadline}', '{row_report.elapsed_time}', '{row_report.dispatch_time}', '{row_report.reminder_time}', '{row_report.status}', '{row_report.priority}', '{row_report.assignee}', '{row_report.lot_code}', '{row_report.picture}', '{row_report.tags}', '{row_report.categories}', '{row_report.justifications}')"
                # print(sql)
                cur.execute(sql)
                conn.commit()
                # cur.execute(f"INSERT INTO [dbo].[TB_SYS_API_TRACKSALES_CLIENTE] values ({index},'{row['{row_report.name}']}',{row.code},'{row.description}',{row.detractors},{row.passives},{row.promoters},{row.dispatches},{row.comments},{row.answers},'{row.main_channel}','{row.create_time}')")
            except Exception as e:
                ...
                # print(e)

        # df_report.to_excel(caminho_destino_report)
        
    data_info = datetime.now()
    data_info = data_info.strftime("%d-%m-%Y %H:%M:%S")
    time.sleep(24 * 60 * 60)
    # time.sleep(20)
    print("Rodou na data: " + f'{data_info}')
    
    cur.close()
    conn.close()
        


