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

cur.execute("SELECT * FROM [dbo].[TB_SYS_API_TRACKSALES_CLIENTE]")

for row in cur:
    print(row)

cur.close()
conn.close()