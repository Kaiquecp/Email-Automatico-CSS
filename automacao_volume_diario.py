import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import win32com.client as win32
import matplotlib.pyplot as plt

server = 'server'
database = 'database'
username = 'user'
password = 'password'
conn_str = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=SQL+Server'
engine = create_engine(conn_str)

file_path = r'C:\Users\user\Documents\SQL Server Management Studio\SELECT\VOLUME CARREGADO CETREL.sql'
with open(file_path, 'r') as file:
    query = file.read()

df = pd.read_sql(query, engine)
df['DATA_INCLUSAO'] = pd.to_datetime(df['DATA_INCLUSAO']).dt.date
agora = datetime.now()
hoje = agora.date()

if agora.hour > 0 and agora.hour < 10:
    data_referencia = hoje - timedelta(days=1)
elif agora.hour > 12 and agora.hour < 23:
    data_referencia = hoje 
else:
    print("Fora do horário de envio de e-mails. Nenhum e-mail será enviado.")
    exit()

data_referencia_formatada = data_referencia.strftime('%d/%m/%Y')
df_diario = df[df['DATA_INCLUSAO'] == data_referencia]
tipos_frota = {4: 'Frota', 13: 'Terceiros', 14: 'Agregados'}
df_diario['FROTA OU AGREGADO'] = df_diario['FROTA OU AGREGADO'].map(tipos_frota)
df['FROTA OU AGREGADO'] = df['FROTA OU AGREGADO'].map(tipos_frota)
resultado_diario = df_diario.groupby('FROTA OU AGREGADO').agg(
    {'DOCUMENTO': 'count', 'VOLUME CARREGADO': 'sum'}
).rename(columns={'DOCUMENTO': 'QUANTIDADE VIAGENS', 'VOLUME CARREGADO': 'SOMA VOLUME CARREGADO'}).reset_index()
resultado_diario['SOMA VOLUME CARREGADO'] = resultado_diario['SOMA VOLUME CARREGADO'] / 1000

categorias = ['Frota', 'Terceiros', 'Agregados']
resultado_diario['FROTA OU AGREGADO'] = pd.Categorical(
    resultado_diario['FROTA OU AGREGADO'], categories=categorias, ordered=True
)
resultado_diario = resultado_diario.sort_values('FROTA OU AGREGADO').reset_index(drop=True)

total_diario = pd.DataFrame({
    'FROTA OU AGREGADO': ['Total'],
    'QUANTIDADE VIAGENS': [resultado_diario['QUANTIDADE VIAGENS'].sum()],
    'SOMA VOLUME CARREGADO': [resultado_diario['SOMA VOLUME CARREGADO'].sum()]
})
resultado_diario = pd.concat([resultado_diario, total_diario], ignore_index=True)

resultado_acumulado = df.groupby('FROTA OU AGREGADO').agg(
    {'DOCUMENTO': 'count', 'VOLUME CARREGADO': 'sum'}
).rename(columns={'DOCUMENTO': 'QUANTIDADE VIAGENS', 'VOLUME CARREGADO': 'SOMA VOLUME CARREGADO'}).reset_index()

resultado_acumulado['SOMA VOLUME CARREGADO'] = resultado_acumulado['SOMA VOLUME CARREGADO'] / 1000

categorias = ['Frota', 'Terceiros', 'Agregados']
resultado_acumulado['FROTA OU AGREGADO'] = pd.Categorical(
    resultado_acumulado['FROTA OU AGREGADO'], categories=categorias, ordered=True
)
resultado_acumulado = resultado_acumulado.sort_values('FROTA OU AGREGADO').reset_index(drop=True)

total_acumulado = pd.DataFrame({
    'FROTA OU AGREGADO': ['Total'],
    'QUANTIDADE VIAGENS': [resultado_acumulado['QUANTIDADE VIAGENS'].sum()],
    'SOMA VOLUME CARREGADO': [resultado_acumulado['SOMA VOLUME CARREGADO'].sum()]
})
resultado_acumulado = pd.concat([resultado_acumulado, total_acumulado], ignore_index=True)

resultado_diario['PERCENTUAL'] = (
    resultado_diario['SOMA VOLUME CARREGADO'].astype(float) /
    resultado_diario['SOMA VOLUME CARREGADO'].astype(float).sum()
) * 100

resultado_diario_grafico = resultado_diario[resultado_diario['FROTA OU AGREGADO'] != 'Total']

resultado_diario_grafico['PERCENTUAL'] = (
    resultado_diario_grafico['SOMA VOLUME CARREGADO'].astype(float) /
    resultado_diario_grafico['SOMA VOLUME CARREGADO'].astype(float).sum()
) * 100

fig, ax = plt.subplots(figsize=(8, 6))
ax.pie(
    resultado_diario_grafico['PERCENTUAL'],
    labels=resultado_diario_grafico['FROTA OU AGREGADO'],
    autopct='%1.1f%%',
    startangle=90,
    colors=['#4CAF50', '#FFC107', '#2196F3']
)
fig, ax = plt.subplots(figsize=(8, 6))
ax.pie(
    resultado_diario_grafico['PERCENTUAL'],
    labels=resultado_diario_grafico['FROTA OU AGREGADO'],
    autopct='%1.1f%%',
    startangle=90,
    colors=['#4CAF50', '#FFC107', '#2196F3']
)
ax.set_title(f'Percentual do Volume Carregado - {data_referencia_formatada}', fontsize=14)
plt.tight_layout()

grafico_path = r'C:\Users\user\Documents\Resumo_Volume_Carregado.png'
plt.savefig(grafico_path)
plt.close()

for col in ['SOMA VOLUME CARREGADO']:
    resultado_diario[col] = resultado_diario[col].map(lambda x: f"{x:,.2f}".replace(",", "."))
    resultado_acumulado[col] = resultado_acumulado[col].map(lambda x: f"{x:,.2f}".replace(",", "."))

resultado_diario_html = resultado_diario.drop(columns=['PERCENTUAL'], errors='ignore')
html_diario = resultado_diario_html.to_html(index=False, classes='table table-bordered')
html_acumulado = resultado_acumulado.to_html(index=False, classes='table table-bordered')

body = f"""
<p>Prezados,</p>
<p>Segue o resumo do volume carregado:</p>

<h3>Resumo Diário do Volume Carregado em Toneladas de {data_referencia_formatada}</h3>
{html_diario}

<h3>Resumo Acumulado Total do Volume Carregado em Toneladas</h3>
{html_acumulado}

<p><strong>Gráfico do Percentual de Volume Carregado:</strong></p>
<img src="cid:GraficoVolume" alt="Gráfico do Percentual de Volume Carregado">
"""

def send_email_with_chart(subject, body, to_email, chart_path):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.BodyFormat = 2
    mail.HTMLBody = body
    mail.To = to_email
    mail.Attachments.Add(chart_path)
    for attachment in mail.Attachments:
        if attachment.FileName == chart_path.split("\\")[-1]:
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "GraficoVolume"
            )

    try:
        mail.Send()
        print(f"E-mail enviado com sucesso para {to_email}")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")

send_email_with_chart(
    "Resumo Diário e Acumulado do Volume Carregado",
    body,
    'email',
    grafico_path
)
