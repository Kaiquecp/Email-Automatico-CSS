import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import win32com.client as win32
import matplotlib.pyplot as plt


# Configurações de conexão
server = '54.207.0.223\\DW_TRANSPARANA,14183'
database = 'db_visual_transparana'
username = 'userTransparana'
password = '3Y9>)C(=)md6'

# Criando a engine SQLAlchemy para a conexão
conn_str = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=SQL+Server'
engine = create_engine(conn_str)

# Lendo a consulta SQL do arquivo
file_path = r'C:\Users\Kaique\Documents\SQL Server Management Studio\SELECT\VOLUME CARREGADO CETREL.sql'
with open(file_path, 'r') as file:
    query = file.read()

# Executando a consulta e carregando o resultado em um DataFrame
df = pd.read_sql(query, engine)

# Convertendo a coluna de data para apenas a data
df['DATA_INCLUSAO'] = pd.to_datetime(df['DATA_INCLUSAO']).dt.date

# Verificar o horário atual
agora = datetime.now()
hoje = agora.date()

if agora.hour > 0 and agora.hour < 10:
    data_referencia = hoje - timedelta(days=1)  # Enviar dados do dia anterior
elif agora.hour > 12 and agora.hour < 23:
    data_referencia = hoje  # Enviar dados do dia atual
else:
    print("Fora do horário de envio de e-mails. Nenhum e-mail será enviado.")
    exit()  # Termina o programa se não estiver no horário adequado

# Formatando data referencia
data_referencia_formatada = data_referencia.strftime('%d/%m/%Y')
# Filtro para a data de referência
df_diario = df[df['DATA_INCLUSAO'] == data_referencia]

# Mapeando os tipos de frota no DataFrame filtrado
tipos_frota = {4: 'Frota', 13: 'Terceiros', 14: 'Agregados'}

# Aplicando a mapeamento
df_diario['FROTA OU AGREGADO'] = df_diario['FROTA OU AGREGADO'].map(tipos_frota)

# Aplicando o mapeamento na coluna 'FROTA OU AGREGADO'
df['FROTA OU AGREGADO'] = df['FROTA OU AGREGADO'].map(tipos_frota)

# Resumo Diário: Segregando corretamente por categoria
resultado_diario = df_diario.groupby('FROTA OU AGREGADO').agg(
    {'DOCUMENTO': 'count', 'VOLUME CARREGADO': 'sum'}
).rename(columns={'DOCUMENTO': 'QUANTIDADE VIAGENS', 'VOLUME CARREGADO': 'SOMA VOLUME CARREGADO'}).reset_index()

# Dividindo o volume carregado por 1000 (transformando para toneladas)
resultado_diario['SOMA VOLUME CARREGADO'] = resultado_diario['SOMA VOLUME CARREGADO'] / 1000

# Garantindo a ordem das categorias
categorias = ['Frota', 'Terceiros', 'Agregados']
resultado_diario['FROTA OU AGREGADO'] = pd.Categorical(
    resultado_diario['FROTA OU AGREGADO'], categories=categorias, ordered=True
)

# Ordenando as categorias e adicionando o total diário
resultado_diario = resultado_diario.sort_values('FROTA OU AGREGADO').reset_index(drop=True)

# Total diário
total_diario = pd.DataFrame({
    'FROTA OU AGREGADO': ['Total'],
    'QUANTIDADE VIAGENS': [resultado_diario['QUANTIDADE VIAGENS'].sum()],
    'SOMA VOLUME CARREGADO': [resultado_diario['SOMA VOLUME CARREGADO'].sum()]
})
resultado_diario = pd.concat([resultado_diario, total_diario], ignore_index=True)

# Resumo Acumulado: Segregando corretamente por categoria
resultado_acumulado = df.groupby('FROTA OU AGREGADO').agg(
    {'DOCUMENTO': 'count', 'VOLUME CARREGADO': 'sum'}
).rename(columns={'DOCUMENTO': 'QUANTIDADE VIAGENS', 'VOLUME CARREGADO': 'SOMA VOLUME CARREGADO'}).reset_index()

# Dividindo o volume carregado por 1000 (transformando para toneladas)
resultado_acumulado['SOMA VOLUME CARREGADO'] = resultado_acumulado['SOMA VOLUME CARREGADO'] / 1000

# Garantindo a ordem das categorias
categorias = ['Frota', 'Terceiros', 'Agregados']
resultado_acumulado['FROTA OU AGREGADO'] = pd.Categorical(
    resultado_acumulado['FROTA OU AGREGADO'], categories=categorias, ordered=True
)

# Ordenando as categorias e adicionando o total acumulado
resultado_acumulado = resultado_acumulado.sort_values('FROTA OU AGREGADO').reset_index(drop=True)

# Total acumulado
total_acumulado = pd.DataFrame({
    'FROTA OU AGREGADO': ['Total'],
    'QUANTIDADE VIAGENS': [resultado_acumulado['QUANTIDADE VIAGENS'].sum()],
    'SOMA VOLUME CARREGADO': [resultado_acumulado['SOMA VOLUME CARREGADO'].sum()]
})
resultado_acumulado = pd.concat([resultado_acumulado, total_acumulado], ignore_index=True)


# Calculando a porcentagem de volume carregado

resultado_diario['PERCENTUAL'] = (
    resultado_diario['SOMA VOLUME CARREGADO'].astype(float) /
    resultado_diario['SOMA VOLUME CARREGADO'].astype(float).sum()
) * 100

# Excluindo a linha "Total" do DataFrame para o gráfico
resultado_diario_grafico = resultado_diario[resultado_diario['FROTA OU AGREGADO'] != 'Total']

# Calculando o percentual para o gráfico sem a linha "Total"
resultado_diario_grafico['PERCENTUAL'] = (
    resultado_diario_grafico['SOMA VOLUME CARREGADO'].astype(float) /
    resultado_diario_grafico['SOMA VOLUME CARREGADO'].astype(float).sum()
) * 100

# Criando o gráfico apenas com as categorias específicas
fig, ax = plt.subplots(figsize=(8, 6))
ax.pie(
    resultado_diario_grafico['PERCENTUAL'],
    labels=resultado_diario_grafico['FROTA OU AGREGADO'],
    autopct='%1.1f%%',
    startangle=90,
    colors=['#4CAF50', '#FFC107', '#2196F3']
)
# Criando o gráfico
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

# Salvando o gráfico como imagem
grafico_path = r'C:\Users\Kaique\Documents\Resumo_Volume_Carregado.png'
plt.savefig(grafico_path)
plt.close()

# Formatando os resultados para o e-mail
for col in ['SOMA VOLUME CARREGADO']:
    resultado_diario[col] = resultado_diario[col].map(lambda x: f"{x:,.2f}".replace(",", "."))
    resultado_acumulado[col] = resultado_acumulado[col].map(lambda x: f"{x:,.2f}".replace(",", "."))

# Removendo a coluna 'PERCENTUAL' antes de gerar o HTML
resultado_diario_html = resultado_diario.drop(columns=['PERCENTUAL'], errors='ignore')
html_diario = resultado_diario_html.to_html(index=False, classes='table table-bordered')
html_acumulado = resultado_acumulado.to_html(index=False, classes='table table-bordered')

# Corpo do e-mail com o gráfico
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

# Função de envio de e-mail com anexo de gráfico
def send_email_with_chart(subject, body, to_email, chart_path):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.BodyFormat = 2  # HTML format
    mail.HTMLBody = body
    mail.To = to_email

    # Anexar gráfico ao e-mail
    mail.Attachments.Add(chart_path)
    # Referenciar o gráfico no corpo do e-mail
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
    'kaique.pimentel@etp-transparana.com.br',
    grafico_path
)