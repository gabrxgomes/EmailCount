import mailbox
import pandas as pd
from collections import defaultdict
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import dateparser

# Function to parse the date from the email headers
def converter_data(data_str):
    return dateparser.parse(data_str)

# Function to calculate time active and number of sent emails per month
def calcular_tempo_ativo_e_emails_por_mes(emails, meu_identificador):
    # Dictionary to store active time and number of emails by month and recipient
    dados_por_mes_e_destinatario = defaultdict(lambda: {'tempo_ativo': 0, 'emails_enviados': 0})

    # Filter emails sent by you
    meus_emails = [email for email in emails if email.get('From') and meu_identificador in email['From']]

    # Group emails by recipient and month
    for email in meus_emails:
        destinatario = obter_destinatario(email)
        data_email = converter_data(email['Date'])

        if data_email:
            chave_mes = f"{data_email.year}-{data_email.month:02d}"

            # Add the count of emails sent and time spent
            dados_por_mes_e_destinatario[(chave_mes, destinatario)]['emails_enviados'] += 1

    return dados_por_mes_e_destinatario

# Function to get the recipient of the emails
def obter_destinatario(email):
    return email.get('To', 'Desconhecido')

# Hide the root window for Tkinter file dialog
root = Tk()
root.withdraw()

# Ask the user to select an MBOX file
caminho_arquivo_mbox = askopenfilename(title="Selecione o arquivo MBOX", filetypes=[("MBOX files", "*.mbox")])

# Check if the file was selected
if not caminho_arquivo_mbox:
    print("Nenhum arquivo foi selecionado. O programa será encerrado.")
else:
    meu_identificador = ''  # Your email identifier

    # Load the MBOX file
    mbox = mailbox.mbox(caminho_arquivo_mbox)

    # Dictionary to store processed data grouped by recipient and month
    dados_agrupados = defaultdict(lambda: defaultdict(lambda: {'emails_enviados': 0, 'tempo_ativo': 0}))

    # Process each message in the MBOX
    for message in mbox:
        if message['From'] and meu_identificador in message['From']:  # Ensure email was sent by you
            destinatario = obter_destinatario(message)
            data_email = converter_data(message['Date'])

            if data_email:
                chave_mes = f"{data_email.year}-{data_email.month:02d}"

                # Add the email to the grouped data by month and recipient
                dados_agrupados[destinatario][chave_mes]['emails_enviados'] += 1

    # Prepare the final data for the report
    dados_resultantes = []
    for destinatario, meses in dados_agrupados.items():
        for mes, dados in meses.items():
            # Format the data for display
            tempo_ativo_formatado = f"{int(dados['tempo_ativo'] // 3600)} horas e {int((dados['tempo_ativo'] % 3600) // 60)} minutos"
            emails_formatado = f"{mes}: {dados['emails_enviados']} emails"

            # Append the data to the report
            dados_resultantes.append({
                'Destinatário': destinatario,
                'Emails Enviados': emails_formatado,
                'Tempo Gasto (h:m:s)': tempo_ativo_formatado if tempo_ativo_formatado else '0 horas e 0 minutos'
            })

    # Check if any data was processed
    if not dados_resultantes:
        print("Nenhum dado foi encontrado ou processado.")
    else:
        # Create a DataFrame with the resulting data
        df = pd.DataFrame(dados_resultantes)

        # Save the DataFrame to an Excel file
        caminho_arquivo_excel = 'relatorio_emails_mbox.xlsx'
        df.to_excel(caminho_arquivo_excel, index=False)
        print(f"Relatório salvo em: {caminho_arquivo_excel}")
