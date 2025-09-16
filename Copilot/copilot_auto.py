import pandas as pd
import win32com.client as win32
import time
import os
from datetime import datetime
import csv

base_dir = os.path.dirname(os.path.abspath(__file__))
template_path = os.path.join(base_dir, 'copilot_templates.csv')

print('==== Iniciando script copilot_auto.py ====')

# Carregar planilha
try:
    df = pd.read_csv("planilha_processada_solaris.csv", sep=',')
    print(f'Planilha carregada com {len(df)} linhas.')
except Exception as e:
    print(f'Erro ao carregar planilha: {e}')
    raise

# Carregar templates de e-mail
templates = {}
with open(template_path, encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        templates[row['tipo'].strip().upper()] = {
            'assunto': row['assunto'],
            'corpo': row['corpo']
        }

# Conectar ao Outlook e capturar assinatura
try:
    print('Conectando ao Outlook...')
    outlook = win32.Dispatch('Outlook.Application')
    print('Conexão com Outlook realizada.')
    print('Capturando assinatura do Outlook...')
    temp_mail = outlook.CreateItem(0)
    temp_mail.Display()
    time.sleep(1)  # Tempo extra para garantir carregamento completo
    assinatura = temp_mail.HTMLBody
    temp_mail.Close(0)
    print('Assinatura capturada.')
except Exception as e:
    print(f'Erro ao conectar ao Outlook ou capturar assinatura: {e}')
    raise

emails_enviados = 0
print('Iniciando envio de e-mails Copilot...')

LINHA_LIMITE = 1122
PULAR_LINHAS = 400

for index, row in df.iterrows():
    if index < PULAR_LINHAS:
        continue
    if index >= LINHA_LIMITE:
        print(f'Parando envio a partir da linha {LINHA_LIMITE}')
        break
    try:
        print(f'\n--- Processando linha {index} ---')
        if str(row.get('E-mail Enviado', '')).strip().lower() == 'true':
            print('E-mail já enviado anteriormente, pulando...')
            continue

        contato_nome = str(row['Contato']).strip().title()
        primeiro_nome = contato_nome.split()[0]
        destinatario = str(row['E-mail']).strip()
        script = str(row['Script a Enviar']).strip().upper()

        if pd.isna(destinatario) or '@' not in destinatario:
            print(f"E-mail inválido para linha {index}: '{destinatario}', pulando...")
            continue

        print(f'Preparando envio de e-mail para {destinatario}...')

        tipo_template = script if script in templates else 'COPILOT'
        assunto = templates[tipo_template]['assunto']
        corpo_personalizado = templates[tipo_template]['corpo'].format(primeiro_nome=primeiro_nome)

        print('Montando e-mail...')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = assunto
        mail.Display()  # Abre o e-mail e insere a assinatura padrão
        time.sleep(1)  # Dá tempo para o Outlook inserir a assinatura
        inspector = mail.GetInspector
        editor = inspector.WordEditor
        print('Inserindo corpo personalizado no e-mail...')
        editor.Range(0, 0).InsertBefore(corpo_personalizado.lstrip().rstrip())
        print('Ajustando tamanho da fonte do corpo...')
        range_corpo = editor.Range(0, len(corpo_personalizado))
        range_corpo.Font.Size = 13

        mail.ReadReceiptRequested = True

        print('Verificando se está pausado pelo painel...')
        while os.path.exists('pause.flag'):
            print("Pausado pelo painel... aguardando liberação.")
            time.sleep(2)

        print('Enviando e-mail...')
        mail.Send()
        print('E-mail enviado com sucesso!')

        df.at[index, 'E-mail Enviado'] = 'TRUE'
        df.at[index, 'Data de Envio'] = datetime.today().strftime('%d/%m/%Y %H:%M')
        emails_enviados += 1
        print(f"E-mail enviado para {destinatario}")
        time.sleep(60)
        print(f'--- Fim do processamento da linha {index} ---')

    except Exception as e:
        print(f"Erro na linha {index}: {e}")
        continue

print('Salvando planilha...')
df.to_csv("planilha_processada_copilot.csv", sep=',', index=False)
print(f'Finalizado. Total de {emails_enviados} e-mails enviados.')