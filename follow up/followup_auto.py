import pandas as pd
import win32com.client as win32
import time
import os
from datetime import datetime
import re
import csv

print('==== Iniciando script followup_auto.py ====' )
try:
    print('Importando módulos...')
    import pandas as pd
    import win32com.client as win32
    import time
    import random
    import os
    from datetime import datetime
    import re
    print('Módulos importados com sucesso.')
except Exception as e:
    print(f'Erro ao importar módulos: {e}')
    raise

# Carregar o CSV
try:
    print('Carregando planilha...')
    df = pd.read_csv("planilha_processada_solaris.csv", sep=',')
    print(f'Planilha carregada com {len(df)} linhas.')
    print(f'Colunas encontradas: {list(df.columns)}')
except Exception as e:
    print(f'Erro ao carregar planilha: {e}')
    raise

# Carregar templates de follow-up
templates = {}
with open('followup_templates.csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        templates[row['tipo'].strip().upper()] = {
            'assunto': row['assunto'],
            'corpo': row['corpo']
        }

# Conectar ao Outlook
try:
    print('Conectando ao Outlook...')
    outlook = win32.Dispatch('outlook.application')
    print('Conexão com Outlook realizada.')
except Exception as e:
    print(f'Erro ao conectar ao Outlook: {e}')
    raise

# Pega a assinatura do Outlook via Display()
try:
    print('Capturando assinatura do Outlook...')
    temp_mail = outlook.CreateItem(0)
    temp_mail.Display()
    time.sleep(1)  # dá tempo pro Outlook inserir a assinatura
    assinatura_html = temp_mail.HTMLBody
    temp_mail.Close(0)  # fecha sem salvar
    print('Assinatura capturada.')
except Exception as e:
    print(f'Erro ao capturar assinatura: {e}')
    raise

emails_enviados = 0
print('Iniciando loop de envio de follow-ups...')

# Abrir arquivo de log de erros
print('Abrindo arquivo de log de erros...')
error_log = open('erros_envio.txt', 'a', encoding='utf-8')
print('Arquivo de log de erros aberto.')

# Defina aqui a linha limite para envio de follow-ups
LINHA_LIMITE = 1122  # Altere para o número desejado (ex: 100)

# Defina aqui quantas linhas deseja pular no início
PULAR_LINHAS = 1039 # Altere para o número desejado (ex: 10)

for index, row in df.iterrows():
    # Pular as X primeiras linhas
    if index < PULAR_LINHAS:
        continue
    # Parar o envio a partir da linha limite
    if index >= LINHA_LIMITE:
        print(f'Parando envio a partir da linha {LINHA_LIMITE}')
        break
    try:
        print(f'\n--- Processando linha {index} ---')
        # Pula se já foi enviado follow-up
        if str(row.get('Follow-up enviado', '')).strip().lower() == 'true':
            print('Follow-up já enviado anteriormente, pulando...')
            continue
        contato_nome = str(row['Contato']).strip().title()
        primeiro_nome = contato_nome.split()[0]
        destinatario = str(row['E-mail']).strip()
        script = str(row['Script a Enviar']).strip()
        print(f'Contato: {contato_nome}, Destinatário: {destinatario}, Script: {script}')

        if pd.isna(destinatario) or '@' not in destinatario:
            print(f"E-mail inválido para linha {index}: '{destinatario}', pulando...")
            continue

        print(f'Preparando envio de follow-up para {destinatario}...')

        # Seleciona template pelo tipo
        tipo_template = script.strip().upper() if script.strip().upper() in templates else 'PADRAO'
        assunto = templates[tipo_template]['assunto']
        corpo_personalizado = templates[tipo_template]['corpo'].format(primeiro_nome=primeiro_nome)
        print('Montando e-mail final de follow-up...')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = assunto
        mail.Display()  # Abre o e-mail e insere a assinatura padrão
        time.sleep(1)  # Dá tempo para o Outlook inserir a assinatura
        inspector = mail.GetInspector
        editor = inspector.WordEditor
        print('Inserindo corpo personalizado no e-mail de follow-up...')
        editor.Range(0,0).InsertBefore(corpo_personalizado.lstrip().rstrip())
        print('Ajustando tamanho da fonte do corpo...')
        range_corpo = editor.Range(0, len(corpo_personalizado))
        range_corpo.Font.Size = 13

        # Solicita confirmação de leitura e entrega
        mail.ReadReceiptRequested = True

        print('Verificando se está pausado pelo painel...')
        while os.path.exists('pause.flag'):
            print("Pausado pelo painel... aguardando liberação.")
            time.sleep(2)
        try:
            print('Enviando e-mail de follow-up...')
            mail.Send()
            print('Follow-up enviado com sucesso!')
            # Marca como enviado e registra data
            df.at[index, 'Follow-up enviado'] = 'TRUE'
            df.at[index, 'Data de envio follow-up'] = datetime.today().strftime('%d/%m/%Y %H:%M')
            emails_enviados += 1
            print(f"Follow-up criado para {destinatario} com assinatura.")
        except Exception as e:
            print(f"Erro ao enviar follow-up para {destinatario}: {e}")
            error_log.write(f"{destinatario},{str(e)}\n")
            continue
        # Pausa maior a cada 60 e-mails (1 hora)
        if emails_enviados % 60 == 0:
            print(f"Pausa longa de 5 minutos após {emails_enviados} follow-ups...")
            time.sleep(300)
            emails_enviados = 0
        else:
            tempo_espera = 60  # 1 follow-up por minuto
            print(f"Aguardando {tempo_espera} segundos antes do próximo follow-up...")
            time.sleep(tempo_espera)
        print(f'--- Fim do processamento da linha {index} ---')
    except Exception as e:
        print(f"Erro inesperado na linha {index}: {e}")
        error_log.write(f"Linha {index} erro inesperado: {str(e)}\n")
        continue

print('Salvando planilha com status de follow-up...')
df.to_csv("planilha_processada_solaris.csv", sep=',', index=False)
print('Planilha salva.')

# --- Processamento de mensagens de erro de entrega (bounces) ---
try:
    print("Processando mensagens de erro de entrega na caixa de entrada...")
    import win32com.client
    outlook_ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook_ns.GetDefaultFolder(6)  # 6 = inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Mais recentes primeiro
    count = 0
    # Palavras-chave para identificar bounces
    bounce_keywords_subject = [
        "couldn't be delivered", "undeliverable", "falha na entrega", "delivery has failed",
        "delivery status notification", "mail delivery failed", "failure notice", "returned mail",
        "não foi possível entregar", "não entregue", "não entregue ao destinatário", "mailbox unavailable",
        "user unknown", "recipient not found", "no such user", "host not found", "domain does not exist",
        "dns error", "mail system error"
    ]
    bounce_keywords_body = [
        "dns", "domain does not exist", "user unknown", "mailbox unavailable", "recipient not found",
        "no such user", "host not found", "could not be delivered", "delivery failed", "não foi possível entregar",
        "usuário desconhecido", "caixa postal inexistente", "mail system error", "invalid address", "address not found"
    ]
    for message in messages:
        if count > 200:
            print(f"Limite de 200 mensagens atingido. Parando processamento de bounces.")
            break
        count += 1
        if getattr(message, 'Class', None) != 43:  # 43 = MailItem
            continue
        subject = (message.Subject or "").lower()
        body = (message.Body or "").lower()
        print(f"Analisando mensagem {count}: Assunto='{subject}'")
        erro_detectado = False
        # Verifica palavras-chave no assunto
        for kw in bounce_keywords_subject:
            if kw in subject:
                print(f"Mensagem identificada como erro de entrega pelo assunto: '{kw}'")
                erro_detectado = True
                break
        # Se não detectou pelo assunto, verifica o corpo
        if not erro_detectado:
            for kw in bounce_keywords_body:
                if kw in body:
                    print(f"Mensagem identificada como erro de entrega pelo corpo: '{kw}'")
                    erro_detectado = True
                    break
        if erro_detectado:
            match = re.search(r'[\w\.-]+@[\w\.-]+', body)
            if match:
                destinatario = match.group(0).lower()
                print(f"Destinatário extraído do erro: {destinatario}")
                idx = df[df['E-mail'].str.lower() == destinatario].index
                if len(idx) > 0:
                    print(f"Destinatário encontrado na planilha, adicionando observação de erro.")
                    obs_antiga = str(df.at[idx[0], 'Observações']) if 'Observações' in df.columns else ''
                    novo_erro = f"Erro de entrega detectado: {subject} - {body[:100]}"
                    if obs_antiga and obs_antiga != 'nan':
                        df.at[idx[0], 'Observações'] = obs_antiga + ' | ' + novo_erro
                    else:
                        df.at[idx[0], 'Observações'] = novo_erro
                else:
                    print(f"Destinatário {destinatario} não encontrado na planilha.")
            else:
                print("Não foi possível extrair destinatário do corpo da mensagem de erro.")
    print("Salvando planilha com observações de erro de entrega...")
    df.to_csv("planilha_processada_solaris.csv", sep=',', index=False)
    print('Planilha salva com observações de erro de entrega.')
except Exception as e:
    print(f"Erro ao processar mensagens de erro de entrega: {e}")

print(f"\n==== Fim do script followup_auto.py ====")
print(f"Total de {emails_enviados} follow-ups enviados com assinatura.\n") 