import pandas as pd
import win32com.client as win32
import time
import os
from datetime import datetime
import re

print('==== Iniciando script email_auto.py ====\n')
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
print('Iniciando loop de envio...')

# Abrir arquivo de log de erros
print('Abrindo arquivo de log de erros...')
error_log = open('erros_envio.txt', 'a', encoding='utf-8')
print('Arquivo de log de erros aberto.')

for index, row in df.iterrows():
    try:
        print(f'\n--- Processando linha {index} ---')
        # Pula se já foi enviado
        if str(row.get('E-mail enviado', '')).strip().lower() == 'true':
            print('E-mail já enviado anteriormente, pulando...')
            continue
        contato_nome = str(row['Contato']).strip().title()
        primeiro_nome = contato_nome.split()[0]
        destinatario = str(row['E-mail']).strip()
        script = str(row['Script a Enviar']).strip()
        print(f'Contato: {contato_nome}, Destinatário: {destinatario}, Script: {script}')

        if pd.isna(destinatario) or '@' not in destinatario:
            print(f"E-mail inválido para linha {index}: '{destinatario}', pulando...")
            continue

        print(f'Preparando envio para {destinatario}...')

        # Escolhe o assunto e corpo
        if script == "C-LEVEL / TI":
            print('Usando template C-LEVEL / TI')
            assunto = "Como aumentar performance e reduzir custos com tecnologia na sua empresa"
            corpo_personalizado = f"""
Olá {primeiro_nome}, tudo bem?

Meu nome é João, sou consultor comercial da Solaris Tech, uma empresa especializada em soluções de tecnologia focadas em três frentes principais:

-  Automação e desenvolvimento para Totvs (Protheus e RM): Criamos rotinas, integrações e automações personalizadas para otimizar processos operacionais e garantir aderência às regras de negócio. Por exemplo, na IZII (Gestão de Benefícios), reduzimos em até 70% o tempo de operação nas áreas de compras e logística dentro do Protheus.

-  Business Intelligence e Dashboards Gerenciais: Desenvolvemos visões gerenciais estratégicas utilizando Power BI e soluções customizadas. Com a Anexbank (Securitização de Crédito), conseguimos integrar dados financeiros, estoque e vendas em dashboards que aumentaram a visibilidade e reduziram retrabalhos em mais de 40%.

-  Big Data, Data Lake e IoT: Estruturamos ambientes escaláveis em cloud para captura e análise em tempo real de dados provenientes de sensores, dispositivos móveis e ERPs — como fizemos no projeto da AgroMetrics, possibilitando decisões mais rápidas e embasadas.

Se fizer sentido para você, gostaria de agendar uma conversa rápida para entender seus desafios atuais e mostrar como podemos ajudar a aumentar a performance, produtividade e reduzir custos na sua empresa.

Fico à disposição e aguardo seu retorno!

Atenciosamente,
"""
        else:
            print('Usando template padrão de fornecimento de TI')
            assunto = "Fornecimento de itens de TI com agilidade e suporte completo"
            corpo_personalizado = f"""
Olá {primeiro_nome}, tudo bem?

Me chamo João, sou consultor comercial da Solaris Tech, e gostaria de me apresentar como um parceiro estratégico para fornecimento de produtos e soluções de tecnologia.

A Solaris atua com distribuição de equipamentos de TI por meio de parcerias com distribuidores oficiais. Atendemos demandas como:

- Notebooks, servidores, switches, roteadores
- Licenciamento de software
- Componentes de rede e infraestrutura
- Suprimentos homologados com garantia

Nosso diferencial está na agilidade de cotação, prazos competitivos e suporte técnico completo, inclusive com acompanhamento pós-entrega junto às áreas técnicas.

Atendemos empresas de médio e grande porte, respeitando os processos internos de compliance, cotações e SLAs definidos pela área de compras.

Caso faça sentido, posso te enviar uma apresentação institucional e iniciamos um mapeamento das categorias com maior recorrência na sua operação.

Fico à disposição!

Atenciosamente,
"""
        print('Montando e-mail final...')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = assunto
        mail.Display()  # Abre o e-mail e insere a assinatura padrão
        time.sleep(1)  # Dá tempo para o Outlook inserir a assinatura
        inspector = mail.GetInspector
        editor = inspector.WordEditor
        print('Inserindo corpo personalizado no e-mail...')
        editor.Range(0,0).InsertBefore(corpo_personalizado.lstrip().rstrip())
        print('Ajustando tamanho da fonte do corpo...')
        range_corpo = editor.Range(0, len(corpo_personalizado))
        range_corpo.Font.Size = 13
        print('Verificando se está pausado pelo painel...')

        # Solicita confirmação de leitura e entrega
        mail.ReadReceiptRequested = True
        mail.OriginatorDeliveryReportRequested = True

        while os.path.exists('pause.flag'):
            print("Pausado pelo painel... aguardando liberação.")
            time.sleep(2)
        try:
            print('Enviando e-mail...')
            mail.Send()
            print('E-mail enviado com sucesso!')
            # Marca como enviado e registra data
            df.at[index, 'E-mail enviado'] = 'TRUE'
            df.at[index, 'Data de envio'] = datetime.today().strftime('%d/%m/%Y %H:%M')
            emails_enviados += 1
            print(f"E-mail criado para {destinatario} com assinatura.")
        except Exception as e:
            print(f"Erro ao enviar para {destinatario}: {e}")
            error_log.write(f"{destinatario},{str(e)}\n")
            continue
        # Pausa maior a cada 60 e-mails (1 hora)
        if emails_enviados % 60 == 0:
            print(f"Pausa longa de 5 minutos após {emails_enviados} e-mails...")
            time.sleep(300)
        else:
            tempo_espera = 60  # 1 e-mail por minuto
            print(f"Aguardando {tempo_espera} segundos antes do próximo e-mail...")
            time.sleep(tempo_espera)
        print(f'--- Fim do processamento da linha {index} ---')
    except Exception as e:
        print(f"Erro inesperado na linha {index}: {e}")
        error_log.write(f"Linha {index} erro inesperado: {str(e)}\n")
        continue

print('Salvando planilha com status de envio...')
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

print(f"\n==== Fim do script email_auto.py ====")
print(f"Total de {emails_enviados} e-mails enviados com assinatura.\n")