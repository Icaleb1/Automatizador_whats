import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ssl
from dotenv import load_dotenv
import os

load_dotenv()

def enviar_email(destinatario, assunto, corpo, anexo=None):
    remetente = os.getenv('EMAIL')
    senha = os.getenv('SENHA')

    # Configuração do e-mail
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = destinatario
    mensagem['Subject'] = assunto

    mensagem.attach(MIMEText(corpo, 'plain'))
    
    if anexo:
        anexo_absoluto = os.path.abspath(anexo)
        if not os.path.isfile(anexo_absoluto):
            print(f"Arquivo não encontrado: {anexo_absoluto}")
            return

        with open(anexo_absoluto, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(anexo)}')

            mensagem.attach(part)
    
    contexto = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as server:
            server.login(remetente, senha)
            server.sendmail(remetente, destinatario, mensagem.as_string())
        print(f"E-mail enviado para {destinatario}")
    except Exception as e:
        print(f"Falha ao enviar e-mail para {destinatario}: {e}")

# Teste de envio de e-mail
enviar_email('calebfernandes080@gmail.com', 'testando o envio de emails', 'corpo', 'erros.csv')
