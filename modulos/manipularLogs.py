import logging
import os

from dotenv import load_dotenv
from enviaEmail import enviar_email


load_dotenv()
remetente = os.getenv('EMAIL')

log_file = 'erros.log'

def configurar_logging():
    logging.basicConfig(
        filename=log_file,
        level=logging.ERROR,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

class EmailHandler(logging.Handler):
    def emit(self, record):
        log_entry = self.format(record)
        enviar_email(remetente, 'Erro na execução', log_entry, log_file)

def adicionar_email_handler():
    email_handler = EmailHandler()
    email_handler.setLevel(logging.ERROR)
    email_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.getLogger().addHandler(email_handler)