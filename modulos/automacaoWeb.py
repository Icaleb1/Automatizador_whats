from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from urllib.parse import quote
from time import sleep
import random
import logging
from openpyxl.styles import PatternFill
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def enviar_anexo(navegador, telefone, anexo):
    intervaloAleatorio = random.uniform(15, 25)
    try:
        # Abrir a conversa com o telefone do cliente
        link_mensagem_whats = f'https://web.whatsapp.com/send?phone={telefone}'
        navegador.get(link_mensagem_whats)

        # Aguarda a página do WhatsApp carregar
        #Funcional!!!
        WebDriverWait(navegador, 15).until(
            EC.presence_of_element_located((By.ID, 'side'))   
        )
    
        # Clicar no ícone de anexar
        #Funcional!!!
        anexar_icone = WebDriverWait(navegador, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@title='Anexar']"))
        )
        anexar_icone.click()


        # Selecionar o input de arquivo e enviar o anexo
        anexar_documento = navegador.find_element(By.XPATH, "//input[@accept='*']")
        anexar_documento.send_keys(anexo)

        # Clicar no botão de enviar
        sleep(intervaloAleatorio)
        botao_enviar = WebDriverWait(navegador, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']"))
        )
        botao_enviar.click()

        sleep(intervaloAleatorio)
        return True

    except (NoSuchElementException, Exception) as e:
        logging.error(f"Erro ao enviar anexo para {telefone}: {e}")
        return False


def inicializar_navegador():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('start-maximized')
    options.add_argument("user-data-dir=/path/to/your/custom/profile")  # Use o caminho para o perfil do usuário
    navegador = webdriver.Chrome(options=options)
    navegador.get("https://web.whatsapp.com/")
    while len(navegador.find_elements(by='id', value='side')) < 1:
        sleep(10)
    return navegador

def processar_clientes(navegador, mensagens, pagina_clientes, workbook, file_path, anexo):
    green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    intervaloAleatorio = random.uniform(15, 35)

    for linha in pagina_clientes.iter_rows(min_row=2):

        if len(linha) < 3:
            logging.error(f"Linha com dados insuficientes: {linha}")
            continue

        matricula = linha[0].value  
        nome = linha[1].value
        telefone = linha[2].value

        if nome is None or telefone is None:
            continue

        primeira_mensagem = f"Olá, {nome}\n\n{mensagens[0]}"

        for i, mensagem in enumerate(mensagens): 
            mensagem_atual = primeira_mensagem if i == 0 else mensagem

            try:
                link_mensagem_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem_atual)}'
                navegador.get(link_mensagem_whats)

                while len(navegador.find_elements(by='id', value='side')) < 1:
                    sleep(intervaloAleatorio)

                navegador.find_element(by='xpath', value='//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p').send_keys(Keys.ENTER)
                sleep(intervaloAleatorio)

                navegador.switch_to.window(navegador.window_handles[0])

                sleep(intervaloAleatorio)

                if anexo:
                    envio_anexo_sucesso = enviar_anexo(navegador, telefone, anexo)
                    if not envio_anexo_sucesso:
                        raise Exception(f"Falha ao enviar o anexo para o telefone {telefone}")

                for cell in linha:
                    cell.fill = green_fill

            except Exception as e:

                logging.error(f"Erro ao processar cliente {nome} ({telefone}): {e}")

                for cell in linha:
                    cell.fill = red_fill
    try:
        workbook.save(file_path)
    except Exception as e:
        logging.error(f"Erro ao salvar o arquivo: {e}")