from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from urllib.parse import quote
from time import sleep
import random
import logging
from openpyxl.styles import PatternFill


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

def processar_clientes(navegador, mensagens, pagina_clientes, workbook, file_path):
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