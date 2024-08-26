import logging
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from urllib.parse import quote
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import threading
from enviaEmail import enviar_email
import random
from verificarVersao import compararVersoes

compararVersoes()

log_file = 'erros.log'
logging.basicConfig(
    filename=log_file,
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class EmailHandler(logging.Handler):
    def emit(self, record):
        log_entry = self.format(record)
        enviar_email('calebfernandes080@gmail.com', 'Erro na execução', log_entry, log_file)

email_handler = EmailHandler()
email_handler.setLevel(logging.ERROR)
email_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger().addHandler(email_handler)


# Configurações iniciais
root = tk.Tk()
root.withdraw()


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

def carregar_arquivo():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        messagebox.showerror('Error', 'Nenhum arquivo selecionado. Encerrando o programa.')
        sys.exit()
    return file_path

def escolher_pagina(workbook):
    paginas = workbook.sheetnames
    pagina_selecionada = simpledialog.askstring('Seleção de Página', f'Selecione a página a ser processada:\n\n{", ".join(paginas)}')
    if pagina_selecionada not in paginas:
        messagebox.showerror('Erro', 'Página selecionada não encontrada. Encerrando o programa.')
        sys.exit()
    return pagina_selecionada


def obter_num_mensagens():
    while True:
        try:
            num_mensagens = simpledialog.askinteger('Quantidade de mensagens', 'Quantas mensagens deseja enviar?')
            if num_mensagens is None:
                messagebox.showwarning('Atenção!', 'Digite um número válido!')
                sys.exit()
            if num_mensagens <= 0:
                raise ValueError('O número de mensagens deve ser positivo.')
            return num_mensagens
        except ValueError as e:
            messagebox.showerror('Erro', f'Valor inválido: {e}. Por favor, insira um número inteiro positivo.')

def obter_mensagens(num_mensagens):
    mensagens = []
    for i in range(num_mensagens):
        mensagem = simpledialog.askstring('Mensagem', f'Digite a mensagem {i+1}:')
        if mensagem is None:
            messagebox.showwarning('Atenção', 'Operação cancelada.')
            sys.exit()
        mensagens.append(mensagem)
    return mensagens

def inicializar_navegador():
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir=/path/to/your/custom/profile")  # Use o caminho para o perfil do usuário
    navegador = webdriver.Chrome(options=options)
    navegador.get("https://web.whatsapp.com/")
    while len(navegador.find_elements(by='id', value='side')) < 1:
        sleep(10)
    return navegador

def processar_clientes(navegador, mensagens, pagina_clientes):
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

    workbook.save(file_path)

def main():
    try:
        file_path = carregar_arquivo()
        workbook = openpyxl.load_workbook(file_path)
        pagina_selecionada = escolher_pagina(workbook)
        pagina_clientes = workbook[pagina_selecionada]

        num_mensagens = obter_num_mensagens()
        mensagens = obter_mensagens(num_mensagens)
        navegador = inicializar_navegador()
        processar_clientes(navegador, mensagens, pagina_clientes)
        messagebox.showinfo('Concluído', 'Mensagens enviadas.')
    except Exception as e:
        logging.error(f"Erro na função principal: {e}")
    finally:
        try:
            navegador.quit()
        except Exception as e:
            logging.error(f"Erro ao tentar fechar o navegador: {e}")
        else:
            open(log_file, 'w').close()


if __name__ == '__main__':
    main()
