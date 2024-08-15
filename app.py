import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import traceback

root = tk.Tk()
root.withdraw()

def logErro(mensagem):
    with open('log_erros.txt', 'a') as log_file:
        log_file.write(f'{mensagem}\n')

def carregarArquivo():
    arquivoSelecionado = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not arquivoSelecionado:
        logErro('Nenhum arquivo selecionado. Encerrando o programa.')
        sys.exit()
    return arquivoSelecionado

def obterNumMensagens():
    while True:
        try:
            numMensagens = simpledialog.askinteger('Quantidade de mensagens ', 'Quantas mensagens deseja enviar? ')
            if numMensagens is None:
                logErro('Nenhum número válido foi inserido!')
                sys.exit()
            if numMensagens <= 0:
                raise ValueError('O número de mensagens deve ser positivo.')
            break
        except ValueError as e:
            logErro(f'Valor inválido: {e}. Por favor, insira um número inteiro positivo.')
    return numMensagens

def obterMensagens(numMensagens):
    mensagens = []
    for i in range(numMensagens):
        mensagem = simpledialog.askstring('Mensagem', f'Digite a mensagem {i+1}:')
        if mensagem is None:
            logErro('Operação cancelada pelo usuário.')
            sys.exit()
        mensagens.append(mensagem)
    return mensagens

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

def enviarMensagens(navegador, mensagens, paginaClientes):
    linhaVerde = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
    linhaVermelha = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    for linha in paginaClientes.iter_rows(min_row=2):
        matricula = linha[0].value  
        nome = linha[1].value
        telefone = linha[2].value
        email = linha[3].value
        valor_debito = linha[4].value

        if nome is None or telefone is None:
            continue

        primeiraMensagem = f"Olá, {nome}\n\n{mensagens[0]}"

        for i, mensagem in enumerate(mensagens): 
            if i == 0:
                mensagemAtual = primeiraMensagem
            else:
                mensagemAtual = mensagem

            try:
                linkMensagemWhats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagemAtual)}'
                navegador.get(linkMensagemWhats)

                while len(navegador.find_elements(by='id', value='side')) < 1:
                    sleep(10)

                navegador.find_element(by='xpath', value='//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p').send_keys(Keys.ENTER)
                sleep(15)

                navegador.switch_to.window(navegador.window_handles[0])

                sleep(10)

                for cell in linha:
                    cell.fill = linhaVerde

            except Exception as e:
                logErro(f'Não foi possível enviar mensagem para {nome}, ({telefone}). Erro: {e}')
                for cell in linha:
                    cell.fill = linhaVermelha

try:
    arquivoSelecionado = carregarArquivo()

    numMensagens = obterNumMensagens()

    mensagens = obterMensagens(numMensagens)

    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir=/path/to/your/custom/profile")
    navegador = webdriver.Chrome(options=options)

    navegador.get("https://web.whatsapp.com/")

    while len(navegador.find_elements(by='id', value='side')) < 1:
        sleep(5)

    workbook = openpyxl.load_workbook(arquivoSelecionado)
    paginaClientes = workbook['Planilha1']

    enviarMensagens(navegador, mensagens, paginaClientes)

    workbook.save(arquivoSelecionado)

    messagebox.showinfo('Concluído', 'Mensagens enviadas.')

except Exception as e:
   logErro(f'Ocorreu um erro inesperado: {traceback.format_exc()}')

finally:
    navegador.quit()