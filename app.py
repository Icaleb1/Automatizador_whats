"""
automatizar essa planilha com PYTHON e 
assimilar a EVO para enviar as mensagens automaticamente para os clientes

"""

#Descrever passos manuais e transformar em código
#Ler planilha e guardar informações
#Criar links personalizados do whats e enviar MSGs
#Com base nos dados da planilha

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

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

image_path = resource_path("botao.png")

file_path = filedialog.askopenfilename(
    title="Selecione o arquivo Excel",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not file_path:
    messagebox.showerror('Error', 'nenhum arquivo selecionado. Encerrando o programa.')
    sys.exit()

while True:
    try:
        num_mensagens = simpledialog.askinteger('Quantidade de mensagens ', 'Quantas mensagens deseja enviar? ')
        if num_mensagens is None:
            messagebox.showwarning('Atenção!', 'Digite um número válido!')
            sys.exit()
        if num_mensagens <= 0:
            raise ValueError('O número de mensagens deve ser positivo.')
        break
    except ValueError as e:
        messagebox.showerror('Erro', f'Valor inválido: {e}. Por favor, insira um número inteiro positivo.')

mensagens = []

for i in range(num_mensagens):
    mensagem = simpledialog.askstring('Mensagem', f'Digite a mensagem {i+1}:')
    if mensagem is None:
        messagebox.showwarning('Atenção', 'Operação cancelada.')
        sys.exit()
    mensagens.append(mensagem)

options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=/path/to/your/custom/profile")  # Use o caminho para o perfil do usuário
navegador = webdriver.Chrome(options=options)

navegador.get("https://web.whatsapp.com/")

while len(navegador.find_elements(by='id', value='side')) < 1:
    sleep(5)

workbook = openpyxl.load_workbook(file_path)
pagina_clientes = workbook['Planilha1']

green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

for linha in pagina_clientes.iter_rows(min_row=2):
    matricula = linha[0].value  
    nome = linha[1].value
    telefone = linha[2].value
    email = linha[3].value

    valor_debito = linha[4].value

    if nome is None or telefone is None:
        continue

    primeira_mensagem = f"Olá, {nome}\n\n{mensagens[0]}"

    for i, mensagem in enumerate(mensagens): 

        if i == 0:
            mensagem_atual = primeira_mensagem
        else:
            mensagem_atual = mensagem

        try:
            link_mensagem_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem_atual)}'
            navegador.get(link_mensagem_whats)

            while len(navegador.find_elements(by='id', value='side')) < 1:
                sleep(10)

            navegador.find_element(by='xpath', value='//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p').send_keys(Keys.ENTER)
            sleep(15)

            navegador.switch_to.window(navegador.window_handles[0])

            sleep(10)

            for cell in linha:
                cell.fill = green_fill

        except Exception as e:
            messagebox.showerror('Erro', f'Não foi possível enviar mensagem para {nome}, ({telefone}). Detalhes do erro foram registrados.')

            for cell in linha:
                cell.fill = red_fill

workbook.save(file_path)

messagebox.showinfo('Concluído', 'Mensagens enviadas.')
