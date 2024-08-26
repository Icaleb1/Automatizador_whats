from tkinter import messagebox
import webbrowser
import requests
import sys
import os

def verificarVersaoAtual():
    try:
        if getattr(sys, 'frozen', False):
            # Executando como um executável PyInstaller
            caminho_versao = os.path.join(sys._MEIPASS, 'versao.txt')
        else:
            # Executando no ambiente de desenvolvimento
            caminho_versao = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'versao.txt')

        with open(caminho_versao, 'r') as arquivo:
            versaoAtual = arquivo.read().strip()
            return versaoAtual
    except FileNotFoundError:
        print('Arquivo da versão atual e local não encontrado.')
        return None

def obterVersaoGithub():
    url = 'https://raw.githubusercontent.com/Icaleb1/Automatizador_whats/main/versao.txt'
    response = requests.get(url)
    
    if response.status_code == 200:
        versaoGithub = response.text.strip()
        return versaoGithub
    else:
        print(f'Erro ao acessar o arquivo: {response.status_code}')
        return None

def compararVersoes():
    versaoAtual = verificarVersaoAtual()
    versaoGithub = obterVersaoGithub()
    
    if versaoAtual and versaoGithub:
        if versaoAtual != versaoGithub:
            resposta = messagebox.askyesno('Atualização disponível!',
                f'Versão local: {versaoAtual}\nVersão mais recente: {versaoGithub}\n\nDeseja atualizar agora?')
            if resposta:
                link_drive = "https://drive.google.com/drive/folders/1G78v6PE6Vuk6x2RW11KmYiHMqEOyPCaD?usp=drive_link"
                webbrowser.open(link_drive)
    else:
        print('Não foi possível realizar a comparação de versões.')