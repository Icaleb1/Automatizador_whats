from tkinter import messagebox
import webbrowser
import requests

def verificarVersaoAtual():
    try:
        with open('versao.txt', 'r') as arquivo:
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
        if versaoAtual == versaoGithub:
            print('Seu software já está atualizado.')
        else:
            messagebox.showinfo('Atualização disponível!', 
                f'Versão local: {versaoAtual}, Versão mais recente: {versaoGithub}')
            link_drive = "https://drive.google.com/drive/folders/1G78v6PE6Vuk6x2RW11KmYiHMqEOyPCaD?usp=drive_link"
            webbrowser.open(link_drive)
    else:
        print('Não foi possível realizar a comparação de versões.')
