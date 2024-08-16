import requests
import os
import subprocess

def verificarAtualizações():
    versao = 'https://github.com/Icaleb1/Automatizador_whats/blob/Caleb/version.txt'
    executavel = 'https://github.com/Icaleb1/Automatizador_whats/blob/Caleb/dist/app.exe'

    resposta = requests.get(versao)
    versaoMaisRecente = resposta.text.strip()

    with open('versao.txt', 'r') as arquivo:
        versaoAtual = arquivo.read().strip()

    if versaoMaisRecente > versaoAtual:
        print(f'Nova versão disponível: {versaoMaisRecente}')
        baixarNovaVersao(executavel)
    else:
        print(f'Você já está usando a versão mais recente.')

    def baixarNovaVersao(executavel):
        resposta = requests.get(executavel)
        with open('app.exe', 'wb') as arquivo:
            arquivo.write(resposta.content)
        
        print('Download concluído. Atualizando...')

        with open('Atualizar.bat', 'w') as arquivoBat:
            arquivoBat.write("timeout /t 3 /nobreak > NUL\n")
            arquivoBat.write("move /Y meu_script_latest.exe meu_script.exe\n")
            arquivoBat.write("start meu_script.exe\n")
        subprocess.run(['atualizar.bat'])
        os._exit(0)

if __name__ == '__main__':
    verificarAtualizações()    
