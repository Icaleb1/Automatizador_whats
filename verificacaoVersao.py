import requests
import os
import subprocess

def verificarAtualizacoes():
    versaoUrl = 'https://githubusercontent.com/Icaleb1/Automatizador_whats/raw/Caleb/versao.txt'
    executavelUrl = 'https://github.com/Icaleb1/Automatizador_whats/raw/Caleb/dist/app.exe'

    # Obtém a versão mais recente do servidor
    resposta = requests.get(versaoUrl)
    if resposta.status_code != 200:
        print(f'Erro ao acessar a versão: {resposta.status_code}')
        return

    versaoMaisRecente = resposta.text.strip()

    # Lê a versão local
    with open('versao.txt', 'r') as arquivo:
        versaoAtual = arquivo.read().strip()

    # Compara as versões
    if versaoMaisRecente > versaoAtual:
        print(f'Nova versão disponível: {versaoMaisRecente}')
        baixarNovaVersao(executavelUrl)
    else:
        print('Você já está usando a versão mais recente.')

def baixarNovaVersao(executavel_url):
    temp_exe_path = 'app_temp.exe'  # Use um nome temporário para o arquivo baixado

    # Baixa o novo executável
    resposta = requests.get(executavel_url)
    if resposta.status_code != 200:
        print(f'Erro ao baixar o executável: {resposta.status_code}')
        return
    
    with open(temp_exe_path, 'wb') as arquivo:
        arquivo.write(resposta.content)
    
    print('Download concluído. Atualizando...')

    # Cria o script em lote para substituir o executável e reiniciar
    with open('Atualizar.bat', 'w') as arquivoBat:
        arquivoBat.write("timeout /t 3 /nobreak > NUL\n")
        arquivoBat.write("move /Y app_temp.exe app.exe\n")
        arquivoBat.write("start app.exe\n")
    
    # Executa o script em lote e encerra o script Python
    subprocess.run(['Atualizar.bat'])
    os._exit(0)

if __name__ == '__main__':
    verificarAtualizacoes()
