import requests

# URL do arquivo bruto no GitHub
url = "https://raw.githubusercontent.com/usuario/repositorio/main/versao.txt"

# Solicitação HTTP GET
response = requests.get(url)

if response.status_code == 200:
    versao = response.text.strip()
    print(f"Versão atual: {versao}")
else:
    print(f"Erro ao acessar o arquivo: {response.status_code}")
