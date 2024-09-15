import subprocess
import os

def criar_executavel():
    sep = ';' if os.name == 'nt' else ':'
    comando = [
        'pyinstaller',
         '--onefile',
        '--add-data', f'enviaEmail.py{sep}.',
        '--add-data', f'verificarVersao.py{sep}.',
        '--add-data', f'versao.txt{sep}.',
        '--add-data', f'.env{sep}.',
        '--add-data', f'erros.log{sep}.',
        '--add-data', f'modulos{sep}modulos',  # Inclui a pasta modulos
        '--noconsole',
        'app.py'
    ]
    
    try:
        subprocess.run(comando, check=True)
        print("Executável criado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao criar o executável: {e}")

if __name__ == "__main__":
    criar_executavel()
