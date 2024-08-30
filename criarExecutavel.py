import subprocess

def criar_executavel():
    comando = [
        'pyinstaller',
        '--onefile',
        '--add-data', 'enviaEmail.py;.',
        '--add-data', 'verificarVersao.py;.',
        '--add-data', 'versao.txt;.',
        '--add-data', '.env;.',
        '--add-data', 'erros.log;.',
        '--add-data', 'modulos;modulos',  # Inclui a pasta modulos
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
