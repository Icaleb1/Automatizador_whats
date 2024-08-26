from cx_Freeze import setup, Executable

# Configurações para incluir arquivos adicionais
build_exe_options = {
    "packages": ["os"],
    "include_files": [
        ("enviaEmail.py", "enviaEmail.py"),  
        ("verificarVersao.py", "verificarVersao.py"),
        ("versao.txt", "versao.txt"),  
        (".env", ".env")  
    ]
}

setup(
    name="BotWhats",
    version="1.0",
    description="Automatizador do WhatsApp",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py")]
)
