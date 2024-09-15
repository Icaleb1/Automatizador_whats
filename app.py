import logging
from tkinter import messagebox

import openpyxl
from modulos.interface import carregar_arquivo, escolher_pagina, obter_mensagens, obter_num_mensagens, carregar_anexo
from modulos.automacaoWeb import inicializar_navegador, processar_clientes
from modulos.manipularArquivos import resource_path
from modulos.manipularLogs import configurar_logging, adicionar_email_handler
from verificarVersao import compararVersoes

def main():
    configurar_logging()
    adicionar_email_handler()
    compararVersoes()

    try:
        anexo = carregar_anexo()
        file_path = carregar_arquivo()
        workbook = openpyxl.load_workbook(file_path)
        pagina_selecionada = escolher_pagina(workbook)
        pagina_clientes = workbook[pagina_selecionada]

        num_mensagens = obter_num_mensagens()
        mensagens = obter_mensagens(num_mensagens)
        navegador = inicializar_navegador()
        processar_clientes(navegador, mensagens, pagina_clientes, workbook, file_path, anexo)
        messagebox.showinfo('Concluído', 'Mensagens enviadas.')
    except Exception as e:
        logging.error(f"Erro na função principal: {e}")
    finally:
        try:
            navegador.quit()
        except Exception as e:
            logging.error(f"Erro ao tentar fechar o navegador: {e}")
        else:
            open('erros.log', 'w').close()

if __name__ == '__main__':
    main()
