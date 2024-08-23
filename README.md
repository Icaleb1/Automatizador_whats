# WhatsApp Message Automation Script

Este projeto contém um script em Python que automatiza o envio de mensagens pelo WhatsApp, utilizando Selenium para interação com o WhatsApp Web. Além disso, ele integra a funcionalidade de envio de e-mails com logs de erros como anexo, permitindo que os administradores sejam notificados automaticamente em caso de falhas durante a execução.

## Funcionalidades

- **Envio Automático de Mensagens**: O script lê uma planilha Excel com informações dos clientes (nome, telefone) e envia mensagens personalizadas via WhatsApp Web.
- **Interface Gráfica**: Utiliza a biblioteca Tkinter para selecionar o arquivo Excel, determinar a quantidade e o conteúdo das mensagens.
- **Registro de Erros**: Erros durante a execução são registrados em um arquivo de log (`erros.log`) e enviados automaticamente por e-mail.
- **Envio de E-mail**: Em caso de erro, um e-mail é enviado automaticamente com o log anexo, usando as configurações do Gmail.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação principal do projeto.
- **Selenium**: Usado para automatizar a interação com o WhatsApp Web.
- **Tkinter**: Interface gráfica para facilitar a interação do usuário com o script.
- **OpenPyXL**: Manipulação de arquivos Excel para leitura e escrita de dados.
- **Logging**: Registro de erros em um arquivo de log.
- **smtplib**: Envio de e-mails através do Gmail.
- **dotenv**: Carregamento de variáveis de ambiente para segurança das credenciais.

## Como Usar

1. **Instalação de Dependências**: Instale as bibliotecas necessárias utilizando `pip`:
   ```bash
   pip install -r requirements.txt
   ```

2. **Configuração**: Crie um arquivo `.env` na raiz do projeto com as seguintes variáveis:
   ```
   EMAIL=seu-email@gmail.com
   SENHA=sua-senha-do-email
   ```

3. **Execução**: Execute o script principal (`app.py`):
   ```bash
   python app.py
   ```

4. **Seleção do Arquivo**: Escolha o arquivo Excel contendo a lista de clientes.

5. **Configuração das Mensagens**: Insira o número de mensagens e os respectivos textos.

6. **Envio de Mensagens**: O script abrirá o WhatsApp Web e começará a enviar as mensagens.

## Estrutura do Projeto

- **app.py**: Script principal que gerencia a interface gráfica, leitura do arquivo Excel, envio de mensagens via WhatsApp e registro de erros.
- **enviaEmail.py**: Script responsável por enviar e-mails com logs de erros.
- **erros.log**: Arquivo de log onde os erros são registrados.
- **requirements.txt**: Lista de dependências do projeto.

## Considerações Finais

Este projeto foi desenvolvido para automatizar o envio de mensagens de forma eficiente e segura. Certifique-se de que o caminho para o perfil do usuário no Chrome esteja correto e de que as permissões de acesso ao e-mail estejam configuradas corretamente.

