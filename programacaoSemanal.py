# Bibliotecas e credenciais necessárias para executar este programa:
# 1. Instale o `gspread` para interagir com o Google Sheets.
#    Comando: `pip install gspread`

# 2. Instale o `oauth2client` para gerenciar a autenticação usando credenciais da API Google.
#    Comando: `pip install oauth2client`

# 3. Acesse a API Google Sheets e a API Google Drive para obter permissões.
#    - Vá para https://console.developers.google.com/ e crie um projeto.
#    - Ative a "Google Sheets API" e a "Google Drive API" para o projeto.
#    - Em "Credenciais", crie uma conta de serviço e gere um arquivo de chave JSON.
#    - Salve o arquivo de credenciais com oNomeQueQuiser.json na mesma pasta do script.
#    - Adicione permissões de edição na planilha para o e-mail associado à conta de serviço.

# Nota: Certifique-se de que o `credentials.json` esteja no mesmo diretório que o script ou o executável.

# Outros requisitos:
# - Python 3.x: Certifique-se de estar usando a versão 3 ou superior para garantir compatibilidade com as bibliotecas que estou usando nesse projeto.
# - Biblioteca padrão `datetime`, `os`, e `sys`, que já vêm com Python.

import gspread  # Biblioteca para interagir com o Google Sheets
from oauth2client.service_account import ServiceAccountCredentials  # Autenticação de conta de serviço (É importante que você crie a sua para poder usar esse script)
from datetime import datetime  # Biblioteca para trabalhar com datas e horas
import os  # Biblioteca para manipulação de arquivos e diretórios
import sys  # Biblioteca para funções e variáveis do sistema

# Configuração para identificar o caminho correto da pasta atual (Isso porque transformei o aplicativo em executável e o enviei para o Supervisor que usará minha ferramenta)
# 'sys.frozen' verifica se o código está sendo executado como um executável
if getattr(sys, 'frozen', False):
    pasta_atual = os.path.dirname(sys.executable)  # Para um executável, pega o diretório do executável
else:
    pasta_atual = os.path.dirname(os.path.abspath(__file__))  # Para um script, pega o diretório do script

# Caminho para o arquivo de credenciais JSON (a chave que você cria após criar a sua conta de serviço)
credentials_path = os.path.join(pasta_atual, 'credentials.json')

# Definindo o escopo de permissões para acessar o Google Sheets e o Google Drive
escopo = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"] # Um detalhe importante
# Autenticando com as credenciais e escopo especificados
creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, escopo)
client = gspread.authorize(creds)  # Autorizando o cliente do gspread

# Abrindo a planilha e acessando a primeira aba
sheet = client.open("Programação semanal").sheet1

# Buscando todos os registros da planilha
dados = sheet.get_all_records()
# Verifica se a planilha está vazia e encerra o programa se não houver dados
if not dados:
    print("A planilha está vazia. Verifique e tente novamente.")
    sys.exit() # Eu optei apenas por sair, você pode criar uma mensagem de erro adicional se preferir e depois encerrar o executável.

# Nome do arquivo de saída, incluindo a data atual no formato dia-mês-ano
nome_arquivo = f'Programação - {datetime.now().date().strftime("%d-%m-%Y")}.txt'

# Criando e escrevendo o cabeçalho do arquivo de texto
with open(nome_arquivo, 'w', encoding='utf-8') as mensagem:
    mensagem.write('*=========== PROGRAMAÇÃO DIÁRIA - NOI =========*\n\n')
    
    # Lista de dias da semana, usada para exibir o dia correspondente à previsão de volta
    dias_da_semana = ['Segunda-feira', 'terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']

    # Loop para processar cada linha (registro) da planilha
    for row in dados:
        # Preparando os dados para formatação
        tarefa = row['Tarefa'].replace('\n', ',\n')  # Formata a tarefa com quebras de linha
        num_protocolo = row['Protocolo']
        status_os = row['Ordem de serviço']
        colab1, colab2 = row['Colaborador 1'], row['Colaborador 2']
        observacao = row['Observações']
        prev_volta = datetime.strptime(row['Previsão de volta'], '%d/%m/%Y').date()  # Converte a data de string para datetime
        placa_viatura = str(row['Placa']).upper()  # Placa da viatura, convertida para maiúsculas

        # Ajusta o texto para a previsão de volta
        if prev_volta == datetime.now().date():
            prev_volta = '*Previsão de volta:* Hoje'  # Se for hoje, indica "Hoje"
        else:
            dia_semana = prev_volta.weekday()  # Pega o dia da semana
            prev_volta = prev_volta.strftime("%d/%m/%Y")
            prev_volta = f'*Previsão de volta:* {prev_volta} ({dias_da_semana[dia_semana]})'

        # Formata o nome dos colaboradores em maiúsculas e com o sobrenome abreviado
        nome_inicial_c1 = colab1.split()[0]
        sobrenome_abrev_c1 = colab1.split()[1][0] + '.'
        nome_c1 = (nome_inicial_c1 + ' ' + sobrenome_abrev_c1).upper()

        nome_inicial_c2 = colab2.split()[0]
        sobrenome_abrev_c2 = colab2.split()[1][0] + '.'
        nome_c2 = (nome_inicial_c2 + ' ' + sobrenome_abrev_c2).upper()

        # Montando o bloco de mensagem com as informações formatadas
        bloco_msg = f'*-> {nome_c1} e {nome_c2}*\n'
        bloco_msg += f'*N° Protocolo:* {num_protocolo}\n'
        bloco_msg += f'*Tarefas:* {tarefa}\n'
        bloco_msg += f'*Viatura:* {placa_viatura}\n'
        bloco_msg += f'*Observações:* _{observacao}_\n'
        bloco_msg += f'{prev_volta}\n'

        # Escreve o bloco de mensagem no arquivo de saída
        mensagem.write(bloco_msg + '\n')

