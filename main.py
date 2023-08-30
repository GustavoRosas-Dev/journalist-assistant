# Autor: Gustavo Rosas
# GitHub: https://github.com/GustavoRosas-Dev/journalist-assistant
# Data: 2023-07-20

# Este script requer configuração básica antes da execução. Consulte o README para obter instruções completas.
# https://github.com/GustavoRosas-Dev/journalist-assistant

#region LIBRARIES (combined)

#region LIBRARIES (script)

from __future__ import print_function
import os
import os.path  # Importa o módulo os.path para manipulação de caminhos de arquivo
import sys

from google.auth.transport.requests import Request  # Importa a classe Request do módulo google.auth.transport.requests
from google.oauth2.credentials import Credentials  # Importa a classe Credentials do módulo google.oauth2.credentials
from google_auth_oauthlib.flow import InstalledAppFlow  # Importa a classe InstalledAppFlow do módulo google_auth_oauthlib.flow
from googleapiclient.discovery import build  # Importa a função build do módulo googleapiclient.discovery
from googleapiclient.errors import HttpError  # Importa a classe HttpError do módulo googleapiclient.errors

import pyautogui # PopUp ao terminar (substituir pelo sg.PopUp do PySimpleGUI)

#endregion

#region LIBRARIES (UI)

import PySimpleGUI as sg
import re
import time
import webbrowser
import threading

#endregion

#endregion

#region LISTS (script)

paginas = [] # OK

linhas_vazias = [] # OK

total_links_encontrados = [] # OK

nome_itens_nao_encontrados = [] # OK

#endregion

#region DEFS (script)

# Sempre precisam ser executadas (a cada requisição)
#region AUTHORIZATION DEFS
def credentials_sheets():
    global creds  # Declara a variável creds como global

    creds = None
    SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets']

    # O arquivo token.json armazena o acesso e os tokens de atualização do usuário, e é
    # criado automaticamente quando o fluxo de autorização é concluído pela primeira vez.
    if os.path.exists('token-sheets.json'): # SE EXISTIR

        # TENTE USAR ESTES TOKENS:
        try:
            creds = Credentials.from_authorized_user_file('token-sheets.json', SCOPES_SHEETS)

        # SE O TOKEN DO SHEETS TIVER EXPIRADO...
        except:
            # aqui o script tentou acessar o Google com o token e não conseguiu.

            def popup_estilizado():
                # Definir o layout personalizado do popup
                layout = [
                    [sg.Text("Os Tokens do Google expiraram!\nClique em 'Reiniciar' e autentique-se novamente.",
                             font=(estilo_fonte, 10), text_color=text_color_erro, background_color=background_color_erro)],
                    [sg.Button("Reiniciar", key="-REINICIAR-", button_color=button_color_erro)]
                ]

                # Criar a janela do popup
                return sg.Window("Acesso negado", layout, keep_on_top=True,
                     icon="Logo/logo-journalist-oficial.ico", background_color=background_color_erro)

            # CRIA UMA JANELA QUE SIMULA UM POPUP (Porque: Pois o PySimpleGUI não permite estilização,
            # de forma nativa, em sg.Popups) P/ TER UM BOTÃO COM TEXTO PERSONALIZADO.
            popup = popup_estilizado()

            # Exibir o popup personalizado
            event, values = popup.read()
            popup.close()

            # Verificar o evento do botão
            if event == "-REINICIAR-":
                # Ação quando o botão "OK" é clicado

                #region APAGANDO OS TOKEN EXPIRADOS
                if os.path.exists('token-sheets.json'):
                    os.remove(path='token-sheets.json')  # REMOVE O DO TOKEN-SHEETS
                    try:
                        os.remove(path='token-drive.json')  # REMOVE O DO TOKEN-DRIVE
                    except:
                        pass
                #endregion

                #region REINICIANDO O PROGRAMA

                # Fechar a janela atual
                window.close()

                python = sys.executable  # Armazena o caminho para o interpretador Python atualmente em execução

                current_file = os.path.abspath(__file__)  # Obtém o caminho absoluto do arquivo Python atual

                os.chdir(os.path.dirname(
                    current_file))  # Altera o diretório de trabalho atual para o diretório onde o arquivo Python está localizado

                os.system(
                    python + " " + current_file)  # Executa um novo processo do Python, passando o caminho do interpretador e do arquivo atual como argumentos

                sys.exit()  # Encerra o programa original (anterior), NÃO o programa atual.
                #endregion

    # Se não houver credenciais disponíveis (válidas), permita que o usuário faça login.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials-sheets.json', SCOPES_SHEETS)
            creds = flow.run_local_server(port=0)
        # Salve as credenciais para a próxima execução
        with open('token-sheets.json', 'w') as token:
            token.write(creds.to_json()) # Converte o token que foi escrito p/ JSON

    return SCOPES_SHEETS

def credentials_drive():
    global creds  # Declara a variável creds como global

    creds = None  # Inicializa as credenciais como None
    SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive'] # Define os escopos de acesso para a API do Sheets

    if os.path.exists('token-drive.json'):
        try:
            creds = Credentials.from_authorized_user_file('token-drive.json', SCOPES_DRIVE)

        # SE O TOKEN DO DRIVE TIVER EXPIRADO...
        except:
            # aqui o script tentou acessar o Google com o token e não conseguiu.

            def popup_estilizado():
                # Definir o layout personalizado do popup
                layout = [
                    [sg.Text("Os Tokens do Google expiraram!\nClique em 'Reiniciar' e autentique-se novamente.",
                             font=(estilo_fonte, 10), text_color=text_color_erro,
                             background_color=background_color_erro)],
                    [sg.Button("Reiniciar", key="-REINICIAR2-", button_color=button_color_erro)]
                ]

                # Criar a janela do popup
                return sg.Window("Acesso negado", layout, keep_on_top=True,
                                 icon="Logo/logo-journalist-oficial.ico", background_color=background_color_erro)

            # CRIA UMA JANELA QUE SIMULA UM POPUP (Porque: Pois o PySimpleGUI não permite estilização,
            # de forma nativa, em sg.Popups) P/ TER UM BOTÃO COM TEXTO PERSONALIZADO.
            popup = popup_estilizado()

            # Exibir o popup personalizado
            event, values = popup.read()
            popup.close()

            # Verificar o evento do botão
            if event == "-REINICIAR2-":

                # region APAGANDO OS TOKEN EXPIRADOS
                if os.path.exists('token-drive.json'):
                    os.remove(path='token-drive.json')  # REMOVE O DO TOKEN-DRIVE
                    try:
                        os.remove(path='token-sheets.json')  # REMOVE O DO TOKEN-SHEETS
                    except:
                        pass
                # endregion

                # Fechar a janela atual
                window.close()

                # region REINICIANDO O PROGRAMA
                python = sys.executable
                current_file = os.path.abspath(__file__)
                os.chdir(os.path.dirname(current_file))
                os.system(python + " " + current_file)
                sys.exit()
                # endregion

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials-drive.json', SCOPES_DRIVE)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token-drive.json', 'w') as token:
            token.write(creds.to_json()) # Converte o token que foi escrito p/ JSON

    return SCOPES_DRIVE

#endregion

#region ACTION DEFS

def get_paginas_planilha():
    global paginas

    try:
        # Cria o serviço Sheets usando as credenciais
        service = build('sheets', 'v4', credentials=creds)
        # print('debug service', service)

        SAMPLE_SPREADSHEET_ID = id_planilha_google_sheets
        #print('debug SAMPLE_SPREAD...', SAMPLE_SPREADSHEET_ID)

        # Chama a API do Sheets
        sheet = service.spreadsheets()
        #print('debug service', service)
        spreadsheet = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID).execute()
        #print('debug spreadsheet', spreadsheet)
        sheets = spreadsheet.get('sheets', [])
        #print('debug sheets', sheets)

        # Percorre todas as páginas (abas) da planilha
        for s in sheets:
            sheet_title = s['properties']['title']
            #print('debug sheet_title', sheet_title)
            paginas.append(sheet_title)
            #print('debug páginas:', paginas)

        return paginas

    except HttpError as err:
        print(err)


def get_qtd_itens_sheets():
    global total_itens
    global linhas_vazias
    global soma_linhas_vazias

    values = []  # Inicializa a variável 'values' com uma lista vazia

    try:
        # Cria o serviço Sheets usando as credenciais
        service = build('sheets', 'v4', credentials=creds)
        #print('debug service', service)

        SAMPLE_SPREADSHEET_ID = id_planilha_google_sheets

        # Chama a API do Sheets
        sheet = service.spreadsheets()
        # print('debug service', service)
        spreadsheet = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID).execute()
        # print('debug spreadsheet', spreadsheet)
        sheets = spreadsheet.get('sheets', [])
        # print('debug sheets', sheets)

        # AQUI ESTA O ERRO:
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        #print('debug result:', result)
        values = result.get('values', [])
        #print('debug values:', values)

        # Se não tiver nada na planilha
        if not values:
            # print('Nenhum dado encontrado.')
            return

        # Do contrário
        else:
            # Percorra todas as linhas de valores
            for row in values:
                # Verifique se a linha está vazia
                if len(row) == 0:

                    linhas_vazias.append(1)
                    continue  # Pula para a próxima iteração do loop se a linha estiver vazia

                # E mostre o valor de cada linha
                #print(row[0])  # Supondo que o nome do arquivo esteja na primeira coluna (índice 0)

            # LINHAS VAZIAS
            for vazio in linhas_vazias:
                soma_linhas_vazias += vazio

            # print('\nLinhas vazias:', soma_linhas_vazias)

    except HttpError as err:
        print(err)

    if len(values) != 0:
        total_itens = len(values)
        #print('\nTOTAL ITENS:', total_itens)
    else:
        error_popup('AVISO: Não encontrei nenhum item na planilha.\nEncerrando...')
        sys.exit() # Encerra o programa

    total_itens = len(values)

    try:
        total_itens = len(values) - linhas_vazias[0]
    except:
        pass


def get_item_name_sheets():
    global item_name

    try:
        # Cria o serviço Sheets usando as credenciais
        service = build('sheets', 'v4', credentials=creds)
        #print('debug service', service)

        SAMPLE_SPREADSHEET_ID = id_planilha_google_sheets

        # Chama a API do Sheets
        sheet = service.spreadsheets()
        # print('debug service', service)
        spreadsheet = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID).execute()
        # print('debug spreadsheet', spreadsheet)
        sheets = spreadsheet.get('sheets', [])
        # print('debug sheets', sheets)

        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])

        # Verifique se há valores na lista
        if not values:
            #print('Nenhum dado encontrado.')
            return

        # Verifique se a linha está vazia
        if len(values[controlador_item]) == 0:
            #print('\nLinha vazia encontrada\n')
            return

        item_name = values[controlador_item][0]
        #print('\nPróximo item:', item_name)

    except HttpError as err:
        print(err)


def search_on_drive_ANTIGO(item_name):
    global web_view_link
    global item_encontrado
    global total_links_encontrados

    try:
        service = build('drive', 'v3', credentials=creds)

        # Call the Drive v3 API
        results = service.files().list(
            q=f"name contains '{item_name}'",
            pageSize=10,
            fields="nextPageToken, files(id, name, webViewLink)"
        ).execute()

        items = results.get('files', [])

        # A lista items está vazia? Nenhum item foi encontrado.
        if not items:
            item_encontrado = False
            web_view_link = None
            #print(f'Nenhum arquivo com o nome "{item_name}" foi encontrado.')  # Imprime uma mensagem se nenhum arquivo for encontrado

            # Adicionando o nome do item que não foi encontrado na lista (se ele não estiver lá)
            if item_name not in nome_itens_nao_encontrados:
                nome_itens_nao_encontrados.append(item_name)

        else:
            item_encontrado = True
            total_links_encontrados.append(1)

            file = items[0]
            file_name = file['name']
            file_id = file['id']

            web_view_link = file['webViewLink']
            # print(f'Arquivo "{file_name}" encontrado.'
            #       f'- Segue o link: {web_view_link}')

            return web_view_link  # Retorna o link de visualização do arquivo encontrado


    except HttpError as error:
        if error.resp.status == 400:
            print("Ocorreu um erro na solicitação. Verifique os parâmetros fornecidos.")
        else:
            print(f"Ocorreu um erro: {error}")

def get_child_folder_links_1(parent_folder_url):
    try:
        service = build('drive', 'v3', credentials=creds)

        # Obtém o ID da pasta-pai a partir do URL fornecido
        parent_folder_id = parent_folder_url.split('/')[-1]

        # Busca os arquivos que possuem a pasta-pai definida como a pasta desejada
        results = service.files().list(
            q=f"'{parent_folder_id}' in parents",
            pageSize=10,
            fields="files(name, webViewLink)"
        ).execute()

        items = results.get('files', [])

        # Verifica se foram encontrados arquivos filhos
        if not items:
            print('\nNenhuma pasta encontrada.')
        else:
            # print('Pastas filhas encontradas:')
            for item in items:
                file_name = item['name']
                web_view_link = item['webViewLink']
                # print(f'Nome: {file_name} - Link: {web_view_link}')

    except HttpError as error:
        if error.resp.status == 400:
            print("Ocorreu um erro na solicitação. Verifique os parâmetros fornecidos.")
        else:
            print(f"Ocorreu um erro: {error}")

def search_on_drive_1(item_name):
    global web_view_link
    global item_encontrado
    global total_links_encontrados

    try:
        service = build('drive', 'v3', credentials=creds)

        # Call the Drive v3 API
        results = service.files().list(
            q=f"name contains '{item_name}' and '{id_folder_google_drive}' in parents",
            pageSize=10,
            fields="nextPageToken, files(id, name, webViewLink)"
        ).execute()

        items = results.get('files', [])

        # A lista items está vazia? Nenhum item foi encontrado.
        if not items:
            item_encontrado = False
            web_view_link = None
            # print(f'Nenhum arquivo com o nome "{item_name}" foi encontrado.')  # Imprime uma mensagem se nenhum arquivo for encontrado

            # Adicionando o nome do item que não foi encontrado na lista (se ele não estiver lá)
            if item_name not in nome_itens_nao_encontrados:
                nome_itens_nao_encontrados.append(item_name)

        else:
            item_encontrado = True
            total_links_encontrados.append(1)

            file = items[0]
            file_name = file['name']
            file_id = file['id']

            web_view_link = file['webViewLink']
            # print(f'Arquivo "{file_name}" encontrado.'
            #       f'- Segue o link: {web_view_link}')

            return web_view_link  # Retorna o link de visualização do arquivo encontrado


    except HttpError as error:
        if error.resp.status == 400:
            print("Ocorreu um erro na solicitação. Verifique os parâmetros fornecidos.")
        else:
            print(f"Ocorreu um erro: {error}")


def get_child_folder_links(parent_folder_url):
    child_folder_links = []  # Lista para armazenar os links das pastas filhas

    SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive']
    creds = Credentials.from_authorized_user_file('token-drive.json', SCOPES_DRIVE)

    try:
        service = build('drive', 'v3', credentials=creds)
        #print('DEBUG SERVICE:', service)

        # Obtém o ID da pasta-pai a partir do URL fornecido
        parent_folder_id = parent_folder_url.split('/')[-1]
        #print('DEBUG PARENT_FOLDER_ID:', parent_folder_id)


        # ERRO AQUI: Busca os arquivos que possuem a pasta-pai definida como a pasta desejada
        results = service.files().list(
            q=f"'{parent_folder_id}' in parents",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            fields="nextPageToken, files(id, name, mimeType, webViewLink)").execute()

        items = results.get('files', [])

        if not items:
            print('Nenhum arquivo encontrado.')
            return

        # print('Files:')
        for item in items:
            name = item['name']
            item_type = 'Folder' if item['mimeType'] == 'application/vnd.google-apps.folder' else 'File'
            web_view_link = item['webViewLink']
            item_id = item['id']
            if item_type == 'Folder': # Se o item for do tipo folder
                child_folder_links.append(web_view_link)  # Adiciona o link na lista
            # print(f'{name} (ID: {item_id} - {item_type}: {web_view_link})')

    except HttpError as error:
        if error.resp.status == 400:
            error_popup("Erro na solicitação: Verifique os parâmetros fornecidos.")
        else:
            #print(f"Ocorreu um erro: {error}")

            credentials_error_popup() # Exiba um popup de erro de credenciais

            #region DELETA OS TOKENS DA PASTA RAIZ
            token_sheets_path = 'token-sheets.json'  # Caminho para o arquivo do token do Google Sheets
            token_drive_path = 'token-drive.json'  # Caminho para o arquivo do token do Google Drive

            if os.path.exists(token_sheets_path):
                os.remove(token_sheets_path)
                #print(f'Arquivo {token_sheets_path} removido')
            else:
                pass

            if os.path.exists(token_drive_path):
                os.remove(token_drive_path)
                #print(f'Arquivo {token_drive_path} removido')
            else:
                pass

            #endregion

    return child_folder_links  # Retorna a lista de links das pastas filhas

def search_on_drive(item_name, parent_folder_url):
    global web_view_link
    global item_encontrado
    global total_links_encontrados

    # print('DEBUG 1')
    child_folder_links = get_child_folder_links(parent_folder_url)  # Obtém os links das pastas filhas
    # print('CHILD FOLDER LINKS:', child_folder_links)
    # print('DEBUG 2')

    try:
        service = build('drive', 'v3', credentials=creds)

        item_encontrado = False  # Variável para indicar se o item foi encontrado

        for child_folder_link in child_folder_links:
            # Obtém o ID da pasta-filha a partir do link
            child_folder_id = child_folder_link.split('/')[-1]
            #print('DEBUG child_folder_id', child_folder_id)

            # Realiza a busca dentro da pasta-filha
            results = service.files().list(
                q=f"name contains '{item_name}' and '{child_folder_id}' in parents",
                pageSize=100,
                fields="nextPageToken, files(id, name, webViewLink)",
                supportsAllDrives = True,
                includeItemsFromAllDrives = True,
            ).execute()

            # if results['files'] != []:
            #     print('DEBUG Results', results)

            items = results.get('files', [])

            # A lista items está vazia? Nenhum item foi encontrado.
            if items:
                item_encontrado = True
                total_links_encontrados.append(1)

                file = items[0]
                file_name = file['name']
                file_id = file['id']

                web_view_link = file['webViewLink']
                # print(f'Arquivo "{file_name}" encontrado na pasta {child_folder_link}.'
                #       f'- Segue o link: {web_view_link}')

                return web_view_link  # Retorna o link de visualização do arquivo encontrado

        # O item não foi encontrado em nenhuma pasta filha
        if not item_encontrado:
            # print(f'Nenhum arquivo com o nome "{item_name}" foi encontrado nas pastas filhas.')


            # Adicionando o nome do item que não foi encontrado na lista (se ele não estiver lá)
            if item_name not in nome_itens_nao_encontrados:
                nome_itens_nao_encontrados.append(item_name)

    except HttpError as error:
        if error.resp.status == 400:
            print("Ocorreu um erro na solicitação. Verifique os parâmetros fornecidos.")
        else:
            print(f"Ocorreu um erro: {error}")

def paste_link_on_sheets():
    try:
        # Cria o serviço Sheets usando as credenciais
        service = build('sheets', 'v4', credentials=creds)
        #print('debug service', service)

        SAMPLE_SPREADSHEET_ID = id_planilha_google_sheets

        # Chama a API do Sheets
        sheet = service.spreadsheets()
        # print('debug service', service)
        spreadsheet = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID).execute()
        # print('debug spreadsheet', spreadsheet)
        sheets = spreadsheet.get('sheets', [])
        # print('debug sheets', sheets)
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        #print(len(values))

        found_item = False  # Flag para indicar se o item foi encontrado

        # Percorre cada linha da matriz 'values' (p/ localizar a célula com o item_name)
        for row in range(len(values)):
            # Percorre cada coluna da linha atual da matriz 'values'
            for col in range(len(values[row])):
                # Verifica se o elemento atual é igual ao 'item_name'
                if values[row][col] == item_name:
                    # Se o item_name for encontrado
                    if item_encontrado:
                        found_item = True

                        # Obter o intervalo de células a serem atualizadas
                        coluna = coluna_linha[0]
                        cell_range = f'{pagina_escolhida}!{coluna}{row + int(linha_escolhida)}'

                        # Definir a URL e o nome do link
                        url = web_view_link
                        name = item_name

                        # Criar a fórmula do link
                        link_formula = (f'=HYPERLINK("{url}"; "{name}")')

                        # Criar o corpo da solicitação
                        request_body = {
                            'range': cell_range,
                            'values': [[link_formula]],
                            'majorDimension': 'ROWS',
                        }

                        # Atualizar a célula com a fórmula do link
                        response = sheet.values().update(
                            spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range=cell_range,
                            valueInputOption='USER_ENTERED',
                            body=request_body
                        ).execute()

                        # Imprimir mensagem de sucesso
                        # print('Célula atualizada com sucesso!')

                    # Se o item NÃO foi encontrado
                    else:
                        continue

    except HttpError as err:
        print(err)


#endregion

#endregion

#region DEFS (UI)
def close_success_popup(window):
    time.sleep(4)  # Espera por 2 segundos
    window.close()

def make_success_popup():
    # Cores para o popup de sucesso
    text_color_sucesso = '#008000'
    background_color_sucesso = '#D6FFD6'

    # Configurações do popup de sucesso
    layout = [
        [sg.Text('Sucesso ao conectar', font=(estilo_fonte, 12), text_color=text_color_sucesso, background_color=background_color_sucesso)],
    ]
    window = sg.Window('Journalist Assistant', layout, finalize=True, keep_on_top=True, no_titlebar=True, background_color=background_color_sucesso, relative_location=(0,-220))

    # Define o tempo de exibição do popup (em milissegundos)
    timeout = 4000

    # Loop de eventos
    while True:
        event, values = window.read(timeout=timeout)

        # Verifica se o tempo de exibição foi atingido ou se o usuário fechou a janela
        if event == sg.TIMEOUT_EVENT or event == sg.WINDOW_CLOSED:
            break

    window.close()

def make_final_sucess_popup():

    # Função anônima (lambda) que atualiza o texto do botão 'Clique aqui para iniciar'
    (lambda window: window['-BTN_INICIAR-'].update('Clique aqui para iniciar', disabled=False,
                                                   button_color='#7186c7'))(window2)
    (lambda window: window['-TOTAL_ITENS_UI-'].update('0'))(window2)
    (lambda window: window['-ITENS_RESTANTES_UI-'].update('0'))(window2)
    window2.refresh()  # Força a atualização imediata da interface gráfica

    # Cores para o popup de sucesso
    text_color_sucesso = '#008000'
    background_color_sucesso = '#D6FFD6'

    # Configurações do popup de sucesso
    layout = [
        [sg.Text('Terminei de colar os links na planilha =D', font=(estilo_fonte, 12), text_color=text_color_sucesso, background_color=background_color_sucesso)],
    ]
    window = sg.Window('Journalist Assistant', layout, finalize=True, keep_on_top=True, no_titlebar=True,
                       background_color=background_color_sucesso, relative_location=(0,-240))

    # Define o tempo de exibição do popup (em milissegundos)
    timeout = 4000

    # Loop de eventos
    while True:
        event, values = window.read(timeout=timeout)

        # Verifica se o tempo de exibição foi atingido ou se o usuário fechou a janela
        if event == sg.TIMEOUT_EVENT or event == sg.WINDOW_CLOSED:
            break

    window.close()

    # Reiniciar interface (desativado)
    '''#region REINICIAR INTERFACE (REBOOT)

    # Fechar a janela atual
    window.close()
    window2.close()

    python = sys.executable  # Armazena o caminho para o interpretador Python atualmente em execução

    current_file = os.path.abspath(__file__)  # Obtém o caminho absoluto do arquivo Python atual

    os.chdir(os.path.dirname(
        current_file))  # Altera o diretório de trabalho atual para o diretório onde o arquivo Python está localizado

    os.system(
        python + " " + current_file)  # Executa um novo processo do Python, passando o caminho do interpretador e do arquivo atual como argumentos

    sys.exit()  # Encerra o programa original (anterior), NÃO o programa atual.
    # endregion'''


def error_popup(text):
    # Cores para o popup de erro

    # Configurações do popup de sucesso
    layout = [
        [sg.Text(str(text), font=(estilo_fonte, 12), text_color=text_color_erro,
                     background_color=background_color_erro)],
    ]
    window = sg.Window('Journalist Assistant', layout, finalize=True, keep_on_top=True, no_titlebar=True, background_color=background_color_erro, relative_location=(0,-220))

    # Define o tempo de exibição do popup (em milissegundos)
    timeout = 4000

    # Loop de eventos
    while True:
        event, values = window.read(timeout=timeout)

        # Verifica se o tempo de exibição foi atingido ou se o usuário fechou a janela
        if event == sg.TIMEOUT_EVENT or event == sg.WINDOW_CLOSED:
            break

    window.close()

def make_final_error_popup():
    # Cores para o popup de erro

    # Configurações do popup de sucesso
    layout = [
        [sg.Text('Opsss! Algo está errado.\nRevise os links informados e tente novamente.', font=(estilo_fonte, 12), text_color=text_color_erro, background_color=background_color_erro)],
    ]
    window = sg.Window('Journalist Assistant', layout, finalize=True, keep_on_top=True, no_titlebar=True, background_color=background_color_erro, relative_location=(0,-220))

    # Define o tempo de exibição do popup (em milissegundos)
    timeout = 4000

    # Loop de eventos
    while True:
        event, values = window.read(timeout=timeout)

        # Verifica se o tempo de exibição foi atingido ou se o usuário fechou a janela
        if event == sg.TIMEOUT_EVENT or event == sg.WINDOW_CLOSED:
            break

    window.close()

def credentials_error_popup():
    # Cores para o popup de erro

    # Configurações do popup de sucesso
    layout = [
        [sg.Text('Opsss! Parece que suas credenciais do Google expiraram.', font=(estilo_fonte, 12), text_color=text_color_erro, background_color=background_color_erro)],
    ]
    window = sg.Window('Journalist Assistant', layout, finalize=True, keep_on_top=True, no_titlebar=True, background_color=background_color_erro, relative_location=(0,-220))

    # Define o tempo de exibição do popup (em milissegundos)
    timeout = 4000

    # Loop de eventos
    while True:
        event, values = window.read(timeout=timeout)

        # Verifica se o tempo de exibição foi atingido ou se o usuário fechou a janela
        if event == sg.TIMEOUT_EVENT or event == sg.WINDOW_CLOSED:
            break

    window.close()

#region Validações (letras e números)

#region Validar apenas letras
def validar_letras(texto):
    padrao = r'^[a-zA-Z]+$'
    return re.match(padrao, texto)

#endregion

#region Validar apenas números

def validar_numeros(texto):
    padrao = r'^[0-9]+$'
    return re.match(padrao, texto)
#endregion

#endregion

#endregion

#region APARÊNCIA (UI)

#region Logo

sg.set_global_icon('Logo/logo-journalist-oficial.ico')  # Substitua 'caminho_para_o_icone.ico' pelo caminho real para o arquivo de ícone

#endregion

#region Tema

sg.theme('BlueMono')  # Definindo o tema da interface

#endregion

#region Fonte

estilo_fonte = "Open"

#endregion

#region Paleta de Cores

titulo_campo = '#300138'
cor_titulo = '#4e5d8c'
cor_descritivo = '#71798c'

#endregion

#region Erro

text_color_erro='#FF0000'
background_color_erro='#FFD6D6'
button_color_erro='#FF0000'

#endregion

#region Sucesso

text_color_sucesso = '#008000'
background_color_sucesso = '#D6FFD6'
button_color_sucesso = '#008000'

#endregion

#endregion

#region LAYOUT "Janela Principal" (UI)

# region Frames

frame_pagina = [
    [sg.Text('Nome da página', font=(estilo_fonte, 10), text_color=cor_titulo)],
    [sg.DropDown('', size=(13, 1), font=(estilo_fonte, 9), background_color='white',
                 text_color='#7186c7', key='-PAGINA-')]
]

frame_coluna = [
    [sg.Text('Letra da coluna', font=(estilo_fonte, 10), text_color=cor_titulo)],
    [sg.Input(key='-COLUNA-', font=(estilo_fonte, 9), text_color='#7186c7',
              tooltip='Insira a letra da coluna onde estão os nomes para pesquisar', size=(13, 1))]
]

frame_linha = [
    [sg.Text('Número da linha', font=(estilo_fonte, 10), text_color=cor_titulo)],
    [sg.Input(key='-LINHA-', font=(estilo_fonte, 9), text_color='#7186c7',
              tooltip='Insira o nº da linha onde começam os nomes para pesquisar', size=(13, 1))]
]

frame_configs = [
    [sg.Frame('', frame_pagina, border_width=0), sg.Frame('', frame_coluna, border_width=0),
     sg.Frame('', frame_linha, border_width=0)],
    [sg.Button('Clique aqui para iniciar', font=(estilo_fonte, 10), expand_x=True, key='-BTN_INICIAR-', disabled=False)]
]

frame_contador = [
    [sg.Text(' ' * 10),  # Espaçamento entre os elementos
     sg.Text('Total itens:', font=(estilo_fonte, 10), text_color=cor_titulo),
     sg.Text('0', font=(estilo_fonte, 10), text_color=cor_titulo, key='-TOTAL_ITENS_UI-'),
     sg.Text(' ' * 5),  # Espaçamento entre os elementos
     sg.Text('Itens restantes:', font=(estilo_fonte, 10), text_color=cor_titulo),
     sg.Text('0', font=(estilo_fonte, 10), text_color=cor_titulo, key='-ITENS_RESTANTES_UI-')]
]

frame_relatorio = [
    [sg.Output(size=(49, 10), font=(estilo_fonte, 10), text_color=cor_titulo, key='-OUTPUT-')]
]

# endregion

# region Layout

layout2 = [
    [sg.Text('Conecte os links do Google antes de iniciar', font=(estilo_fonte, 9, "bold"), text_color=cor_titulo),
     sg.Button('', image_filename='images/unlinked ok-01.png', size=(16, 16), key='-BTN_JANELA_LINKS-',
               button_color='#aab6d3', border_width=0, tooltip='Clique aqui para conectar/ desconectar')],
    [sg.Frame('Configurações da Planilha', frame_configs, font=(estilo_fonte, 8), title_color=cor_descritivo, )],
    [sg.Frame('', frame_contador, border_width=0, font=(estilo_fonte, 8), title_color=cor_descritivo)],
    [sg.Frame('Relatório', frame_relatorio, border_width=0, font=(estilo_fonte, 8), title_color=cor_descritivo)],
    [sg.Text(expand_x=True), sg.Text("GitHub", font=(estilo_fonte, 9), enable_events=True, text_color=cor_titulo, key='-GITHUB_TEXT-'), sg.Image("images/GitHub-02.png", key='-GITHUB_ICON-', enable_events=True)]

]

# endregion

# endregion

window2 = sg.Window('Journalist Assistant', layout2, finalize=True, keep_on_top=True, grab_anywhere=True)

#region EVENTS "Janela Principal" (UI)

# region *** ORIENTAÇÃO - PASSO 01

print('PASSO 01:\n- Conecte os links do Google antes de iniciar.\n')

# endregion

while True:
    event2, values2 = window2.read()

    # region EVENTO - Botão "Fechar"

    if event2 == sg.WINDOW_CLOSED:
        break

    # endregion

    # region EVENTO - Botão '-BTN_JANELA_LINKS-'

    elif event2 == '-BTN_JANELA_LINKS-':

        # region Update image link (UNLINKED)

        window2['-BTN_JANELA_LINKS-'].update(image_filename='images/unlinked ok-01.png')

        # endregion

        #region Criação da 'window' (Janela Secundária)
        layout = [
        [sg.Text('URL Google Sheets:', font=(estilo_fonte, 8, "bold"), text_color=cor_titulo)],
        [sg.Input(key='-URL_SHEETS-', text_color='gray',
                  tooltip='Insira o URL de uma planilha do Google Sheets')],
        [sg.Text('URL Google Drive:', font=(estilo_fonte, 8, "bold"), text_color=cor_titulo)],
        [sg.Input(key='-URL_DRIVE-', text_color='gray',
                  tooltip='Insira o URL de uma pasta do Google Drive')],
        [sg.Button('Conectar', font=(estilo_fonte, 10), expand_x=True, key='-BTN_CONECTAR-')]
    ]

        window = sg.Window('Journalist Assistant', layout, keep_on_top=True, grab_anywhere=True)

        #endregion

        #region EVENTOS Janela Secundária

        while True:
            event, values = window.read()

            # region EVENTO - Botão "Fechar"

            if event == sg.WINDOW_CLOSED:
                break

            # endregion

            # region EVENTO - Botão "Conectar"

            elif event == '-BTN_CONECTAR-':

                # Variavéis que representam as URLS inseridas
                url_sheets = values['-URL_SHEETS-']
                url_drive = values['-URL_DRIVE-']


                # region DEF's (UI)
                def extrair_id_planilha(url_sheets):
                    pattern = r'https://docs.google.com/spreadsheets/d/([^/]+)'
                    match = re.search(pattern, url_sheets)
                    #print('debug match (antes) -->', match)
                    if match:
                        #print('debug match (depois)--> ', match)
                        #print('debug match.group(1) (depois)--> ', match.group(1))
                        return match.group(1)
                    return None


                def extrair_id_drive_folders(url_drive):
                    if 'https://drive.google.com/drive/u/0/' in url_drive:
                        # print('------------- U/0 -------------')
                        pattern = r'https://drive.google.com/drive/u/0/folders/([^/?]+)'
                        match = re.search(pattern, url_drive)
                        # print('debug match (antes) -->', match)
                        if match:
                            # print('debug match (depois)--> ', match)
                            # print('debug match.group(1) (depois)--> ', match.group(1))
                            return match.group(1)
                        return None

                    else:
                        # print('------------- normal ---------------')
                        pattern = r'https://drive.google.com/drive/folders/([^/?]+)'
                        match = re.search(pattern, url_drive)
                        # print('debug match (antes) -->', match)
                        if match:
                            # print('debug match (depois)--> ', match)
                            # print('debug match.group(1) (depois)--> ', match.group(1))
                            return match.group(1)
                        return None



                # endregion

                # Aqui acionamos as defs, ao mesmo tempo que armazenamos o valor do return nas respectivas variáveis
                id_planilha_google_sheets = extrair_id_planilha(url_sheets)
                id_folder_google_drive = extrair_id_drive_folders(url_drive)

                # Se a URL do Sheets existir:
                if id_planilha_google_sheets:
                    # Se a URL do Drive existir:
                    if id_folder_google_drive:

                        # print('DRIVE--->', id_folder_google_drive)
                        # print('SHEETS--->', id_planilha_google_sheets)

                        #region Ações

                        # region --- ESCOPO

                        # Escopo necessário para acessar o Google Sheets
                        #SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets']
                        # Escopo necessário para acessar o Google Drive
                        #SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive']

                        # endregion

                        # region VARIABLES (script)

                        total_itens = 0
                        item_name = None
                        web_view_link = None
                        item_encontrado = None  # Bool
                        lista_nome_itens_nao_encontrados = ''

                        soma_total_links_encontrados = 0
                        controlador_item = -1
                        soma_linhas_vazias = 0

                        # endregion

                        #region Autenticações (Sheets + Drive)

                        credentials_sheets()
                        credentials_drive()

                        #endregion

                        #region Obter as páginas
                        get_paginas_planilha()
                        # print('Páginas:', paginas)
                        # print('Quant. de páginas:', len(paginas))
                        # for pagina in paginas:
                        #     print('Página:', pagina)

                        #endregion

                        # region Espelhar páginas no dropdown

                        window2['-PAGINA-'].update(values=paginas)

                        # endregion

                        # region Mude a imagem (unliked) do ícone, para "linked".

                        window2['-BTN_JANELA_LINKS-'].update(image_filename='images/linked ok-01.png')

                        # endregion

                        print('\t███████████████████████\n\t███ Conectado com sucesso ███\n\t███████████████████████')

                        # region *** ORIENTAÇÃO - PASSO 02

                        print(
                            '\nPASSO 02:\n- Preencha todos os campos acima e, em seguida, clique no botão "Clique aqui para iniciar".')

                        # endregion

                        window.close()

                        # region Exiba o PopUp de sucesso

                        make_success_popup()

                        # endregion

                        # endregion

                    #region PopUp ERRO: Se o URL do Drive não existir...

                    else:
                        sg.popup("URL inválida.\nPor favor, insira uma URL válida do Google Drive.", keep_on_top=True,
                                 relative_location=(0, 0), title="Journalist Assistant", icon="Logo/logo-journalist-oficial.ico",
                                 font=(estilo_fonte, 10), text_color=text_color_erro,
                                 background_color=background_color_erro, button_color=button_color_erro)

                    #endregion

                #region PopUp ERRO: Se o URL do Sheets não existir...
                else:
                    sg.popup("URL inválida.\nPor favor, insira uma URL válida do Google Sheets.", keep_on_top=True,
                             relative_location=(0, 0), title="Journalist Assistant", icon="Logo/logo-journalist-oficial.ico",
                             font=(estilo_fonte, 10), text_color=text_color_erro,
                             background_color=background_color_erro, button_color=button_color_erro)

                #endregion

            #endregion

        #endregion

    #endregion

    # region EVENTO - Botão "Clique aqui para iniciar"

    elif event2 == '-BTN_INICIAR-':

        linha_escolhida = values2['-LINHA-']
        pagina_escolhida = values2['-PAGINA-']
        coluna_linha = values2['-COLUNA-']

        # Se os campos de pagina e coluna_linha estiverem preenchidos
        if pagina_escolhida and coluna_linha and linha_escolhida:
            # Função anônima (lambda) que atualiza o texto + a cor do botão 'Clique aqui para iniciar'
            (lambda window: window['-BTN_INICIAR-'].update('Aguarde! Estou trabalhando...', disabled=True,
                                                           button_color='#E0E0E0'))(window2)
            window2.refresh()  # Força a atualização imediata da interface gráfica

            # Verifica se foi uma letra que o usuário inseriu no campo 'Letra da coluna'
            if validar_letras(coluna_linha):
                # Verifica se foi um número que o usuário inseriu no campo 'Número da linha'
                if validar_numeros(linha_escolhida):

                    # region --- SAMPLES (script)

                    # ID e intervalo de uma planilha de exemplo
                    SAMPLE_SPREADSHEET_ID = '1q3WqsfepC0GIgp7KQZ5vK57DB55ZBM_ilMOEfYd3eBU'
                    SAMPLE_RANGE_NAME = f'{pagina_escolhida}!{coluna_linha[0]}{linha_escolhida}:{coluna_linha[0]}500'

                    # endregion

                    # region LÓGICA (script)

                    if __name__ == '__main__':

                        # region Autenticações (Sheets + Drive)
                        credentials_sheets()
                        credentials_drive()

                        # endregion

                        get_child_folder_links(id_folder_google_drive)

                        # A partir daqui já conseguimos acessar a variavel total_itens
                        get_qtd_itens_sheets()
                        #print('Total de itens a serem verificados:', (total_itens - linhas_vazias[0]) + 1)

                        # Função anônima (lambda) que atualiza o texto com a quantidade de itens
                        (lambda window: window['-TOTAL_ITENS_UI-'].update(total_itens))(window2)
                        window2.refresh()  # Força a atualização imediata da interface gráfica

                        # Aqui é onde devemos criar a lógica: Planilha-Drive-Planilha
                        inicio = 0
                        fim = total_itens

                        # Com o "fim" sendo igual ao "total de itens", temos o comprimento da lista, ou o total de nomes que deveremos encontrar links.
                        while inicio < fim:
                            inicio += 1
                            controlador_item += 1

                            # Função anônima (lambda) que atualiza o texto com a quantidade de itens
                            (lambda window: window['-ITENS_RESTANTES_UI-'].update((fim+1) - inicio))(window2)
                            window2.refresh()  # Força a atualização imediata da interface gráfica

                            # PLanilha
                            get_item_name_sheets()
                            #print('Control...', controlador_item)

                            # Drive
                            search_on_drive(item_name, id_folder_google_drive)

                            # Planilha
                            paste_link_on_sheets()

                        # Relatório (SOMAS)
                        print('\n\t█████████████████████████\n\t████ Confira o relatório abaixo ████\n\t█████████████████████████')

                        for link in total_links_encontrados:
                            soma_total_links_encontrados += link
                        print('\n- Links encontrados:', soma_total_links_encontrados)

                        print('- Linhas vazias:', soma_linhas_vazias)

                        # Exibindo a lista dos nomes que não foram encontrados links no Drive
                        print('- Itens que eu NÃO encontrei link:')
                        for nome_nao_encontrados in nome_itens_nao_encontrados:
                            if nome_nao_encontrados != None:
                                print('\t  -', nome_nao_encontrados)

                        # Se ENCONTROU 1 ou mais links
                        if soma_total_links_encontrados != 0:

                            #region PopUp do final

                            make_final_sucess_popup()

                            #endregion

                        # Se NÃO encontrou nenhum link
                        else:

                            # region PopUp Error do final

                            make_final_error_popup()

                            #endregion

                        # Final
                        # pyautogui.alert('Terminei de colar os links na planilha =D',
                        #                 title='Journalist Assistant')

                else:
                    # region PopUp ERRO 'Número da linha'
                    sg.popup(
                        f"A entrada fornecida '{linha_escolhida}' é inválida.\nPor favor, insira apenas números no campo 'Número da linha'.",
                        keep_on_top=True,
                        relative_location=(0, 0), title="Journalist Assistant", icon="Logo/logo.ico",
                        font=(estilo_fonte, 10), text_color=text_color_erro,
                        background_color=background_color_erro, button_color=button_color_erro)

                    # endregion

            else:
                # region PopUp ERRO 'Letra da coluna'
                sg.popup(
                    f"A entrada fornecida '{coluna_linha}' é inválida.\nPor favor, insira apenas letras no campo 'Letra da coluna'.",
                    keep_on_top=True,
                    relative_location=(0, 0), title="Journalist Assistant", icon="Logo/logo.ico",
                    font=(estilo_fonte, 10), text_color=text_color_erro,
                    background_color=background_color_erro, button_color=button_color_erro)
                # endregion

            # endregion

            # Fechar a janela secundária e sair do loop
            #window2.close()
            # break

        # Se NÃO estiverem preenchidos
        else:
            sg.popup("Todos os campos devem estar preenchidos para 'Começar'.",
                     keep_on_top=True,
                     relative_location=(0, -50), title="Journalist Assistant", icon="Logo/logo-journalist-oficial.ico",
                     font=(estilo_fonte, 10), text_color='#FF0000', background_color='#FFD6D6',
                     button_color='#FF0000', )

    # endregion

    #region EVENTO - Link do GitHub

    elif event2 == '-GITHUB_TEXT-' or '-GITHUB_ICON-':
        webbrowser.open_new_tab('https://github.com/GustavoRosas-Dev')

    #endregion

#endregion
