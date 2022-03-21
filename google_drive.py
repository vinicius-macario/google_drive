### SETUP 

#Bibliotecas
pip install --upgrade xlrd==2.0.1

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from google.oauth2 import service_account
from googleapiclient.discovery import build

import pandas as pd
import os
import io
from googleapiclient.http import MediaIoBaseDownload

from google.colab import drive

#Montar o Drive
drive.mount('/content/gdrive') # Uma janela será exibida para dar permissão

#Credenciais Google Drive
#É possível utilizar uma biblioteca do Google para acessar os arquivos sem as credenciais no Colab, mas nesse caso ele irá solicitar mais uma permissão.
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name('/content/gdrive/MyDrive/arquivo_credenciais.json', scope)
client = gspread.authorize(credentials)

service = build('drive', 'v3', credentials=credentials)

#####

#Listar aquivos na Pasta
topFolderId = 'XXXX'  #Id da pasta do Google Drive
items = []
pageToken = ""
while pageToken is not None:
    response = service.files().list(q="'" + topFolderId + "' in parents", pageSize=1000, pageToken=pageToken, fields="nextPageToken, files(id, name)").execute()
    items.extend(response.get('files', []))
    pageToken = response.get('nextPageToken')
    
#Dataframe com o ID e Nome do Arquivo
file_id = []
nome = []
for f in range(0, len(items)):
    file_id.append(items[f].get('id'))
    nome.append(items[f].get('name'))
arquivo = pd.DataFrame(list(zip(file_id, nome)), columns=['File ID', 'Nome'])

#Função pra baixar os arquivos
def baixar_planilha_colab(file_id, nome):
  request = service.files().get_media(fileId=file_id)
  fh = io.FileIO("/content/gdrive/MyDrive/pasta/{0}".format(nome), mode='wb')
  downloader = MediaIoBaseDownload(fh, request)
  done = False
  while done is False:
      status, done = downloader.next_chunk()
      print(nome)
      
#Baixar as planilhas
for index, row in arquivo.iterrows():
    file_id = row['File ID']
    nome = row['Nome']
    baixar_planilha_colab(file_id, nome)

#Função pra abrir os arquivos
def abrir_arquivo(path, nome_arquivo):
    df = pd.read_excel(path, sheet_name='Pasta1', usecols='B:O', skiprows=6, nrows=8 ) #Parâmetros para acessar somente parte da planilha
    #outros ajustes podem ser incluídos aqui
    return df
 
# Listar arquivos do diretorio, apliar a função e consolidar em um único DF
diretorio = '/content/gdrive/MyDrive/pasta/'
base_aux = pd.DataFrame()
for index, row in arquivo_aux.iterrows():
  nome_arquivo = row['Nome']
  path = diretorio + nome_arquivo
  resultado = abrir_arquivo(path, nome_arquivo)
  base_aux = pd.concat([base_aux, resultado])

#Salvar o dataframe final em um arquivo
base_aux.to_excel('/content/gdrive/MyDrive/arquivo_final.xlsx', index=False)
