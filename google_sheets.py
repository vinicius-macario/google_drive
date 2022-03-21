
#bibliotecas
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

#Credenciais
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name('/content/gdrive/MyDrive/arquivo_credenciais.json', scope)
client = gspread.authorize(credentials)

#Obter o conteúdo da planilha
gsheets = client.open_by_key('ID_DA_PLANILHA').worksheet('NOME_ABA') 
base_ = pd.DataFrame(gsheets.get_all_records())

#Inserir conteúdo
df.fillna('N/A', inplace=True)
gsheets.update('A1'.format(linha), df.values.tolist())

# Atualizar a partir da última linha preenchida 
linha = len(gsheets.col_values(1)) + 1
gsheets.update('A{}'.format(linha), df.values.tolist())
