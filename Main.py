from Google import Create_Service
import os
import argparse

CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
parser = argparse.ArgumentParser(description='Copia de planilhas')
parser.add_argument('target', metavar='target', type=str, help='informe planilha alvo')
parser.add_argument('source', metavar='source', type=str, help='informe planilha fonte')

args = parser.parse_args()
arg_font = parser.parse_args()
source_sheet_id = arg_font.source
target_sheet_id = args.target
 
worksheet_id = '0'

new_target_sheet_id = ''

listaWorksheet = []
clearConsole = lambda: os.system('cls' if os.name in ('nt', 'dos') else 'clear')


def criarServico():
    print("criando serviço")

def listarWorksheet():
    mySpreadsheet =service.spreadsheets().get(
    spreadsheetId = source_sheet_id
    ).execute()
    
    print("Listando worksheets da planilha fonte")
    for sheet in mySpreadsheet['sheets']:
        listaWorksheet.append(str(sheet['properties']['sheetId']))
        print(sheet['properties']['title'] + ' (' + str(sheet['properties']['sheetId']) + ")")


def copiarSpreadsheet(idWorksheed):
    service.spreadsheets().sheets().copyTo(
        spreadsheetId = source_sheet_id,
        sheetId = idWorksheed,
        body = {'destinationSpreadsheetId' : target_sheet_id}
    ).execute()


listarWorksheet()
for i in range (len(listaWorksheet)):
    print("Copiando worksheet de ID " + listaWorksheet[i] + " para planilha alvo")
    copiarSpreadsheet(listaWorksheet[i])
    

