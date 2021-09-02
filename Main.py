from Google import Create_Service
import os
import argparse
import re
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
parser = argparse.ArgumentParser(description='Spreadsheet copy')
parser.add_argument('target', metavar='target', type=str, help='the target spreadsheet')
parser.add_argument('source', metavar='source', type=str, help='the original spreadsheet')


args = parser.parse_args()
arg_font = parser.parse_args()
source_sheet_id = arg_font.source
target_sheet_id = args.target
 
worksheet_id = '0'

new_target_sheet_id = ''

worksheet_list = []
specs = '^[Ss][Pp][Ee][Cc]\w?'
feature = '^[Ff][Ee][Aa][Tt][Uu][Rr][Ee][Ss]?$'
experiences = '^[Ee][Xx][Pp][Ee][Rr][Ii][Ee][Nn][Cc][Ee]\w?'
network = '^[Ee][Xx][Pp][Ee][Rr][Ii][Ee][Nn][Cc][Ee] [Nn][Ee][Tt][Ww][Oo][Rr][Kk]\w?'
hw_info = '^[Hh][Ww] [Ii][Nn][Ff][Oo]\w?'

clearConsole = lambda: os.system('cls' if os.name in ('nt', 'dos') else 'clear')


def criarServico():
    print("criando serviço")

def list_worksheet():
    mySpreadsheet =service.spreadsheets().get(
    spreadsheetId = source_sheet_id
    ).execute()
    
    print("Listando worksheets da planilha fonte")
    for sheet in mySpreadsheet['sheets']:
        if re.match(specs,sheet['properties']['title']) or re.match(experiences,sheet['properties']['title']) or re.match(network,sheet['properties']['title']) or re.match(feature,sheet['properties']['title']) or re.match(hw_info,sheet['properties']['title']):
            worksheet_list.append(str(sheet['properties']['sheetId']))
            print(f"{sheet['properties']['title']} ({sheet['properties']['sheetId']})")


def spreadsheet_copier(idWorksheed):
    service.spreadsheets().sheets().copyTo(
        spreadsheetId = source_sheet_id,
        sheetId = idWorksheed,
        body = {'destinationSpreadsheetId' : target_sheet_id}
    ).execute()


list_worksheet()
for i in range (len(worksheet_list)):
    print("Copiando worksheet de ID " + worksheet_list[i] + " para planilha alvo")
    spreadsheet_copier(worksheet_list[i])

