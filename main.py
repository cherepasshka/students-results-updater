import httplib2
import os
import googleapiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


# Авторизуемся и получаем service — экземпляр доступа к API
def connect(credentials_file):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        credentials_file,
        ['https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    return googleapiclient.discovery.build('sheets', 'v4', http=httpAuth)


def have_sheet(name, all_sheets):
    for sheet in all_sheets:
        if sheet['properties']['title'] == name:
            return True
    return False


def add_sheet(sheet_name, row_count, column_count):
    results = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet_name,
                            "gridProperties": {
                                "rowCount": row_count,
                                "columnCount": column_count
                            }
                        }
                    }
                }
            ]
        }
    ).execute()
    return results


def read(sheet_name, from_cell, to_cell):
    range = '{name}!{from_cell}:{to_cell}'.format(name=sheet_name, from_cell=from_cell, to_cell=to_cell)
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range,
        majorDimension='ROWS'
    ).execute()
    return values


def clear_sheet(sheet_name, start_cell, row_count, column_count):
    last_row = int(start_cell[1:]) + row_count - 1
    end_cell = '{letter}{row}'.format(letter=chr(ord(start_cell[0]) + column_count - 1), row=last_row)
    range = '{name}!{from_cell}:{end_cell}'.format(name=sheet_name, from_cell=start_cell, end_cell=end_cell)
    request = service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=range).execute()
    return request


def transpose(data):
    return [[data[i][j] for i in range(len(data))] for j in range(len(data[0]))]


def write(sheet_name, from_cell, data_to_write, dimension):
    if dimension == 'COLUMNS':
        data_to_write = transpose(data_to_write)
    
    last_row = int(from_cell[1:]) + len(data_to_write) - 1
    end_cell = '{letter}{row}'.format(letter=chr(ord(from_cell[0]) + len(data_to_write[0]) - 1), row=last_row)
    print(from_cell, end_cell)
    range = '{name}!{from_cell}:{end_cell}'.format(name=sheet_name, from_cell=from_cell, end_cell=end_cell)
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": range,
                 "majorDimension": "ROWS",
                 "values": data_to_write},
            ]
        }
    ).execute()
    return values


def prepare_sheet(sheet_name, sheet_list):
    if not have_sheet(sheet_name, sheet_list):
        add_sheet(sheet_name, 1000, 26)
    clear_sheet(sheet_name, 'A1', row_count=1000, column_count=26)


files = os.listdir()
if 'sheet_id.txt' not in files:
    print('Error: current directory should contain sheet_id.txt file with id of Google Sheet')
    exit(1)


# Файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'credentials.json'

with open('sheet_id.txt', 'r') as f:
    spreadsheet_id = f.read()

service = connect(CREDENTIALS_FILE)
spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
sheet_list = spreadsheet.get('sheets')

print('Service account has successfully connected to Google sheet')

working_sheet = sheet_list[0]['properties']['title']
prepare_sheet(working_sheet, sheet_list)

with open('out.txt', 'r') as f:
    data = f.read()
data = [x.split() for x in data.split('\n')]
write(working_sheet, 'A1', data, "ROWS")
