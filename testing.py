import os
import time

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_sheet_id(service, spreadsheet_id, sheet_name):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', [])
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    return None

def main(namesColLetter, populateColLetter, spreadSheetLink, populateText='Present', sheet_name='MASTER SHEET'):
    credentials = None
    spreadsheet_id = spreadSheetLink.split('/')[5]
    if os.path.exists('token.json'):
        credentials = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(credentials.to_json())

    service = build('sheets', 'v4', credentials=credentials)
    sheets = service.spreadsheets()

    sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
    if sheet_id is None:
        print(f"Sheet with name '{sheet_name}' not found.")
        return


    myPeopleList = []
    with open("names", "r") as file:
        for person in file.readlines():
            eachPerson = person.strip('\n').strip(" ").split('@')[0].lower()
            myPeopleList.append(eachPerson)
            print(eachPerson)

    count = 0
    threshold = 41
    for row in range(2, 242):
        count += 1
        nameOrEmail = ((sheets.values().get(spreadsheetId=spreadsheet_id, range=f'{namesColLetter}{row}').execute())['values'][0][0]).strip(' ').split('@')[0].lower()
        print(nameOrEmail)

        if nameOrEmail in myPeopleList:
            sheets.values().update(spreadsheetId=spreadsheet_id, range=f'{populateColLetter}{row}', valueInputOption='USER_ENTERED', body={'values': [[populateText]]}).execute()

            # Set background color to green
            requests = [
                {
                    'updateCells': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': row - 1,
                            'endRowIndex': row,
                            'startColumnIndex': ord(populateColLetter) - ord('A'),
                            'endColumnIndex': ord(populateColLetter) - ord('A') + 1
                        },
                        'rows': [
                            {
                                'values': [
                                    {
                                        'userEnteredFormat': {
                                            'backgroundColor': {
                                                'green': 1.0,
                                                'blue': 0.0,
                                                'red': 0.0
                                            }
                                        }
                                    }
                                ]
                            }
                        ],
                        'fields': 'userEnteredFormat.backgroundColor'
                    }
                }
            ]
            sheets.batchUpdate(spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()

        print(count)
        if count == threshold:
            time.sleep(61)
            threshold += 39

if __name__ == '__main__':
    linkToSpreadSheet = "https://docs.google.com/spreadsheets/d/1_djs85b4rOoZiUE9qz_LJ0CnDHgEinPzJlyWM9-u_FI/edit#gid=1909685657"
    main('O', 'I', linkToSpreadSheet, 'Present')
