import os
import time

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def main(namesColLetter, populateColLetter, spreadSheetLink,populateText='Present'):
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
    # Gotta put sheet frist followed by exlamation mark then the cell range
    result = sheets.values().get(spreadsheetId=spreadsheet_id, range='A1:B1').execute()


    myPeopleList = []
    with open("names", "r") as file:
        for person in file.readlines():
            myPeopleList.append(person.strip('\n').lower())
    count = 0
    threshold = 57
    for row in range(2,241):
        count += 1
        if ((sheets.values().get(spreadsheetId=spreadsheet_id, range=f'{namesColLetter}{row}').execute())['values'][0][0]).lower() in myPeopleList:
            sheets.values().update(spreadsheetId=spreadsheet_id, range=f'{populateColLetter}{row}', valueInputOption='USER_ENTERED', body={'values': [[populateText]]}).execute()


        print(count)
        if count == threshold:
            time.sleep(60)
            threshold += 57


if __name__ == '__main__':
    linkToSpreadSheet = "https://docs.google.com/spreadsheets/d/1_djs85b4rOoZiUE9qz_LJ0CnDHgEinPzJlyWM9-u_FI/edit#gid=1909685657"
    main('A', 'D', linkToSpreadSheet,'Present')