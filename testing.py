import os
import time

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# after the slash D on the URL is the sheet ID
spreadsheet_id = "1_djs85b4rOoZiUE9qz_LJ0CnDHgEinPzJlyWM9-u_FI"
def set_cell_background_color(service, spreadsheet_id, sheet_info, cell_range, color):
    # Set the background color of a cell in the specified sheet
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_info['sheetId'],
                    "startRowIndex": int(cell_range[1:]) - 1,
                    "endRowIndex": int(cell_range[1:]),
                    "startColumnIndex": ord(cell_range[0]) - ord("A"),
                    "endColumnIndex": ord(cell_range[0]) - ord("A") + 1,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": {"green": color}}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
    ]

    body = {"requests": requests}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def main():
    credentials = None
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
    # Gotta put sheet first followed by exclamation mark then the cell range
    result = sheets.values().get(spreadsheetId=spreadsheet_id, range='A1:B1').execute()

    myPeopleList = []
    with open("names", "r") as file:
        for person in file.readlines():
            myPeopleList.append(person.strip('\n').lower())
    count = 0
    threshold = 57
    for row in range(2, 241):
        count += 1
        cell_value = (sheets.values().get(spreadsheetId=spreadsheet_id, range=f'A{row}').execute())['values'][0][0].lower()
        if cell_value in myPeopleList:
            sheets.values().update(
                spreadsheetId=spreadsheet_id,
                range=f'D{row}',
                valueInputOption='USER_ENTERED',
                body={'values': [['Present']]}
            ).execute()
            set_cell_background_color(service, spreadsheet_id, result['sheets'][0], f'D{row}',0.8)  # Setting cell background color to green

        print(count)
        if count == threshold:
            time.sleep(60)
            threshold += 57


if __name__ == '__main__':
    main()
