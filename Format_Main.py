from __future__ import print_function
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import time
import datetime
from dateutil.relativedelta import relativedelta
import json

from GSheet_Format_Func import *


SCOPES = [
    'https://www.googleapis.com/auth/drive.metadata',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
]


def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file("token.json",SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else: 
            flow= InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json","w") as token:
            token.write(creds.to_json())

    #MAKE A COPY OF THE FILE
    try:
        #access drive service
        service = build("drive","v3",credentials= creds)

        #Copy the file
        copied_file = {
            'name': 'Monthly Report Testing.xlsx',  # Name of the copied report
            'parents': ['1seyqMYdCx139_ngueNCYqFJ7KKSH6mTH']  # Destination folder ID
        }
        copied_file_response = service.files().copy(
            fileId='1Aj1m-20WtahHevbe9wmK5xU78kLQ4gfoa0sx1uLadFk', #Programmatic Access - Monthly Reports ID
            body=copied_file
        ).execute()

        if copied_file_response:
            print('File copied successfully.')
        else:
            print('File copy operation did not return a response.')
        
        service = build('sheets', 'v4', credentials=creds)
        copied_spreadsheet_id = copied_file_response['id']
        copied_spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=copied_spreadsheet_id).execute()
    
        ## Iterate through the sheets and print their sheetIds
        #for sheet in copied_spreadsheet_metadata['sheets']:
        #    sheet_id = sheet['properties']['sheetId']
        #    sheet_name = sheet['properties']['title']
        #    print(f"Sheet Name: {sheet_name}, Sheet ID: {sheet_id}")

        sheet_names_to_delete = [
            'Social Post - Span',
            'Social Post - Eng',
            'OO TikTok Export',
            'OO Twitch Export',
            'OO IG Export',
            'OO Fb Posts Export',
            'OO Fb Vids Export',
            'OO YT Export',
            'OO Twitter Export',
            'EG accounts',
            'OO Export (SPRINKLR)',
            'CC Social Export (SPRINKLR)',
            'Followers (CC)',
            'Impressions (CC)',
            'Engagement (CC)',
            'Video Views (CC)',
            'Video Views Trend (CC)']

        for name in sheet_names_to_delete:
            delete_sheet_by_name(copied_spreadsheet_id, name)

        # Get the list of sheets in the document
        sheet_metadata = service.spreadsheets().get(spreadsheetId=copied_spreadsheet_id).execute()
        sheets = sheet_metadata['sheets']

        for sheet_info in sheets:
            sheet_name = sheet_info['properties']['title']

            # Apply common operations to the sheet
            apply_common_operations_to_sheet(copied_spreadsheet_id, sheet_name)

            #change font size
            change_font_size(copied_spreadsheet_id, sheet_name, "A3", 12)

            # Wait for a minute before processing the next sheet
            time.sleep(40)

        print("Functions applied to all sheets successfully.")



    except HttpError as error:
        print(error)

if __name__ == "__main__":
    main()



    


