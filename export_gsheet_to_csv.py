#!/usr/bin/python3
#
# Get your Google spread sheet on Gsuite
#
# Before using this script, please set your sheet id to the following variable.
#   GSHEET_ID  = 'your sheet id'
#
# Usage: export_gsheet_to_csv.py [-m (sheet name)] [-o output-csv-file-name]
#
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import csv
import sys
import getopt

# Google spreadsheet conf
GSHEET_ID   = ''
SHEET_NAME  = 'sheet1'
SHEET_RANGE = 'A1:C10'

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CREDENT= 'credentials.json'
TOKEN  = 'token.pickle'

def get_credential():
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = None
    if os.path.exists(TOKEN):
        with open(TOKEN, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENT, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token)
    if not creds:
        sys.exit("No credential")
    return creds

def get_spreadsheet(sheet_id, sheet_range_str, creds):
    sheet = build('sheets', 'v4', credentials=creds).spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=sheet_range_str).execute()
    values = result.get('values', [])
    if not values:
        sys.exit('No data found.')
    return values

def export_csv(csv_data_matrix, fname):
    with open(fname, 'w', encoding='shift_jis') as f:
        csv.writer(f).writerows(csv_data_matrix)


if __name__ == '__main__':
    output_csv = "sheet.csv"
    try:
        if sys.argv:
            opts, args = getopt.getopt(sys.argv[1:], 'm:o:')
            for o, a in opts:
                if o == "-o":
                    output_csv = a
                elif o == "-m":
                    SHEET_NAME = a
                else:
                    raise getopt.GetoptError("Invalid value")
    except getopt.GetoptError:
        sys.exit("Usage: " + os.path.basename(__file__) +
                 " [-m (sheet name)] [-o output-csv-file-name]")

    credentials = get_credential()
    sheet_data  = get_spreadsheet(GSHEET_ID, SHEET_NAME + "!" + SHEET_RANGE,
                                  credentials)
    export_csv(sheet_data, output_csv)
