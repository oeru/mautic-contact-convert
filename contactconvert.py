#!/usr/bin/python

# use a Google Sheet to set up feeds to harvest
# contactconvert.py GOOGLE-SHEETS-URL (readable by bot)
#
## deprecated - requires install of https://pypi.python.org/pypi/validate_email
#
# from https://developers.google.com/sheets/quickstart/python

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Lists a set of people's details in a comma separated format
    """
    # get user credentials to access the spreadsheet
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    # set the ID of the spreadsheet
    spreadsheetId = '1gJZoh5iNFweWaAfdG2QiNvP6tlDUWBymJ6NOoRGKqG8'
    rangeName = 'Sheet1!A2:G2'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('First Column:')
        for row in values:
            print(row)

    rangeName = "Sheet1!A3:G"
    result = service.spreadsheets().values().get(
    spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('User Data:')
        for row in values:
            print(row)

if __name__ == '__main__':
    main()

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
#SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
#CLIENT_SECRET_FILE = 'client_secret.json'
#APPLICATION_NAME = 'Google Sheets API Python Quickstart'


#import os
#import sys
## work with JSON data
#import json
#from hashlib import md5
#import urllib2
#import BeautifulSoup
## interact with Google spreadsheets
#import gspread
## authenticate to access the Google api
#from oauth2client.service_account import ServiceAccountCredentials

#DOC_ID = "1gJZoh5iNFweWaAfdG2QiNvP6tlDUWBymJ6NOoRGKqG8"
#GOOGLE_API_KEY = "google-servicekey.json"
#SCOPE = []'https://spreadsheets.good.com/feeds']

#credentials = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_API_KEY, SCOPE)

#gs = gspread.authorize(credentials)

#sheet = gs.open_by_key(DOC_ID).sheet1

#all_content = sheet.get_all_values()

#print(all_content)

# import the spread sheet and print it as a series of rows
# in the form:
# "title", "firstname", "lastname", "email", "organisation", "position", "website", "tags"
# where tags include "primary, cc1, cc2, cc3"

# Should disqualify rows without a valid email (do an email SMTP query to verify!)

#STARTROW = 1    # 0 indexed
