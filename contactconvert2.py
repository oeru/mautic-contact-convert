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

import sys
from collections import OrderedDict
import csv
import datetime
from validate_email import validate_email

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'
SHEET_ID = '17cB9E3Pj5syTlldrEG95JlyKoKh086H7vorP7pCz1Bk'
SHEET_RANGE = 'Sheet1!A2:D'
DELIMITER = ','
ENCLOSURE = '"'

SHEET_HEADINGS = ["email", "name", "first", "last"]
CSV_HEADINGS = ["First Name", "Last Name", "Email"]
FIELD_ORDER = [2, 3, 0]

#counter = 0

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
    credential_path = os.path.join(credential_dir,                                  'sheets.googleapis.com-python-quickstart.json')

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
    # open the file for storing the results - give it a useful Name
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H.%M.%S")
    path = "csv"
    filename = "contacts-" + timestamp + ".csv"
    fullname = path + "/" + filename
    print("Writing CSV to {}".format(fullname))
    with open(fullname, 'w') as csvfile:
        # first add the headers
        contacts = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)

        # first, write the header row
        contacts.writerow(CSV_HEADINGS)

        # get user credentials to access the spreadsheet
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
        service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
        result = service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=SHEET_RANGE).execute()

        # then print the data
        rows = result.get('values', [])
        if not rows:
            print('No data found.')
        else:
            print('User Data:')
            for row in rows:
                dodgy = False # flag rows with dodgy data
                contact = []
                # ok - quick and dirty
                for index in FIELD_ORDER:
                    try:
                        # email is the 7th field in the spreadsheet
                        # checking MX: validate_email(row[index].encode("utf-8").strip(), check_mx=True):
                        if index == 6 and not  validate_email(row[index].encode("utf-8").strip()):
                            dodgy = True
                            #row[index].append(" **invalid**")
                        contact.append(row[index].encode("utf-8").strip())
                    except:
                        contact.append("Null")
                        dodgy = True

                if dodgy:
                    print('**** Dodgy: {}'.format(contact))
                else:
                    print('{} {} ({})'.format(contact[0], contact[1], contact[2]))
                    contacts.writerow(contact)


if __name__ == '__main__':
    main()

# import the spread sheet and print it as a series of rows
# in the form:
# "title", "firstname", "lastname", "email", "organisation", "position", "website", "tags"
# where tags include "primary, cc1, cc2, cc3"

# Should disqualify rows without a valid email (do an email SMTP query to verify!)

