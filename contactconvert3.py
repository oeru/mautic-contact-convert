#!/usr/bin/python

# use a Google Sheet to set up feeds to harvest
# contactconvert.py GOOGLE-SHEETS-URL (readable by bot)
#
## deprecated - requires install of https://pypi.python.org/pypi/validate_email
#
# from https://developers.google.com/sheets/quickstart/python
#
# Columns from Spreadsheet:
# a 0 Company
# b 1 Website
# c 2 Title
# d 3 First Name
# e 4 Last Name
# f 5 Position
# g 6 Email
# h 7 Notes
# i 8 Title
# j 9 First Name
# k 10 Last Name
# l 11 Position
# m 12 Email
# n 13 Title
# o 14 First Name
# p 15 Last Name
# q 16 Position
# r 17 Email
# s 18 Title
# t 19 First Name
# u 20 Last Name
# v 21 Position
# w 22 Email

from __future__ import print_function
import httplib2
import os

from googleapiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from oauth2client import file

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
SHEET_ID = '1gJZoh5iNFweWaAfdG2QiNvP6tlDUWBymJ6NOoRGKqG8'
SHEET_RANGE = 'Sheet1!A3:W'
DELIMITER = ','
ENCLOSURE = '"'

SHEET_HEADINGS = ["Company", "Website", "Title", "First Name", "Last Name", "Position", "Email"]
CSV_HEADINGS = ["Title", "First Name", "Last Name", "Email", "Institution", "Position", "Website", "Tags"]
FIELD_ORDER = [
    [8, 9, 10, 12, 0, 11, 1],
    [13, 14, 15, 17, 0, 16, 1],
    [18, 19, 20, 22, 0, 21, 1]
    ] # index 3 (0-6) is "email" in each case...
CC = ['first', 'second', 'third']


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
    with open(path + "/" + filename, 'w') as csvfile:
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
                # ok - quick and dirty
                for contact_fields in FIELD_ORDER:
                    dodgy = False # flag rows with dodgy data
                    contact = [] # create a new contact
                    pos = CC[FIELD_ORDER.index(contact_fields)]
                    #print('--- {} ---'.format(pos))
                    for field in contact_fields:
                        try:
                            value = row[field].encode("utf-8").strip()
                            #print('field: {} = {}'.format(field, value))
                            # email is the 7th field in the spreadsheet
                            # checking MX: validate_email(row[index].encode("utf-8").strip(), check_mx=True):
                            if contact_fields.index(field) == 3: # this is an email...
                                #print('email: {}'.format(value))
                                if not validate_email(value):
                                    dodgy = True
                                    #print('dodgy.')
                                #else:
                                    #print('valid.')
                            #row[index].append(" **invalid**")
                            # fill in the relevant contact field
                            contact.append(value)
                        except:
                            contact.append("Null")
                            dodgy = True

                    if not dodgy:
                        # title first last of institution (email, website)
                        contact.append(pos)
                        print('{} {} {} of {} ({}, {}) - {}'.format(contact[0], contact[1], contact[2], contact[4], contact[3], contact[6], contact[7]))
                        contacts.writerow(contact)
                    #else:
                        #print('**** Dodgy: {}'.format(contact))


if __name__ == '__main__':
    main()

# import the spread sheet and print it as a series of rows
# in the form:
# "title", "firstname", "lastname", "email", "organisation", "position", "website", "tags"
# where tags include "primary, cc1, cc2, cc3"

# Should disqualify rows without a valid email (do an email SMTP query to verify!)

#STARTROW = 1    # 0 indexed
