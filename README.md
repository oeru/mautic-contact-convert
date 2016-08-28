# mautic-contact-convert
A script to transform an ad hoc spreadsheet configuration into a CSV suitable for import as Mautic Contacts

This Python script reads data directly from a Google Spreadsheet, reorganises it, and formats it as double-quote enclosed comma separated values suitable for import into a Mautic contact database.

==Getting Google Spreadsheet access credentials:==
You must have a Google Dev account.

Used this howto: https://developers.google.com/sheets/quickstart/python

==Testing this script==

You have to define the sheet you're wanting to access, the credentials to use (you must've shared the sheet with the email in the credentials) and then resulting

'''python contactconvert.py'''
