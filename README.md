# mautic-contact-convert
A script to transform an ad hoc spreadsheet configuration into a CSV suitable for import as Mautic Contacts

This Python script reads data directly from a Google Spreadsheet, reorganises it, and formats it as double-quote enclosed comma separated values suitable for import into a Mautic contact database.

==Getting Google Spreadsheet access credentials:==
You must have a Google Dev account.

Used this howto: https://developers.google.com/sheets/quickstart/python

Deprecated...
Go to https://console.developers.google.com/ which (as of 26 Aug 2016) should show you the API Manager. Select the "Credentials" option (left column) and "Create credentials", selecting "Service Account keys".

Save the resulting file as google-servicekeys.json

The email provided in that file will be crucial - you will need to share the spreadsheet you're referencing with that email address.


==Testing this script==

You have to define the sheet you're wanting to access, the credentials to use (you must've shared the sheet with the email in the credentials) and then resulting

'''python src/contactconvert.py'''
