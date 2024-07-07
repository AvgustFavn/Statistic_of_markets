# Importing required library
import os

import pygsheets

# Create the Client
# Enter the name of the downloaded KEYS
# file in service_account_file
client = pygsheets.authorize(
    service_account_file=f"{os.getcwd()}/test-mps-rentable-02e4e2bdc3d4.json".replace('\\', '/'))

# Sample command to verify successful
# authorization of pygsheets
# Prints the names of the spreadsheet
# shared with or owned by the service
# account
print(client.spreadsheet_titles())
