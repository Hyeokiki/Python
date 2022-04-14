from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("hah-project-8b41c84e8c4a.json", scope)

import gspread

spreadsheet_name = "부자재 입출고 재고 시트"
client = gspread.authorize(creds)
spreadsheet = client.open(spreadsheet_name)

## by name
sheet = spreadsheet.worksheet("출고")

print(sheet)

print(sheet.get_all_values())