import gspread
from google.oauth2.service_account import Credentials

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file('credentials.json', scopes=scopes)
gc = gspread.authorize(creds)
ws = gc.open('บัญชี HDG 69').worksheet('มี.ค.69')

for r in [8, 9, 10, 194, 195, 196, 197, 203, 204, 205]:
    row = ws.row_values(r)
    vals = [v for v in row if v.strip()]
    status = str(vals[:4]) if vals else 'EMPTY'
    print(f'Row {r}: {status}')
