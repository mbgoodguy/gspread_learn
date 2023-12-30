import gspread
import os

from dotenv import load_dotenv
from gspread import Client, Spreadsheet

load_dotenv()

table_id = os.getenv("TABLE_ID")

table_url = 'https://docs.google.com/spreadsheets/d/' + table_id


def main():
    gc: Client = gspread.service_account('./gspreadlearn-06870fa5f3a0.json')
    sh: Spreadsheet = gc.open_by_url(table_url)
    print(sh)
    print(type(table_id))
    print(sh.sheet1.get('A1:A2'))


if __name__ == '__main__':
    main()
