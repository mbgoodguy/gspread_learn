import gspread
import os

from dotenv import load_dotenv
from gspread import Client, Spreadsheet, worksheet

load_dotenv()

table_id = os.getenv("TABLE_ID")

table_url = 'https://docs.google.com/spreadsheets/d/' + table_id


def show_available_worksheets(sh: Spreadsheet):
    worksheets = sh.worksheets()

    for ws in worksheets:
        print(ws)
        print(ws.id)
        print(ws.title)


def show_main_worksheet(sh: Spreadsheet):
    main_ws = sh.sheet1
    print('Певрый лист: ', main_ws.title)
    return main_ws


def create_ws_fill_and_delete(sh: Spreadsheet):
    another_ws = sh.add_worksheet('another', rows=3, cols=3)
    print(another_ws)
    input('enter to fill ws')
    another_ws.insert_row(["hello world", "hi"])
    input('enter to fill ws again')
    another_ws.insert_row(list(range(1, 10)))
    input('enter to delete ws')
    sh.del_worksheet(another_ws)


def update_ws_color(sh: Spreadsheet):
    ws = show_main_worksheet(sh)
    ws.update_tab_color('#d400ff')


def main():
    gc: Client = gspread.service_account('./gspreadlearn-06870fa5f3a0.json')
    sh: Spreadsheet = gc.open_by_url(table_url)
    # print(sh)
    # print(type(table_id))
    # print(sh.sheet1.get('A1:A2'))

    show_available_worksheets(sh)
    show_main_worksheet(sh)
    update_ws_color(sh)
    create_ws_fill_and_delete(sh)


if __name__ == '__main__':
    main()
