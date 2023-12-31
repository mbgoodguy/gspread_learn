import string
from pprint import pprint

import requests
import gspread
import os

from dotenv import load_dotenv
from gspread import Client, Spreadsheet, worksheet, Worksheet, Cell
from gspread.utils import rowcol_to_a1

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


def update_ws_color(sh: Spreadsheet):
    ws = show_main_worksheet(sh)
    ws.update_tab_color('#d400ff')


def insert_some_data(ws: Worksheet):
    ws.insert_rows([
        list(range(1, 10)),
        list(string.ascii_lowercase),
        list(string.ascii_uppercase),
        list(string.punctuation),
        list("ABOBA"),
    ])


def create_ws_fill_and_delete(sh: Spreadsheet):
    another_ws = sh.add_worksheet('another', rows=3, cols=3)
    print(another_ws)
    input('enter to fill ws')
    another_ws.insert_row(["hello world", "hi"])
    input('enter to fill ws again')
    another_ws.insert_row(list(range(1, 10)))
    input('enter to delete ws')
    sh.del_worksheet(another_ws)


# def update_table_by_cells(ws: Worksheet):
#     ws.update_cell(row=1, col=1, value='ABOBA')
#     ws.update_cell(row=2, col=1, value='AMOGUS')


def update_table_by_cells(ws: Worksheet):
    rows = 3
    cols = 3
    row = 4
    col = 2
    range_start = rowcol_to_a1(row, col)
    range_end = rowcol_to_a1(row + rows - 1, col + cols - 1)
    cells_range = f'{range_start}:{range_end}'
    print('update range: ', cells_range)
    values = [['a'] * cols] * rows
    print("values", values)
    ws.update(cells_range, values)


def show_all_values_in_ws(ws: Worksheet):
    list_of_lists = ws.get_all_values()  # получаем все значения, даже те, в которых отсутствует значение. Будет '' в пустой ячейке
    pprint(list_of_lists)


def create_and_fill_comments_with_ws(sh: Spreadsheet):
    comments_data = requests.get('https://jsonplaceholder.typicode.com/comments').json()
    header_row = ["postId", "id", "name", "email", "body"]
    rows = [header_row]
    for comment in comments_data:  # type: dict
        rows.append([
            comment.get(key, '')
            for key in header_row
        ])
    comment_ws = sh.add_worksheet("comments", rows=1, cols=len(header_row))
    comment_ws.insert_rows(rows)


def show_ws(ws: Worksheet):
    list_of_dicts = ws.get_all_records()
    pprint(list_of_dicts)  # для вывода большого кол-ва данных


def find_comment_by_author(ws: Worksheet):
    cell: Cell = ws.find('Carmen_Keeling@caroline.name')
    print('Found smthng at row %s and col %s' % (cell.row, cell.col))

    row = ws.row_values(cell.row)
    print(row)


# ф-и для массового обновления, напрмиер каждую 3 или 2 строку
def do_batch_update(ws: Worksheet):
    batches = []  # данные для обновления
    for i in range(1, 20, 2):
        items_count = i + 1
        addr_from = rowcol_to_a1(i, 1)  # откуда заполняем. Начинаем с первой
        addr_to = rowcol_to_a1(i, items_count)  # до куда заполняем
        data_range = f'{addr_from}:{addr_to}'
        print('add range', data_range)

        # какой range хотим обновить и какие values у нас будут
        batch = {
            'range': data_range,
            'values': [[i] * items_count],  # для каждого элемента в цикле for указываем все более длинный values
        }
        batches.append(batch)
    ws.batch_update(batches)  # метод для массового обновления. Принимает список со словарями


def main():
    gc: Client = gspread.service_account('./gspreadlearn-06870fa5f3a0.json')
    sh: Spreadsheet = gc.open_by_url(table_url)
    ws = sh.sheet1
    # ws2 = sh.worksheet('comments')

    # show_available_worksheets(sh)
    # show_main_worksheet(sh)
    # update_ws_color(sh)
    # create_ws_fill_and_delete(sh)
    # insert_some_data(ws)
    # update_table_by_cells(ws)
    # show_all_values_in_ws(ws)
    # create_and_fill_comments_with_ws(sh)
    # show_ws(ws2)  # вывод спсика со словарями в более удобном чем в print формате
    # comments_ws = sh.worksheet('comments')

    # find_comment_by_author(comments_ws)
    do_batch_update(ws)


if __name__ == '__main__':
    main()
