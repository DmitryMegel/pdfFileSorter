import os.path
from tkinter import filedialog as fd

import pandas as pd

"""
Сортировщик PDF файлов

1. Получить путь к расположению Excel файла
2. Получить список листов Excel файла
3. Создать папки на ПК с названиями, соответствующими листам
4. Получить список PDF файлов каждого листа
5. Заполнить папки PDF файлами в том же составе, как указано в Excel файле
"""


def run():
    file_name = open_file()
    sheet_names = get_sheet_names(file_name)
    directory = fd.askdirectory()
    create_folders(directory, sheet_names)


def open_file() -> str:
    types = (('Excel файлы', '*.xls;*.xlsx;*.xlsm'),)
    return fd.askopenfilename(filetypes=types)


def get_sheet_names(file_name: str) -> list:
    file = pd.ExcelFile(file_name)
    return file.sheet_names


def create_folders(directory, names):
    for name in names:
        path = os.path.join(directory, name)
        os.mkdir(path)


if __name__ == "__main__":
    run()
