import os.path
import shutil
from tkinter import filedialog as fd

import openpyxl
import pandas as pd
import numpy as np

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

    pdf_names = get_pdf_names(file_name, sheet_names)

    all_pdf_dir = fd.askdirectory()
    copy(all_pdf_dir, directory, pdf_names)


def open_file() -> str:
    types = (('Excel файлы', '*.xls;*.xlsx;*.xlsm'),)
    return fd.askopenfilename(filetypes=types)


def get_sheet_names(file_name: str) -> list:
    file = pd.ExcelFile(file_name)
    return file.sheet_names


def create_folders(directory: str, names: list):
    for name in names:
        path = os.path.join(directory, name)
        if not os.path.exists(path):
            os.mkdir(path)


def get_pdf_names(file_name, list_names):
    pdf_names = {}
    cols = [1]

    for sheet_name in list_names:
        if sheet_name == 'Общая':
            continue

        dataframe = pd.read_excel(file_name, sheet_name=sheet_name, usecols=cols, skiprows=1)
        datas = dataframe.iloc[:, 0].tolist()
        datas = [i for i in datas if i != ' ']
        pdf_names.__setitem__(sheet_name, datas)

    return pdf_names


def copy(all_pdf_dir: str, save_dir: str, pdf_names: dict):

    for key, val in pdf_names.items():
        for value in val:
            name = f'{value}.pdf'
            path_old = os.path.join(all_pdf_dir, name)
            path_new = os.path.join(save_dir, key, name)

            if not os.path.exists(path_new):
                shutil.copy(path_old, path_new)


if __name__ == "__main__":
    run()
