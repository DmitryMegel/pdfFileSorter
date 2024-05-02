import os.path
import tkinter as tk
from tkinter import filedialog as fd
import pandas as pd
from openpyxl import load_workbook

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


def open_file() -> str:
    types = (('Excel файлы', '*.xls;*.xlsx;*.xlsm'),)
    return fd.askopenfilename(filetypes=types)


if __name__ == "__main__":
    run()
