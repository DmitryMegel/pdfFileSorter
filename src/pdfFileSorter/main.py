import os.path
import shutil
from tkinter import filedialog as fd, Tk, Button, Label, Entry, END

import pandas as pd

"""
Сортировщик PDF файлов

1. Получить путь к расположению Excel файла
2. Получить список листов Excel файла
3. Создать папки на ПК с названиями, соответствующими листам
4. Получить список PDF файлов каждого листа
5. Заполнить папки PDF файлами в том же составе, как указано в Excel файле
"""


class SorterGUI(Tk):

    def __init__(self):
        super().__init__()

        self.excel_path_lb = Label(self, text="Книга Excel:")
        self.excel_path_lb.grid(row=0, column=0)

        self.unsorted_dir_lb = Label(self, text="Папка с PDF файлами:")
        self.unsorted_dir_lb.grid(row=2, column=0)

        self.sorted_dir_lb = Label(self, text="Папка для сортировки:")
        self.sorted_dir_lb.grid(row=4, column=0)

        self.info = Label(self)
        self.info.grid(row=7, column=0)

        self.excel_path_f = Entry(self, width=80)
        self.excel_path_f.grid(row=1, column=0, padx=5, pady=5)

        self.unsorted_dir_f = Entry(self, width=80)
        self.unsorted_dir_f.grid(row=3, column=0, padx=5, pady=5)

        self.sorted_dir_f = Entry(self, width=80)
        self.sorted_dir_f.grid(row=5, column=0, padx=5, pady=5)

        self.excel_path_b = Button(self, text='Выбрать', command=self.select_excel_file)
        self.excel_path_b.grid(row=1, column=1, padx=5, pady=5)

        self.unsorted_dir_b = Button(self, text='Выбрать', command=self.select_unsorted_dir)
        self.unsorted_dir_b.grid(row=3, column=1, padx=5, pady=5)

        self.sorted_dir_b = Button(self, text='Выбрать')
        self.sorted_dir_b.grid(row=5, column=1, padx=5, pady=5)

        self.sort_b = Button(self, text='Сортировать')
        self.sort_b.grid(row=6, column=0, pady=5)

    def select_excel_file(self):
        types = (('Excel файлы', '*.xls;*.xlsx;*.xlsm'),)
        path = fd.askopenfilename(filetypes=types)

        self.excel_path_f.delete(0, END)
        self.excel_path_f.insert(0, path)

    def select_unsorted_dir(self):
        path = fd.askdirectory()

        self.unsorted_dir_f.delete(0, END)
        self.unsorted_dir_f.insert(0, path)


def run():
    sorter = SorterGUI()
    sorter.title("Сортировка PDF файлов на основе книги Excel")
    sorter.mainloop()

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


def get_pdf_names(file_name: str, list_names: list) -> dict:
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


def copy(all_pdf_dir: str, save_dir: str, pdf_names: dict) -> None:

    for key, val in pdf_names.items():
        for value in val:
            name = f'{value}.pdf'
            path_old = os.path.join(all_pdf_dir, name)
            path_new = os.path.join(save_dir, key, name)

            if not os.path.exists(path_new):
                shutil.copy(path_old, path_new)


if __name__ == "__main__":
    run()
