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

        self.excel_path_f = Entry(self, width=80, state='disabled')
        self.excel_path_f.grid(row=1, column=0, padx=5, pady=5)

        self.unsorted_dir_f = Entry(self, width=80, state='disabled')
        self.unsorted_dir_f.grid(row=3, column=0, padx=5, pady=5)

        self.sorted_dir_f = Entry(self, width=80, state='disabled')
        self.sorted_dir_f.grid(row=5, column=0, padx=5, pady=5)

        self.excel_path_b = Button(self, text='Выбрать', command=self.select_excel_file)
        self.excel_path_b.grid(row=1, column=1, padx=5, pady=5)

        self.unsorted_dir_b = Button(self, text='Выбрать', command=self.select_unsorted_dir)
        self.unsorted_dir_b.grid(row=3, column=1, padx=5, pady=5)

        self.sorted_dir_b = Button(self, text='Выбрать', command=self.select_sorted_dir)
        self.sorted_dir_b.grid(row=5, column=1, padx=5, pady=5)

        self.sort_b = Button(self, text='Сортировать', command=self.run)
        self.sort_b.grid(row=6, column=0, pady=5)

    def select_excel_file(self):
        types = (('Excel файлы', '*.xls;*.xlsx;*.xlsm'),)
        path = fd.askopenfilename(filetypes=types)
        self.update_field(path, self.excel_path_f)

    def select_unsorted_dir(self):
        path = fd.askdirectory()
        self.update_field(path, self.unsorted_dir_f)

    def select_sorted_dir(self):
        path = fd.askdirectory()
        self.update_field(path, self.sorted_dir_f)

    def update_field(self, path, field):
        field.config(state='normal')
        field.delete(0, END)
        field.insert(0, path)
        field.config(state='disabled')
        self.info.config(text='')

    def get_sheet_names(self) -> list:
        file = pd.ExcelFile(self.excel_path_f.get())
        return file.sheet_names

    def create_folders(self, names: list):
        for name in names:
            path = os.path.join(self.sorted_dir_f.get(), name)
            if not os.path.exists(path):
                os.mkdir(path)

    def get_pdf_names(self, list_names: list) -> dict:
        pdf_names = {}
        cols = [1]

        for sheet_name in list_names:
            if sheet_name == 'Общая':
                continue

            dataframe = pd.read_excel(self.excel_path_f.get(), sheet_name=sheet_name, usecols=cols, skiprows=1)
            datas = dataframe.iloc[:, 0].tolist()
            datas = [i for i in datas if i != ' ']
            pdf_names.__setitem__(sheet_name, datas)

        return pdf_names

    def save_with_sort(self, pdf_names: dict) -> None:
        not_found_files = list()
        for key, val in pdf_names.items():
            for value in val:
                name = f'{value}.pdf'
                try:
                    path_old = os.path.join(self.unsorted_dir_f.get(), name)
                    path_new = os.path.join(self.sorted_dir_f.get(), key, name)

                    if not os.path.exists(path_new):
                        shutil.copy(path_old, path_new)
                except FileNotFoundError:
                    not_found_files.append(name)

        if not_found_files:
            self.info.config(text=f'Операция выполнена частично. \nНе найдены файлы: {not_found_files}')
        else:
            self.info.config(text='Операция успешно выполнена')

    def run(self):
        try:
            if self.excel_path_f.get() and self.unsorted_dir_f.get() and self.sorted_dir_f.get():
                sheet_names = self.get_sheet_names()
                pdf_names = self.get_pdf_names(sheet_names)
                self.create_folders(sheet_names)
                self.save_with_sort(pdf_names)
            else:
                self.info.config(text='Заполнены не все поля')
        except IndexError:
            self.info.config(text='Книга excel не подходит или имеет ошибки')


def main():
    sorter = SorterGUI()
    sorter.title("Сортировка PDF файлов на основе книги Excel")
    sorter.mainloop()


if __name__ == "__main__":
    main()
