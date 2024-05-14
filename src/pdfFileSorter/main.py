import os.path
import shutil
import webbrowser
from tkinter import *
from tkinter import filedialog

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

        frame_1 = LabelFrame(text='Файл Excel:')
        frame_2 = LabelFrame(text="Папка с PDF файлами:")
        frame_3 = LabelFrame(text="Папка для сортировки:")
        frame_4 = Frame()
        frame_5 = Frame()

        self.excel_path_f = Entry(frame_1, state='disabled')
        self.unsorted_dir_f = Entry(frame_2, state='disabled')
        self.sorted_dir_f = Entry(frame_3, state='disabled')

        self.excel_path_b = Button(frame_1, text='Выбрать', command=self.select_excel_file)
        self.unsorted_dir_b = Button(frame_2, text='Выбрать', command=self.select_unsorted_dir)
        self.sorted_dir_b = Button(frame_3, text='Выбрать', command=self.select_sorted_dir)
        self.sort_b = Button(frame_4, text='Сортировать', command=self.run)
        self.result_b = Button(frame_4, text='Перейти в папку', state='disabled')

        self.info_log = Text(frame_5, wrap=WORD)

        frame_1.pack(fill=X, expand=True)
        self.excel_path_f.pack(side=LEFT, fill=X, expand=True)
        self.excel_path_b.pack(side=LEFT)

        frame_2.pack(fill=X, expand=True)
        self.unsorted_dir_f.pack(side=LEFT, fill=X, expand=True)
        self.unsorted_dir_b.pack(side=LEFT)

        frame_3.pack(fill=X, expand=True)
        self.sorted_dir_f.pack(side=LEFT, fill=X, expand=True)
        self.sorted_dir_b.pack(side=LEFT)

        frame_4.pack()
        self.sort_b.pack(side=LEFT, padx=(10, 10), pady=(10, 10))
        self.result_b.pack(side=LEFT)

        frame_5.pack()
        self.info_log.pack(side=LEFT)

    def select_excel_file(self):
        types = (('Excel файлы', '*.xls;*.xlsx;*.xlsm'),)
        path = filedialog.askopenfilename(filetypes=types)
        self.update_field(path, self.excel_path_f)
        self.info_log.insert(END, f'Выбран excel файл: {path}\n')

    def select_unsorted_dir(self):
        path = filedialog.askdirectory()
        self.update_field(path, self.unsorted_dir_f)
        self.info_log.insert(END, f'Выбрана папка с PDF файлами: {path}\n')

    def select_sorted_dir(self):
        path = filedialog.askdirectory()
        self.update_field(path, self.sorted_dir_f)
        self.info_log.insert(END, f'Выбрана папка, куда будет производиться сортировка: {path}\n')

    def update_field(self, path, field):
        field.config(state='normal')
        field.delete(0, END)
        field.insert(0, path)
        field.config(state='disabled')

    def open_result_dir(self):
        webbrowser.open(self.sorted_dir_f.get())

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
            self.info_log.insert(END, 'WARNING: Операция выполнена частично\n')
            self.info_log.insert(END, f'WARNING: Не найдены файлы: {not_found_files}\n')

    def run(self):
        try:
            if self.excel_path_f.get() and self.unsorted_dir_f.get() and self.sorted_dir_f.get():
                sheet_names = self.get_sheet_names()
                self.info_log.insert(END, f'\nINFO: Список листов файла excel: {sheet_names}\n')
                pdf_names = self.get_pdf_names(sheet_names)
                self.info_log.insert(END, 'INFO: Начало сортировки\n')
                self.create_folders(sheet_names)
                self.info_log.insert(END, f'INFO: Сгенерированы папки\n')
                self.save_with_sort(pdf_names)
                self.info_log.insert(END, f'INFO: PDF файлы распределены по папкам\n')
                self.info_log.insert(END, 'INFO: Сортировка завершена\n')
                self.result_b.config(state='normal', command=self.open_result_dir)
            else:
                self.info_log.insert(END, 'WARNING: Заполнены не все поля\n')
        except IndexError:
            self.info_log.insert(END, 'ERROR: Книга excel не подходит или имеет ошибки\n')


def main():
    sorter = SorterGUI()
    sorter.title("Сортировка PDF файлов на основе книги Excel")
    sorter.mainloop()


if __name__ == "__main__":
    main()
