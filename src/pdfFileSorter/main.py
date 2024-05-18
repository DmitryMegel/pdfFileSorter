import os.path
import shutil
import webbrowser
from datetime import datetime
from threading import Thread
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Progressbar

import openpyxl
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
        frame_1.pack(fill=X, expand=True)
        self.file_path = Entry(frame_1, state='disabled', width=100)
        self.file_path.pack(side=LEFT, fill=X, expand=True)
        self.file_button = Button(frame_1, text='Выбрать', command=self.choose_excel_file)
        self.file_button.pack(side=LEFT)

        frame_2 = LabelFrame(text="Папка с PDF файлами:")
        frame_2.pack(fill=X, expand=True)
        self.pdf_folder = Entry(frame_2, state='disabled')
        self.pdf_folder.pack(side=LEFT, fill=X, expand=True)
        self.pdf_folder_button = Button(frame_2, text='Выбрать', command=lambda: self.choose_folder(self.pdf_folder))
        self.pdf_folder_button.pack(side=LEFT)

        frame_3 = LabelFrame(text="Сортировать в папку:")
        frame_3.pack(fill=X, expand=True)
        self.result_folder = Entry(frame_3, state='disabled')
        self.result_folder_button = Button(frame_3, text='Выбрать', command=lambda: self.choose_folder(self.result_folder))
        self.result_folder.pack(side=LEFT, fill=X, expand=True)
        self.result_folder_button.pack(side=LEFT)

        frame_4 = Frame()
        frame_4.pack()
        self.sort_button = Button(frame_4, text='Сортировать', command=self.start_sort)
        self.to_folder_button = Button(frame_4, text='Перейти в папку')
        self.to_log = Button(frame_4, text='Открыть отчет')
        self.log_file = ''

        frame_5 = Frame()
        frame_5.pack(fill=X, expand=True)
        self.progressbar = Progressbar(frame_5, mode="indeterminate")
        self.info = Label(frame_5)
        self.info.pack(pady=(10, 10))

    def choose_excel_file(self):
        types = (('Excel файлы', '*.xls;*.xlsx;*.xlsm'),)
        path = filedialog.askopenfilename(filetypes=types)
        self.update_fields(path, self.file_path)

    def choose_folder(self, folder):
        path = filedialog.askdirectory()
        self.update_fields(path, folder)

    def update_fields(self, path, field):
        field['state'] = 'normal'
        field.delete(0, END)
        field.insert(0, path)
        field['state'] = 'disabled'

        self.to_folder_button.pack_forget()
        self.to_log.pack_forget()
        self.info['text'] = ''

        if self.file_path.get() and self.pdf_folder.get() and self.result_folder.get():
            self.sort_button.pack(side=LEFT, padx=(10, 10), pady=(10, 10))
        else:
            self.sort_button.pack_forget()

    def get_sheet_names(self) -> list:
        wb = openpyxl.load_workbook(self.file_path.get())

        sh_names = list()
        for sheet in wb.worksheets:
            if sheet.sheet_properties.tabColor and sheet.sheet_properties.tabColor.rgb not in ('FF92D050', 'FFFFC000'):
                sh_names.append(sheet.title)

        return sh_names

    def get_pdf_names(self) -> dict:
        pdfs = {}
        cols = [1]

        for sheet_name in self.get_sheet_names():
            dataframe = pd.read_excel(self.file_path.get(), sheet_name=sheet_name, usecols=cols, skiprows=1)
            if dataframe.columns.values == 'Обозначение':
                datas = dataframe.iloc[:, 0].tolist()
                datas = [i for i in datas if i != ' ']
                pdfs.__setitem__(sheet_name, datas)

        pdfs = dict(filter(lambda x: x[1], pdfs.items()))
        return pdfs

    def create_folder(self, name):
        path = os.path.join(self.result_folder.get(), name)
        if not os.path.exists(path):
            os.mkdir(path)

    def open_log(self):
        path = os.path.abspath(self.log_file)
        webbrowser.open(path)

    def add_log_file(self, infos):
        dt_string = datetime.now().strftime("%d%m%Y_%H%M%S")
        infos = sorted(infos)

        if not os.path.exists('logs'):
            os.mkdir('logs')

        self.log_file = f'logs/log_{dt_string}.txt'

        log_file1 = open(self.log_file, 'w')
        log_file1.write(f'Операция выполнена частично. Не найдены файлы ({len(infos)} шт.):\n')
        log_file1.write('\n'.join(infos))
        log_file1.close()

    def copy_pdf_file(self, sheet_name, name):
        path_from = os.path.join(self.pdf_folder.get(), name)
        path_to = os.path.join(self.result_folder.get(), sheet_name, name)

        if not os.path.exists(path_to):
            shutil.copy(path_from, path_to)

        return name

    def sort_all_files(self, pdfs):
        not_found_files = set()

        for key, val in pdfs.items():
            self.create_folder(key)

            for value in val:
                if pd.isna(value):
                    continue

                name = f'{str(value).strip()}.pdf'.replace('\n', '')
                try:
                    self.copy_pdf_file(key, name)
                except FileNotFoundError:
                    not_found_files.add(name)

        if not_found_files:
            self.add_log_file(not_found_files)
            self.open_log()

        self.info['text'] = 'Сортировка завершена'

    def pack_result_buttons(self):
        self.info['text'] = 'Сортировка завершена'
        self.to_folder_button.pack(side=LEFT, padx=(10, 10))
        self.to_folder_button['command'] = lambda: webbrowser.open(self.result_folder.get())

        if self.log_file:
            self.to_log.pack(side=LEFT, padx=(10, 10))
            self.to_log.config(command=self.open_log)

    def start_sort(self):
        self.progressbar.pack(fill=X, expand=True)
        self.to_folder_button.pack_forget()
        self.to_log.pack_forget()
        self.info['text'] = ''
        self.progressbar.start()

        Thread(target=self.run).start()

    def run(self):
        try:
            pdf_names = self.get_pdf_names()
            self.sort_all_files(pdf_names)
            self.pack_result_buttons()
        except IndexError:
            self.info['text'] = 'Книга excel не подходит или имеет ошибки'
        except:
            self.info['text'] = 'Произошла непредвиденная ошибка'

        self.progressbar.stop()
        self.progressbar.pack_forget()


def main():
    sorter = SorterGUI()
    sorter.title("Сортировка PDF файлов на основе книги Excel")
    sorter.mainloop()


if __name__ == "__main__":
    main()
