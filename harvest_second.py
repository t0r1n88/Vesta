"""
Скрипт для обработки данных  с многостраничных листов Excel
"""
import pandas as pd
import os
from dateutil.parser import ParserError
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl import load_workbook
import time
import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.options.mode.chained_assignment = None

skip_rows = 0
file_standard_merger = 'data/harvest/Приложение_№_1_Чеченская_Республика_01_12 (2).xlsx'
dir_name = 'data/harvest'
path_to_end_folder_merger = 'data/temp'
checkbox_harvest = 0

# Создаем датафрейм куда будем сохранять ошибочные файлы
err_df = pd.DataFrame(columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'])

name_file_standard_merger = file_standard_merger.split('/')[-1]  # получаем имя файла

standard_wb = load_workbook(filename=file_standard_merger)  # Загружаем эталонный файл

standard_sheets = sorted(standard_wb.sheetnames)  # отсортрованный список листов по которому будет вестись сравнение
standard_size_sheets = len(standard_sheets)

dct_df = dict()  # создаем словарь в котором будем хранить да

for sheet in standard_wb.sheetnames:  # Добавляем в словарь датафреймы
    temp_df = pd.read_excel(file_standard_merger, sheet_name=sheet, dtype=str)
    dct_df[sheet] = temp_df

for dirpath, dirnames, filenames in os.walk(dir_name):
    for filename in filenames:
        if filename.endswith('.xlsx') and filename != name_file_standard_merger:  # не обрабатываем эталонный файл
            # Получаем название файла без расширения
            name_file = filename.split('.xlsx')[0]
            print(name_file)
            temb_wb = load_workbook(filename=f'{dirpath}/{filename}')  # загружаем файл

            if standard_size_sheets == len(temb_wb.sheetnames):  # если количество листов одинаково то обрабатываем
                count_errors = 0
                for idx, data in enumerate(dct_df.items()):  # Проводим проверку на совпадение
                    print('*****')
                    print(idx)
                    name_sheet = data[0]  # получаем название листа
                    df = data[1]  # получаем датафрейм
                    print(name_sheet)
                    temp_name_sheet = temb_wb.sheetnames[idx] #
                    print(temb_wb[temp_name_sheet])

                    print('*****')
                    if len(temb_wb[temp_name_sheet][1]) != df.shape[1]:
                        # если количество колонок не совпадает то записываем как ошибку
                        temp_error_df = pd.DataFrame(
                            columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'], data=[
                                [name_file, name_sheet, 'Количество колонок отличается от эталонного',
                                 f'Ожидалось {df.shape[1]} колонок, а в листе {len(temb_wb[temp_name_sheet][1])}']])  # создаем временный датафрейм. потом надо подумать над словарем

                        err_df = pd.concat([err_df, temp_error_df], ignore_index=True)  # добавляем в датафрейм ошибок
                        count_errors += 1
                        # если хоть одна ошибка то проверяем следующий файл
                    else:
                        continue


                    if count_errors != 0:
                        continue
                    # если нет то начинаем обрабатывать листы
                for name_sheet, df in dct_df.items():
                    temp_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=temp_name_sheet,
                                            dtype=str)  # загружаем датафрейм
                    for row in dataframe_to_rows(temp_df, index=False, header=False):
                        standard_wb[name_sheet].append(row)  # добавляем данные
            else:
                continue  # если не совпадает то проверяем следующий файл

# Получаем текущую дату
current_time = time.strftime('%H_%M_%S %d.%m.%Y')
standard_wb.save(f'{path_to_end_folder_merger}/Общая таблица от {current_time}.xlsx')  # сохраняем





