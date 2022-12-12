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
# file_standard_merger = 'data/temp2/Список 24.05.01 Проектирование, производство и эксплуатация ракет и ракетно-космических комплексов.xlsx'
dir_name = 'data/harvest'
# dir_name = 'data/temp2'
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


if checkbox_harvest == 0: # Вариант объеденения по названию листов
    for dirpath, dirnames, filenames in os.walk(dir_name):
        for filename in filenames:
            if filename.endswith('.xlsx') and filename != name_file_standard_merger:  # не обрабатываем эталонный файл
                # Получаем название файла без расширения
                name_file = filename.split('.xlsx')[0]
                print(name_file)
                temb_wb = load_workbook(filename=f'{dirpath}/{filename}')  # загружаем файл
                if len(temb_wb.sheetnames) == standard_size_sheets:  # сравниваем количество листов в файле
                    diff_name_sheets = set(temb_wb.sheetnames).difference(
                        set(standard_sheets))  # проверяем разницу в названиях листов
                    print(diff_name_sheets)
                    if len(diff_name_sheets) != 0:  # если разница в названиях есть то записываем в ошибки и обрабатываем следующий файл
                        temp_error_df = pd.DataFrame(
                            columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'], data=[
                                [name_file, '', 'Названия листов отличаются от эталонных',
                                 f'Отличаются следующие названия листов {diff_name_sheets}']])  # создаем временный датафрейм. потом надо подумать над словарем

                        err_df = pd.concat([err_df, temp_error_df], ignore_index=True)  # добавляем в датафрейм ошибок

                        continue

                    if standard_sheets == sorted(temb_wb.sheetnames):  # если названия листов одинаковые то обрабатываем
                        count_errors = 0

                        for name_sheet, df in dct_df.items():  # Проводим проверку на совпадение
                            print('*****')
                            print(name_sheet)
                            print(len(temb_wb[name_sheet][1]))
                            print(df.shape)
                            print('*****')
                            if len(temb_wb[name_sheet][1]) != df.shape[1]:
                                # если количество колонок не совпадает то записываем как ошибку
                                temp_error_df = pd.DataFrame(
                                    columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'],
                                    data=[[name_file, name_sheet, 'Количество колонок отличается от эталонного',
                                           f'Ожидалось {df.shape[1]} колонок, а в листе {len(temb_wb[name_sheet][1])}']])  # создаем временный датафрейм. потом надо подумать над словарем

                                err_df = pd.concat([err_df, temp_error_df],
                                                   ignore_index=True)  # добавляем в датафрейм ошибок
                                count_errors += 1

                        # если хоть одна ошибка то проверяем следующий файл
                        if count_errors != 0:
                            continue
                        # если нет то начинаем обрабатывать листы
                        for name_sheet, df in dct_df.items():
                            temp_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=name_sheet,
                                                    dtype=str)  # загружаем датафрейм
                            temp_df['Откуда взяты данные'] = name_file
                            for row in dataframe_to_rows(temp_df, index=False, header=False):
                                standard_wb[name_sheet].append(row)  # добавляем данные
                else:
                    continue  # если не совпадает то проверяем следующий файл

    # Получаем текущую дату
    current_time = time.strftime('%H_%M_%S %d.%m.%Y')
    standard_wb.save(f'{path_to_end_folder_merger}/Общая таблица от {current_time}.xlsx')  # сохраняем
    err_df.to_excel(f'{path_to_end_folder_merger}/Файлы с неправильными листами от {current_time}.xlsx',
                    index=False)  # сохраняем ошибки
elif checkbox_harvest == 1: # Вариант объединения по порядку
    for dirpath, dirnames, filenames in os.walk(dir_name):
        for filename in filenames:
            if filename.endswith('.xlsx') and filename != name_file_standard_merger:  # не обрабатываем эталонный файл
                # Получаем название файла без расширения
                name_file = filename.split('.xlsx')[0]
                print(name_file)
                temb_wb = load_workbook(filename=f'{dirpath}/{filename}')  # загружаем файл

                if standard_size_sheets == len(temb_wb.sheetnames):  # если количество листов одинаково то обрабатываем
                    count_errors = 0 # счетчик ошибок
                    dct_name_sheet = {} # создаем словарь где ключ это название листа в эталонном файле а значение это название листа в обрабатываемом файле
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
                            dct_name_sheet[name_sheet] = temp_name_sheet
                            continue

                    if count_errors != 0:
                        continue
                        # если нет то начинаем обрабатывать листы
                    for name_sheet, df in dct_df.items():
                        temp_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=dct_name_sheet[name_sheet],
                                                dtype=str)  # загружаем датафрейм
                        temp_df['Откуда взяты данные'] = name_file
                        for row in dataframe_to_rows(temp_df, index=False, header=False):
                            standard_wb[name_sheet].append(row)  # добавляем данные
                else:
                    continue  # если не совпадает то проверяем следующий файл

    # Получаем текущую дату
    current_time = time.strftime('%H_%M_%S %d.%m.%Y')
    standard_wb.save(f'{path_to_end_folder_merger}/Общая таблица от {current_time}.xlsx')  # сохраняем

    err_df.to_excel(f'{path_to_end_folder_merger}/Файлы с неправильными листами от {current_time}.xlsx',index=False)



