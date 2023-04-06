import pandas as pd
import numpy as np
import os
from dateutil.parser import ParserError
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl import load_workbook
import pytrovich
from pytrovich.detector import PetrovichGenderDetector
from pytrovich.enums import NamePart, Gender, Case
from pytrovich.maker import PetrovichDeclinationMaker
import time
import datetime
import warnings
from collections import Counter
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.options.mode.chained_assignment = None
import sys
import locale
import logging
import tempfile
import re
logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)

# Классы для исключений

class CheckBoxException(Exception):
    """
    Класс для вызовы исключения в случае если неправильно выставлены чекбоксы
    """
    pass

class NotFoundValue(Exception):
    """
    Класс для обозначения того что значение не найдено
    """
    pass


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)






def select_file_template_doc():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_doc
    name_file_template_doc = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_doc():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_doc
    # Получаем путь к файлу
    name_file_data_doc = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_doc():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_doc
    path_to_end_folder_doc = filedialog.askdirectory()


def select_file_data_date():
    """
    Функция для выбора файла с данными для которого нужно разбить по категориям
    :return: Путь к файлу с данными
    """
    global name_file_data_date
    # Получаем путь к файлу
    name_file_data_date = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


# Функциия для слияния таблиц

def convert_columns_to_str(df, number_columns):
    """
    Функция для конвертации указанных столбцов в строковый тип и очистки от пробельных символов в начале и конце
    """

    for column in number_columns:  # Перебираем список нужных колонок
        try:
            df.iloc[:, column] = df.iloc[:, column].astype(str)
            # Очищаем колонку от пробельных символов с начала и конца
            df.iloc[:, column] = df.iloc[:, column].apply(lambda x: x.strip())
        except IndexError:
            messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                                 'Проверьте порядковые номера колонок которые вы хотите обработать.')


def convert_params_columns_to_int(lst):
    """
    Функция для конвератации значений колонок которые нужно обработать.
    Очищает от пустых строк, чтобы в итоге остался список из чисел в формате int
    """
    out_lst = [] # Создаем список в который будем добавлять только числа
    for value in lst: # Перебираем список
        try:
            # Обрабатываем случай с нулем, для того чтобы после приведения к питоновскому отсчету от нуля не получилась колонка с номером -1
            number = int(value)
            if number != 0:
                out_lst.append(value) # Если конвертирования прошло без ошибок то добавляем
            else:
                continue
        except: # Иначе пропускаем
            continue
    return out_lst


def select_file_params_comparsion():
    """
    Функция для выбора файла с параметрами колонок т.е. кокие колонки нужно обрабатывать
    :return:
    """
    global file_params
    file_params = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_first_comparison():
    """
    Функция для выбора  первого файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_first_file_comparison
    # Получаем путь к файлу
    name_first_file_comparison = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_comparison():
    """
    Функция для выбора  второго файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_second_file_comparison
    # Получаем путь к файлу
    name_second_file_comparison = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_comparison():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_comparison
    path_to_end_folder_comparison = filedialog.askdirectory()


def select_end_folder_date():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_date
    path_to_end_folder_date = filedialog.askdirectory()


def select_file_data_groupby():
    """
    Функция для выбора файла с данными
    :return:
    """
    global name_file_data_groupby
    # Получаем путь к файлу
    name_file_data_groupby = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_groupby():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_groupby
    path_to_end_folder_groupby = filedialog.askdirectory()


# Функции для вкладки извлечение данных
def select_file_params_calculate_data():
    """
    Функция для выбора файла c ячейками которые нужно подсчитать
    :return: Путь к файлу
    """
    global name_file_params_calculate_data
    name_file_params_calculate_data = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_files_data_calculate_data():
    """
    Функция для выбора файлов с данными параметры из которых нужно подсчитать
    :return: Путь к файлам с данными
    """
    global names_files_calculate_data
    # Получаем путь к файлу
    names_files_calculate_data = filedialog.askopenfilenames(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_calculate_data():
    """
    Функция для выбора папки куда будут генерироваться файл  с результатом подсчета и файл с проверочной инфомрацией
    :return:
    """
    global path_to_end_folder_calculate_data
    path_to_end_folder_calculate_data = filedialog.askdirectory()


def calculate_data():
    """
    Функция для подсчета данных из файлов
    :return:
    """
    try:
        count = 0
        count_errors = 0
        quantity_files = len(names_files_calculate_data)
        current_time = time.strftime('%H_%M_%S')
        # Состояние чекбокса
        mode_text = mode_text_value.get()

        # Получаем название обрабатываемого листа
        name_list_df = pd.read_excel(name_file_params_calculate_data, nrows=2)
        name_list = name_list_df['Значение'].loc[0]

        # Получаем количество листов в файле, на случай если название листа не совпадает с правильным
        quantity_list_in_file = name_list_df['Значение'].loc[1]

        # Получаем шаблон с данными, первую строку пропускаем, поскольку название обрабатываемого листа мы уже получили
        df = pd.read_excel(name_file_params_calculate_data, skiprows=2)

        # Создаем словарь параметров
        param_dict = dict()

        for row in df.itertuples():
            param_dict[row[1]] = row[2]
        # Создаем словарь для подсчета данных, копируя ключи из словаря параметров, значения в зависимости от способа обработки

        if mode_text == 'Yes':
            result_dct = {key: '' for key, value in param_dict.items()}
        else:
            result_dct = {key: 0 for key, value in param_dict.items()}

            # Создаем датафрейм для контроля процесса подсчета и заполняем словарь на основе которого будем делать итоговую таблицу

        check_df = pd.DataFrame(columns=param_dict.keys())
        # Вставляем колонку для названия файла
        check_df.insert(0, 'Название файла', '')
        for file in names_files_calculate_data:
            # Проверяем чтобы файл не был резервной копией.
            if '~$' in file:
                continue
            # Создаем словарь для создания строки которую мы будем добавлять в проверочный датафрейм
            new_row = dict()
            # Получаем  отбрасываем расширение файла
            full_name_file = file.split('.')[0]
            # Получаем имя файла  без пути
            name_file = full_name_file.split('/')[-1]
            try:

                new_row['Название файла'] = name_file

                wb = openpyxl.load_workbook(file)
                # Проверяем наличие листа
                if name_list in wb.sheetnames:
                    sheet = wb[name_list]
                # проверяем количество листов в файле.Если значение равно 1 то просто берем первый лист, иначе вызываем ошибку
                elif quantity_list_in_file == 1:
                    temp_name = wb.sheetnames[0]
                    sheet = wb[temp_name]
                else:
                    raise Exception
                for key, cell in param_dict.items():
                    result_dct[key] += check_data(sheet[cell].value, mode_text)
                    new_row[key] = sheet[cell].value

                temp_df = pd.DataFrame(new_row, index=['temp_index'])
                check_df = pd.concat([check_df, temp_df], ignore_index=True)
                # check_df = check_df.append(new_row, ignore_index=True)

                count += 1
            # Ловим исключения
            except Exception as err:
                count_errors += 1
                with open(f'{path_to_end_folder_calculate_data}/Необработанные файлы {current_time}.txt', 'a',
                          encoding='utf-8') as f:
                    f.write(f'Файл {name_file} не обработан!!!\n')

        check_df.to_excel(f'{path_to_end_folder_calculate_data}/Проверка вычисления {current_time}.xlsx', index=False)

        # Создание итоговой таблицы результатов подсчета

        finish_result = pd.DataFrame()

        finish_result['Наименование показателя'] = result_dct.keys()
        finish_result['Значение показателя'] = result_dct.values()
        # Проводим обработку в зависимости от значения переключателя

        # Получаем текущее время для того чтобы использовать в названии

        if mode_text == 'Yes':
            # Обрабатываем датафрейм считая текстовые данные
            count_text_df = count_text_value(finish_result)
            count_text_df.to_excel(
                f'{path_to_end_folder_calculate_data}/Подсчет текстовых значений {current_time}.xlsx')
        else:
            finish_result.to_excel(f'{path_to_end_folder_calculate_data}/Итоговые значения {current_time}.xlsx',
                                   index=False)

        if count_errors != 0:
            messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30',
                                f'Обработка файлов завершена!\nОбработано файлов:  {count} из {quantity_files}\n Необработанные файлы указаны в файле {path_to_end_folder_calculate_data}/ERRORS {current_time}.txt ')
        else:
            messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30',
                                f'Обработка файлов успешно завершена!\nОбработано файлов:  {count} из {quantity_files}')
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')


# Функции для слияния таблиц

def select_end_folder_merger():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_merger
    path_to_end_folder_merger = filedialog.askdirectory()


def select_folder_data_merger():
    """
    Функция для выбора папки где хранятся нужные файлы
    :return:
    """
    global dir_name
    dir_name = filedialog.askdirectory()

def select_params_file_merger():
    """
    Функция для выбора файла c ячейками которые нужно подсчитать
    :return: Путь к файлу
    """
    if group_rb_type_harvest.get() == 2:
        global params_harvest
        params_harvest = filedialog.askopenfilename(
            filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
    else:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30','Выберите вариант слияния В и попробуйте снова ')



def select_standard_file_merger():
    """
    Функция для выбора файла c ячейками которые нужно подсчитать
    :return: Путь к файлу
    """
    global file_standard_merger
    file_standard_merger = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def merge_tables():
    """
    Функция для слияния таблиц с одинаковой структурой в одну большую таблицу
    """
    # Получаем значения из полей ввода и проверяем их на тип
    try:
        checkbox_harvest = group_rb_type_harvest.get()
        if checkbox_harvest != 2:
            skip_rows = int(merger_entry_skip_rows.get())
    except ValueError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Введите целое число в поле для ввода количества пропускаемых строк!!!')
    else:
        # Оборачиваем в try
        try:
            # Создаем датафрейм куда будем сохранять ошибочные файлы
            err_df = pd.DataFrame(columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'])

            name_file_standard_merger = file_standard_merger.split('/')[-1]  # получаем имя файла

            standard_wb = load_workbook(filename=file_standard_merger)  # Загружаем эталонный файл

            standard_sheets = sorted(
                standard_wb.sheetnames)  # отсортрованный список листов по которому будет вестись сравнение
            set_standard_sheets = set(standard_sheets)  # создаем множество из листов эталонного файла
            standard_size_sheets = len(standard_sheets)

            "Удаляем пустые и строки с заливкой которые могут тянуться вниз и из этого данные из других файлов начина" \
            "ются с тысячных строк"
            for sheet in standard_wb.sheetnames:
                del_cols_df = pd.read_excel(file_standard_merger,
                                            sheet_name=sheet)  # загружаем датафрейм чтобы узнать сколько есть заполненны строк

                temp_sheet_max_row = standard_wb[sheet].max_row  # получаем последнюю строку
                standard_wb[sheet].delete_rows(del_cols_df.shape[0] + 2, temp_sheet_max_row)  # удаляем все лишнее

            dct_df = dict()  # создаем словарь в котором будем хранить да

            for sheet in standard_wb.sheetnames:  # Добавляем в словарь датафреймы
                temp_df = pd.read_excel(file_standard_merger, sheet_name=sheet, dtype=str)
                dct_df[sheet] = temp_df

            if checkbox_harvest == 0:  # Вариант объединения по названию листов
                for dirpath, dirnames, filenames in os.walk(dir_name):
                    for filename in filenames:
                        if (filename.endswith('.xlsx') and not filename.startswith(
                                '~$')) and filename != name_file_standard_merger:  # не обрабатываем эталонный файл
                            # Получаем название файла без расширения
                            name_file = filename.split('.xlsx')[0]
                            print(name_file)
                            temb_wb = load_workbook(
                                filename=f'{dirpath}/{filename}')  # загружаем файл, для проверки листов
                            """
                            Проверяем наличие листов из эталонного файла в проверяемом файле, если они есть то начинаем 
                            дальнейшую проверку
                            """
                            if set_standard_sheets.issubset(set(temb_wb.sheetnames)):
                                count_errors = 0
                                # проверяем наличие листов указанных в файле параметров
                                for name_sheet, df in dct_df.items():  # Проводим проверку на совпадение
                                    lst_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=name_sheet)
                                    if lst_df.shape[1] != df.shape[1]:
                                        # если количество колонок не совпадает то записываем как ошибку
                                        temp_error_df = pd.DataFrame(
                                            columns=['Название файла', 'Наименование листа', 'Тип ошибки',
                                                     'Описание ошибки'],
                                            data=[[name_file, name_sheet, 'Количество колонок отличается от эталонного',
                                                   f'Ожидалось {df.shape[1]} колонок, а в листе {lst_df.shape[1]}']])  # создаем временный датафрейм. потом надо подумать над словарем

                                        err_df = pd.concat([err_df, temp_error_df],
                                                           ignore_index=True)  # добавляем в датафрейм ошибок
                                        count_errors += 1

                                # если хоть одна ошибка то проверяем следующий файл
                                if count_errors != 0:
                                    continue
                                # если нет то начинаем обрабатывать листы
                                for name_sheet, df in dct_df.items():
                                    temp_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=name_sheet,
                                                            dtype=str, skiprows=skip_rows,header=None)  # загружаем датафрейм
                                    if temp_df.shape[1] > 3:
                                        temp_df = temp_df.dropna(axis=0,thresh=2)

                                    temp_df['Номер строки'] = range(1, temp_df.shape[0] + 1)
                                    temp_df['Откуда взяты данные'] = name_file
                                    for row in dataframe_to_rows(temp_df, index=False, header=False):
                                        standard_wb[name_sheet].append(row)  # добавляем данные

                            elif len(
                                    temb_wb.sheetnames) == standard_size_sheets:  # сравниваем количество листов в файле
                                diff_name_sheets = set(temb_wb.sheetnames).difference(
                                    set(standard_sheets))  # проверяем разницу в названиях листов
                                print(diff_name_sheets)
                                if len(diff_name_sheets) != 0:  # если разница в названиях есть то записываем в ошибки и обрабатываем следующий файл
                                    temp_error_df = pd.DataFrame(
                                        columns=['Название файла', 'Наименование листа', 'Тип ошибки',
                                                 'Описание ошибки'], data=[
                                            [name_file, '', 'Названия листов отличаются от эталонных',
                                             f'Отличаются следующие названия листов {diff_name_sheets}']])  # создаем временный датафрейм. потом надо подумать над словарем

                                    err_df = pd.concat([err_df, temp_error_df],
                                                       ignore_index=True)  # добавляем в датафрейм ошибок

                                    continue

                            else:
                                temp_error_df = pd.DataFrame(
                                    columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'],
                                    data=[[name_file, '', 'Не совпадает количество или название листов в файле',
                                           f'Листы, которые есть в файле: {",".join(temb_wb.sheetnames)}']])  # создаем временный датафрейм. потом надо подумать над словарем

                                err_df = pd.concat([err_df, temp_error_df],
                                                   ignore_index=True)  # добавляем в датафрейм ошибок

                # Получаем текущую дату
                current_time = time.strftime('%H_%M_%S %d.%m.%Y')
                standard_wb.save(
                    f'{path_to_end_folder_merger}/Слияние по варианту А Общая таблица от {current_time}.xlsx')  # сохраняем
                err_out_wb = openpyxl.Workbook()  # создаем объект openpyxl для сохранения датафрейма
                for row in dataframe_to_rows(err_df, index=False, header=True):
                    err_out_wb['Sheet'].append(row)  # добавляем данные
                # устанавливаем размер колонок
                err_out_wb['Sheet'].column_dimensions['A'].width = 40
                err_out_wb['Sheet'].column_dimensions['B'].width = 30
                err_out_wb['Sheet'].column_dimensions['C'].width = 55
                err_out_wb['Sheet'].column_dimensions['D'].width = 100
                err_out_wb.save(f'{path_to_end_folder_merger}/Слияние по варианту А Ошибки от {current_time}.xlsx')

            elif checkbox_harvest == 1:  # Вариант объединения по порядку
                for dirpath, dirnames, filenames in os.walk(dir_name):
                    for filename in filenames:
                        if (filename.endswith('.xlsx') and not filename.startswith(
                                '~$')) and filename != name_file_standard_merger:  # не обрабатываем эталонный файл
                            # Получаем название файла без расширения
                            name_file = filename.split('.xlsx')[0]
                            print(name_file)
                            temb_wb = load_workbook(filename=f'{dirpath}/{filename}')  # загружаем файл

                            if standard_size_sheets == len(
                                    temb_wb.sheetnames):  # если количество листов одинаково то обрабатываем
                                count_errors = 0  # счетчик ошибок
                                dct_name_sheet = {}  # создаем словарь где ключ это название листа в эталонном файле а значение это название листа в обрабатываемом файле
                                for idx, data in enumerate(dct_df.items()):  # Проводим проверку на совпадение
                                    name_sheet = data[0]  # получаем название листа
                                    df = data[1]  # получаем датафрейм
                                    temp_name_sheet = temb_wb.sheetnames[idx]  #
                                    lst_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=temp_name_sheet)
                                    if lst_df.shape[1] != df.shape[1]:
                                        # если количество колонок не совпадает то записываем как ошибку
                                        temp_error_df = pd.DataFrame(
                                            columns=['Название файла', 'Наименование листа', 'Тип ошибки',
                                                     'Описание ошибки'],
                                            data=[
                                                [name_file, name_sheet, 'Количество колонок отличается от эталонного',
                                                 f'Ожидалось {df.shape[1]} колонок, а в листе {lst_df.shape[1]}']])  # создаем временный датафрейм. потом надо подумать над словарем

                                        err_df = pd.concat([err_df, temp_error_df],
                                                           ignore_index=True)  # добавляем в датафрейм ошибок
                                        count_errors += 1

                                    else:
                                        dct_name_sheet[name_sheet] = temp_name_sheet
                                # если хоть одна ошибка то проверяем следующий файл
                                if count_errors != 0:
                                    continue
                                    # если нет то начинаем обрабатывать листы
                                for name_sheet, df in dct_df.items():
                                    temp_df = pd.read_excel(f'{dirpath}/{filename}',
                                                            sheet_name=dct_name_sheet[name_sheet],
                                                            dtype=str, skiprows=skip_rows,header=None)  # загружаем датафрейм
                                    if temp_df.shape[1] > 3:
                                        temp_df = temp_df.dropna(axis=0, thresh=2)
                                    temp_df['Номер строки'] = range(1, temp_df.shape[0] + 1)
                                    temp_df['Откуда взяты данные'] = name_file
                                    for row in dataframe_to_rows(temp_df, index=False, header=False):
                                        standard_wb[name_sheet].append(row)  # добавляем данные
                            else:
                                temp_error_df = pd.DataFrame(
                                    columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'],
                                    data=[[name_file, '', 'Не совпадает количество или название листов в файле',
                                           f'Листы, которые есть в файле: {",".join(temb_wb.sheetnames)}']])  # создаем временный датафрейм. потом надо подумать над словарем

                                err_df = pd.concat([err_df, temp_error_df],
                                                   ignore_index=True)  # добавляем в датафрейм ошибок

                # Получаем текущую дату
                current_time = time.strftime('%H_%M_%S %d.%m.%Y')
                standard_wb.save(
                    f'{path_to_end_folder_merger}/Слияние по варианту Б Общая таблица от {current_time}.xlsx')  # сохраняем

                err_out_wb = openpyxl.Workbook()  # создаем объект openpyxl для сохранения датафрейма
                for row in dataframe_to_rows(err_df, index=False, header=True):
                    err_out_wb['Sheet'].append(row)  # добавляем данные
                # устанавливаем размер колонок
                err_out_wb['Sheet'].column_dimensions['A'].width = 40
                err_out_wb['Sheet'].column_dimensions['B'].width = 30
                err_out_wb['Sheet'].column_dimensions['C'].width = 55
                err_out_wb['Sheet'].column_dimensions['D'].width = 100
                err_out_wb.save(f'{path_to_end_folder_merger}/Слияние по варианту Б Ошибки от {current_time}.xlsx')

            # Если выбран управляемый сбор данных
            elif checkbox_harvest == 2:
                df_params = pd.read_excel(params_harvest, header=None)  # загружаем параметры
                df_params[0] = df_params[0].astype(
                    str)  # делаем данные строковыми чтобы корректно работало обращение по названию листов

                tmp_name_sheets = df_params[0].tolist()  # создаем списки чтобы потом из них сделать словарь
                tmp_skip_rows = df_params[1].tolist()
                dct_manage_harvest = dict(zip(tmp_name_sheets,
                                              tmp_skip_rows))  # создаем словарь где ключ это название листа а значение это сколько строк нужно пропустить
                set_params_sheets = set(
                    dct_manage_harvest.keys())  # создаем множество из ключей(листов) которые нужно обработать
                if not set_params_sheets.issubset(
                        set_standard_sheets):  # проверяем совпадение названий в эталонном файле и в файле параметров
                    diff_value = set(dct_manage_harvest.keys()).difference(set(standard_sheets))  # получаем разницу

                    messagebox.showerror('',
                                         f'Не совпадают следующие названия листов в файле параметров и в эталонном файле\n'
                                         f'{diff_value}!')
                # начинаем обработку
                for dirpath, dirnames, filenames in os.walk(dir_name):
                    for filename in filenames:
                        if (filename.endswith('.xlsx') and not filename.startswith(
                                '~$')) and filename != name_file_standard_merger:  # не обрабатываем эталонный файл
                            # Получаем название файла без расширения
                            name_file = filename.split('.xlsx')[0]
                            print(name_file)
                            temb_wb = load_workbook(filename=f'{dirpath}/{filename}')  # загружаем файл
                            if set_params_sheets.issubset(set(temb_wb.sheetnames)):
                                count_errors = 0
                                # проверяем наличие листов указанных в файле параметров
                                for name_sheet, skip_r in dct_manage_harvest.items():  # Проводим проверку на совпадение количества колонок
                                    lst_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=name_sheet)
                                    if lst_df.shape[1] != dct_df[name_sheet].shape[1]:
                                        # если количество колонок не совпадает то записываем как ошибку
                                        temp_error_df = pd.DataFrame(
                                            columns=['Название файла', 'Наименование листа', 'Тип ошибки',
                                                     'Описание ошибки'],
                                            data=[[name_file, name_sheet, 'Количество колонок отличается от эталонного',
                                                   f'Ожидалось {dct_df[name_sheet].shape[1]} колонок, а в листе {lst_df.shape[1]}']])  # создаем временный датафрейм. потом надо подумать над словарем
                                        err_df = pd.concat([err_df, temp_error_df],
                                                           ignore_index=True)  # добавляем в датафрейм ошибок
                                        count_errors += 1
                                #
                                # если хоть одна ошибка то проверяем следующий файл
                                if count_errors != 0:
                                    continue
                                # если нет то начинаем обрабатывать листы
                                for name_sheet, skip_r in dct_manage_harvest.items():
                                    temp_df = pd.read_excel(f'{dirpath}/{filename}', sheet_name=name_sheet,
                                                            skiprows=skip_r,
                                                            dtype=str, header=None)  # загружаем датафрейм

                                    if temp_df.shape[1] > 3:
                                        temp_df = temp_df.dropna(axis=0,thresh=2)
                                    temp_df['Номер строки'] = range(1, temp_df.shape[0] + 1)
                                    temp_df['Откуда взяты данные'] = name_file
                                    for row in dataframe_to_rows(temp_df, index=False, header=False):
                                        standard_wb[name_sheet].append(row)  # добавляем данные
                            else:
                                temp_error_df = pd.DataFrame(
                                    columns=['Название файла', 'Наименование листа', 'Тип ошибки', 'Описание ошибки'],
                                    data=[[name_file, '', 'Не совпадает количество или название листов в файле',
                                           f'Листы, которые есть в файле: {",".join(temb_wb.sheetnames)}']])  # создаем временный датафрейм. потом надо подумать над словарем

                                err_df = pd.concat([err_df, temp_error_df],
                                                   ignore_index=True)  # добавляем в датафрейм ошибок

                # # Получаем текущую дату
                current_time = time.strftime('%H_%M_%S %d.%m.%Y')
                standard_wb.save(
                    f'{path_to_end_folder_merger}/Слияние по варианту В Общая таблица от {current_time}.xlsx')  # сохраняем
                err_out_wb = openpyxl.Workbook()  # создаем объект openpyxl для сохранения датафрейма
                for row in dataframe_to_rows(err_df, index=False, header=True):
                    err_out_wb['Sheet'].append(row)  # добавляем данные
                # устанавливаем размер колонок
                err_out_wb['Sheet'].column_dimensions['A'].width = 40
                err_out_wb['Sheet'].column_dimensions['B'].width = 30
                err_out_wb['Sheet'].column_dimensions['C'].width = 55
                err_out_wb['Sheet'].column_dimensions['D'].width = 100
                err_out_wb.save(f'{path_to_end_folder_merger}/Слияние по варианту В Ошибки от {current_time}.xlsx')

        except NameError:
            messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                                 f'Выберите папку с файлами,эталонный файл и папку куда будут генерироваться файлы')
        except PermissionError:
            messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                                 f'Закройте файл выбранный эталонным или файлы из обрабатываемой папки')
        except FileNotFoundError:
            messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                                 f'Выберите файл с параметрами!\n'
                                 f'Если вы выбрали файл с параметрами, а ошибка повторяется,то перенесите папку \n'
                                 f'с файлами которые вы хотите обработать в корень диска. Проблема может быть в \n '
                                 f'в слишком длинном пути к обрабатываемым файлам')
        # except:
        #     logging.exception('AN ERROR HAS OCCURRED')
        #     messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
        #                          'Возникла ошибка!!! Подробности ошибки в файле error.log')
        else:
            messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30',
                                'Создание общей таблицы успешно завершено!!!')


def count_text_value(df):
    """
    Функция для подсчета количества вариантов того или иного показателя
    :param df: датафрейм с сырыми данными. Название показателя значение показателя(строка разделенная ;)
    :return: обработанный датафрейм с мультиндексом, где (Название показателя это индекс верхнего уровня, вариант показателя это индекс второго уровня а значение это сколько раз встречался
    этот вариант в обрабатываемых файлах)
    """
    data = dict()

    #
    for row in df.itertuples():
        value = row[2]
        if type(value) == float or type(value) == int:
            continue
        # Создаем список, разделяя строку по ;
        lst_value = row[2].split(';')[:-1]
        #     # Отрезаем последний элемент, поскольку это пустое значение
        temp_df = pd.DataFrame({'Value': lst_value})
        counts_series = temp_df['Value'].value_counts()
        # Делаем индекс колонкой и превращаем в обычную таблицу
        index_count_values = counts_series.reset_index()
        # Итерируемся по таблице.Это делается чтобы заполнить словарь на основе которого будет создаваться итоговая таблица
        for count_row in index_count_values.itertuples():
            # print(count_row)
            # Заполняем словарь
            data[(row[1], count_row[1])] = count_row[2]
    # Создаем на основе получившегося словаря таблицу
    out_df = pd.Series(data).to_frame().reset_index()
    out_df = out_df.set_index(['level_0', 'level_1'])
    out_df.index.names = ['Название показателя', 'Вариант показателя']
    out_df.rename(columns={0: 'Количество'}, inplace=True)
    return out_df


def check_data(cell, text_mode):
    """
    Функция для проверки значения ячейки. Для обработки пустых значений, строковых значений, дат
    :param cell: значение ячейки
    :return: 0 если значение ячейки не число
            число если значение ячейки число(ха звучит глуповато)
    думаю функция должна работать с дополнительным параметром, от которого будет зависеть подсчет значений навроде галочек или плюсов в анкетах или опросах.
    """
    # Проверяем режим работы. если текстовый, то просто складываем строки
    if text_mode == 'Yes':
        if cell is None:
            return ''
        else:
            temp_str = str(cell)
            return f'{temp_str};'
    # Если режим работы стандартный. Убрал подсчет строк и символов в числовом режиме, чтобы не запутывать.
    else:
        if cell is None:
            return 0
        if type(cell) == int:
            return cell
        elif type(cell) == float:
            return cell
        else:
            return 0


def generate_docs_other():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных Веста Обработка таблиц и создание документов ver 1.30)
    :return:
    """
    try:
        name_column = entry_name_column_data.get()
        name_type_file = entry_type_file.get()
        name_value_column = entry_value_column.get()

        # Считываем данные
        # Добавил параметр dtype =str чтобы данные не преобразовались а использовались так как в таблице
        df = pd.read_excel(name_file_data_doc, dtype=str)
        # Заполняем Nan
        df.fillna(' ',inplace=True)
        lst_date_columns = []

        for idx,column in enumerate(df.columns):
            if 'дата' in column.lower():
                lst_date_columns.append(idx)

        # Конвертируем в пригодный строковый формат
        for i in lst_date_columns:
            df.iloc[:, i] = pd.to_datetime(df.iloc[:, i], errors='coerce', dayfirst=True)
            df.iloc[:, i] = df.iloc[:, i].apply(create_doc_convert_date)


        # Конвертируем датафрейм в список словарей
        data = df.to_dict('records')
        # Получаем состояние  чекбокса объединения файлов в один

        mode_combine = mode_combine_value.get()
        # Получаем состояние чекбокса создания индвидуального файла
        mode_group = mode_group_doc.get()

        # В зависимости от состояния чекбоксов обрабатываем файлы
        if mode_combine == 'No':
            if mode_group == 'No':
                # Создаем в цикле документы
                for idx,row in enumerate(data):
                    doc = DocxTemplate(name_file_template_doc)
                    context = row
                    # print(context)
                    doc.render(context)
                    # Сохраняенм файл
                    # получаем название файла и убираем недопустимые символы < > : " /\ | ? *
                    name_file = f'{name_type_file} {row[name_column]}'
                    name_file = re.sub(r'[<> :"?*|\\/]',' ',name_file)
                    # проверяем файл на наличие, если файл с таким названием уже существует то добавляем окончание
                    if os.path.exists(f'{path_to_end_folder_doc}/{name_file}.docx'):
                        doc.save(f'{path_to_end_folder_doc}/{name_file}_{idx}.docx')

                    doc.save(f'{path_to_end_folder_doc}/{name_file}.docx')
            else:
                # Отбираем по значению строку

                single_df = df[df[name_column] == name_value_column]
                # Конвертируем датафрейм в список словарей
                single_data = single_df.to_dict('records')
                # Проверяем количество найденных совпадений
                # очищаем от запрещенных символов
                name_file = f'{name_type_file} {name_value_column}'
                name_file = re.sub(r'[<> :"?*|\\/]', ' ', name_file)
                if len(single_data) == 1:
                    for row in single_data:
                        doc = DocxTemplate(name_file_template_doc)
                        doc.render(row)
                        # Сохраняенм файл
                        doc.save(f'{path_to_end_folder_doc}/{name_file}.docx')
                elif len(single_data) > 1:
                    for idx,row in enumerate(single_data):
                        doc = DocxTemplate(name_file_template_doc)
                        doc.render(row)
                        # Сохраняем файл

                        doc.save(f'{path_to_end_folder_doc}/{name_file}_{idx}.docx')
                else:
                    raise NotFoundValue



        else:
            if mode_group == 'No':
                # Список с созданными файлами
                files_lst = []
                # Создаем временную папку
                with tempfile.TemporaryDirectory() as tmpdirname:
                    print('created temporary directory', tmpdirname)
                    # Создаем и сохраняем во временную папку созданные документы Word
                    for row in data:
                        doc = DocxTemplate(name_file_template_doc)
                        context = row
                        doc.render(context)
                        # Сохраняем файл
                        #очищаем от запрещенных символов
                        name_file = f'{row[name_column]}'
                        name_file = re.sub(r'[<> :"?*|\\/]', ' ', name_file)

                        doc.save(f'{tmpdirname}/{name_file}.docx')
                        # Добавляем путь к файлу в список
                        files_lst.append(f'{tmpdirname}/{name_file}.docx')
                    # Получаем базовый файл
                    main_doc = files_lst.pop(0)
                    # Запускаем функцию
                    combine_all_docx(main_doc, files_lst)
            else:
                raise CheckBoxException

    except NameError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'В таблице не найдена указанная колонка {e.args}')
    except PermissionError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Закройте все файлы Word созданные Вестой')
        logging.exception('AN ERROR HAS OCCURRED')
    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except CheckBoxException:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Уберите галочку из чекбокса Поставьте галочку, если вам нужно создать один документ\nдля конкретного значения (например для определенного ФИО)'
                             )
    except NotFoundValue:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Указанное значение не найдено в выбранной колонке\nПроверьте наличие такого значения в таблице'
                             )
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')

    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30', 'Создание документов завершено!')


def check_date_columns(i, value):
    """
    Функция для проверки типа колонки. Необходимо найти колонки с датой
    :param i:
    :param value:
    :return:
    """
    try:
        itog = pd.to_datetime(str(value), infer_datetime_format=True)
    except:
        pass
    else:
        return i


def set_rus_locale():
    """
    Функция чтобы можно было извлечь русские названия месяцев
    """
    locale.setlocale(
        locale.LC_ALL,
        'rus_rus' if sys.platform == 'win32' else 'ru_RU.UTF-8')


def combine_all_docx(filename_master, files_lst):
    """
    Функция для объединения файлов Word взято отсюда
    https://stackoverflow.com/questions/24872527/combine-word-document-using-python-docx
    :param filename_master: базовый файл
    :param files_list: список с созданными файлами
    :return: итоговый файл
    """
    # Получаем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    number_of_sections = len(files_lst)
    # Открываем и обрабатываем базовый файл
    master = Document(filename_master)
    composer = Composer(master)
    # Перебираем и добавляем файлы к базовому
    for i in range(0, number_of_sections):
        doc_temp = Document(files_lst[i])
        composer.append(doc_temp)
    # Сохраняем файл
    composer.save(f"{path_to_end_folder_doc}/Объединеный файл от {current_time}.docx")


def calculate_age(born):
    """
    Функция для расчета текущего возраста взято с https://stackoverflow.com/questions/2217488/age-from-birthdate-in-python/9754466#9754466
    :param born: дата рождения
    :return: возраст
    """

    try:

        # today = date.today()
        selected_date = pd.to_datetime(raw_selected_date, dayfirst=True)
        # return today.year - born.year - ((today.month, today.day) < (born.month, born.day))
        return selected_date.year - born.year - ((selected_date.month, selected_date.day) < (born.month, born.day))

    except ValueError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Введена некорректная дата относительно которой нужно провести обработку\nПример корректной даты 01.09.2022')
        logging.exception('AN ERROR HAS OCCURRED')
        quit()


def convert_date(cell):
    """
    Функция для конвертации даты в формате 1957-05-10 в формат 10.05.1957(строковый)
    """

    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date

    except TypeError:
        print(cell)
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Проверьте правильность заполнения ячеек с датой!!!')
        logging.exception('AN ERROR HAS OCCURRED')
        quit()


def create_doc_convert_date(cell):
    """
    Функция для конвертации даты при создании документов
    :param cell:
    :return:
    """
    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date
    except ValueError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'
    except TypeError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'



def processing_date_column(df, lst_columns):
    """
    Функция для обработки столбцов с датами. конвертация в строку формата ДД.ММ.ГГГГ
    """
    # получаем первую строку
    first_row = df.iloc[0, lst_columns]

    lst_first_row = list(first_row)  # Превращаем строку в список
    lst_date_columns = []  # Создаем список куда будем сохранять колонки в которых находятся даты
    tupl_row = list(zip(lst_columns,
                        lst_first_row))  # Создаем список кортежей формата (номер колонки,значение строки в этой колонке)

    for idx, value in tupl_row:  # Перебираем кортеж
        result = check_date_columns(idx, value)  # проверяем является ли значение датой
        if result:  # если да то добавляем список порядковый номер колонки
            lst_date_columns.append(result)
        else:  # иначе проверяем следующее значение
            continue
    for i in lst_date_columns:  # Перебираем список с колонками дат, превращаем их в даты и конвертируем в нужный строковый формат
        df.iloc[:, i] = pd.to_datetime(df.iloc[:, i], errors='coerce', dayfirst=True)
        df.iloc[:, i] = df.iloc[:, i].apply(create_doc_convert_date)


def extract_number_month(cell):
    """
    Функция для извлечения номера месяца
    """
    return cell.month


def extract_name_month(cell):
    """
    Функция для извлечения названия месяца
    Взято отсюда https://ru.stackoverflow.com/questions/1045154/Вывод-русских-символов-из-pd-timestamp-month-name
    """
    return cell.month_name(locale='Russian')


def extract_year(cell):
    """
    Функция для извлечения года рождения
    """
    return cell.year


def calculate_date():
    """
    Функция для разбиения по категориям, подсчета текущего возраста и выделения месяца,года
    :return:
    """
    try:
        # делаем глобальным значение даты.Дада я знаю что это костыль
        global raw_selected_date
        raw_selected_date = entry_date.get()

        name_column = entry_name_column.get()
        # Устанавливаем русскую локаль
        set_rus_locale()

        # Считываем файл
        df = pd.read_excel(name_file_data_date)
        # Конвертируем его в формат даты
        # В случае ошибок заменяем значение NaN
        df[name_column] = pd.to_datetime(df[name_column], dayfirst=True, errors='coerce')
        # Создаем шрифт которым будем выделять названия таблиц
        font_name_table = Font(name='Arial Black', size=15, italic=True)

        # Создаем файл excel
        wb = openpyxl.Workbook()
        # Создаем листы
        # Переименовываем лист чтобы в итоговом файле не было пустого листа
        ren_sheet = wb['Sheet']
        ren_sheet.title = 'Итоговая таблица'

        # wb.create_sheet(title='Итоговая таблица', index=0)
        wb.create_sheet(title='Свод по возрастам', index=1)
        wb.create_sheet(title='Свод по месяцам', index=2)
        wb.create_sheet(title='Свод по годам', index=3)
        wb.create_sheet(title='Свод по 1-ПК', index=4)
        wb.create_sheet(title='Свод по 1-ПО', index=5)
        wb.create_sheet(title='Свод по СПО-1', index=6)
        wb.create_sheet(title='Свод по категориям Росстата', index=7)

        # Подсчитываем текущий возраст
        df['Текущий возраст'] = df[name_column].apply(calculate_age)

        # Получаем номер месяца
        df['Порядковый номер месяца рождения'] = df[name_column].apply(extract_number_month)

        # Получаем название месяца
        df['Название месяца рождения'] = df[name_column].apply(extract_name_month)

        # Получаем год рождения
        df['Год рождения'] = df[name_column].apply(extract_year)

        # Присваиваем категорию по 1-ПК
        df['1-ПК Категория'] = pd.cut(df['Текущий возраст'], [0, 24, 29, 34, 39, 44, 49, 54, 59, 64, 101, 10000],
                                      labels=['моложе 25 лет', '25-29', '30-34', '35-39',
                                              '40-44', '45-49', '50-54', '55-59', '60-64',
                                              '65 и более',
                                              'Возраст  больше 101'])
        # Приводим к строковому виду, иначе не запишется на лист
        df['1-ПК Категория'] = df['1-ПК Категория'].astype(str)

        # Присваиваем категорию по 1-ПО
        df['1-ПО Категория'] = pd.cut(df['Текущий возраст'],
                                      [0, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25,
                                       26, 27, 28,
                                       29, 34, 39, 44, 49, 54, 59, 64, 101],
                                      labels=['моложе 14 лет', '14 лет', '15 лет',
                                              '16 лет',
                                              '17 лет', '18 лет', '19 лет', '20 лет',
                                              '21 год', '22 года',
                                              '23 года', '24 года', '25 лет',
                                              '26 лет', '27 лет', '28 лет', '29 лет',
                                              '30-34 лет',
                                              '35-39 лет', '40-44 лет', '45-49 лет',
                                              '50-54 лет',
                                              '55-59 лет',
                                              '60-64 лет',
                                              '65 лет и старше'])
        # Приводим к строковому виду, иначе не запишется на лист
        df['1-ПО Категория'] = df['1-ПО Категория'].astype(str)

        # Присваиваем категорию по 1-СПО
        df['СПО-1 Категория'] = pd.cut(df['Текущий возраст'],
                                       [0, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 34,
                                        39,
                                        101],
                                       labels=['моложе 13 лет', '13 лет', '14 лет', '15 лет', '16 лет', '17 лет',
                                               '18 лет',
                                               '19 лет', '20 лет'
                                           , '21 год', '22 года', '23 года', '24 года', '25 лет', '26 лет', '27 лет',
                                               '28 лет',
                                               '29 лет',
                                               '30-34 лет', '35-39 лет', '40 лет и старше'])
        ## Приводим к строковому виду, иначе не запишется на лист
        df['СПО-1 Категория'] = df['СПО-1 Категория'].astype(str)

        # Присваиваем категорию по Росстату
        df['Росстат Категория'] = pd.cut(df['Текущий возраст'],
                                         [0, 4, 9, 14, 19, 24, 29, 34, 39, 44, 49, 54, 59, 64, 69, 200],
                                         labels=['0-4', '5-9', '10-14', '15-19', '20-24', '25-29', '30-34',
                                                 '35-39', '40-44', '45-49', '50-54', '55-59', '60-64', '65-69',
                                                 '70 лет и старше'])
        ## Приводим к строковому виду, иначе не запишется на лист
        df['Росстат Категория'] = df['Росстат Категория'].astype(str)

        # Заполняем пустые строки
        df.fillna('Не заполнено!!!', inplace=True)

        # заполняем сводные таблицы
        # Сводная по возрастам

        df_svod_by_age = df.groupby(['Текущий возраст']).agg({name_column: 'count'})
        df_svod_by_age.columns = ['Количество']

        for r in dataframe_to_rows(df_svod_by_age, index=True, header=True):
            wb['Свод по возрастам'].append(r)

        # Сводная по месяцам
        df_svod_by_month = df.groupby(['Название месяца рождения']).agg({name_column: 'count'})
        df_svod_by_month.columns = ['Количество']

        # Сортируем индекс чтобы месяцы шли в хоронологическом порядке
        # Взял отсюда https://stackoverflow.com/questions/40816144/pandas-series-sort-by-month-index
        df_svod_by_month.index = pd.CategoricalIndex(df_svod_by_month.index,
                                                     categories=['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                                                                 'Июль',
                                                                 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'],
                                                     ordered=True)
        df_svod_by_month.sort_index(inplace=True)

        for r in dataframe_to_rows(df_svod_by_month, index=True, header=True):
            wb['Свод по месяцам'].append(r)

        # Сводная по годам
        df_svod_by_year = df.groupby(['Год рождения']).agg({name_column: 'count'})
        df_svod_by_year.columns = ['Количество']

        for r in dataframe_to_rows(df_svod_by_year, index=True, header=True):
            wb['Свод по годам'].append(r)

        # Сводная по 1-ПК
        df_svod_by_1PK = df.groupby(['1-ПК Категория']).agg({name_column: 'count'})
        df_svod_by_1PK.columns = ['Количество']

        for r in dataframe_to_rows(df_svod_by_1PK, index=True, header=True):
            wb['Свод по 1-ПК'].append(r)

        # Сводная по 1-ПО
        df_svod_by_1PO = df.groupby(['1-ПО Категория']).agg({name_column: 'count'})
        df_svod_by_1PO.columns = ['Количество']

        for r in dataframe_to_rows(df_svod_by_1PO, index=True, header=True):
            wb['Свод по 1-ПО'].append(r)

        # Сводная по СПО-1
        df_svod_by_SPO1 = df.groupby(['СПО-1 Категория']).agg({name_column: 'count'})
        df_svod_by_SPO1.columns = ['Количество']

        for r in dataframe_to_rows(df_svod_by_SPO1, index=True, header=True):
            wb['Свод по СПО-1'].append(r)

        # Сводная по Росстату
        df_svod_by_Ros = df.groupby(['Росстат Категория']).agg({name_column: 'count'})
        df_svod_by_Ros.columns = ['Количество']

        # Сортируем индекс
        df_svod_by_Ros.index = pd.CategoricalIndex(df_svod_by_Ros.index,
                                                   categories=['0-4', '5-9', '10-14', '15-19', '20-24', '25-29',
                                                               '30-34',
                                                               '35-39', '40-44', '45-49', '50-54', '55-59', '60-64',
                                                               '65-69',
                                                               '70 лет и старше', 'nan'],
                                                   ordered=True)
        df_svod_by_Ros.sort_index(inplace=True)

        for r in dataframe_to_rows(df_svod_by_Ros, index=True, header=True):
            wb['Свод по категориям Росстата'].append(r)

        for r in dataframe_to_rows(df, index=False, header=True):
            wb['Итоговая таблица'].append(r)

        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Сохраняем итоговый файл
        wb.save(f'{path_to_end_folder_date}/Результат обработки колонки {name_column} от {current_time}.xlsx')
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'В таблице нет такой колонки!\nПроверьте написание названия колонки')
        logging.exception('AN ERROR HAS OCCURRED')
    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')

    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')
    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30', 'Данные успешно обработаны')


def groupby_category():
    """
    Функция для подсчета выбранной колонки по категориям
    :return:
    """
    try:
        df = pd.read_excel(name_file_data_groupby)
        df.columns = list(map(str, list(df.columns)))
        # Создаем шрифт которым будем выделять названия таблиц
        font_name_table = Font(name='Arial Black', size=15, italic=True)
        # Создаем файл excel
        wb = openpyxl.Workbook()

        # Проверяем наличие возможных дубликатов ,котороые могут получиться если обрезать по 30 символов
        lst_length_column = [column[:30] for column in df.columns]
        check_dupl_length = [k for k,v in Counter(lst_length_column).items() if v>1]

        # проверяем наличие объединенных ячеек
        check_merge = [column for column in df.columns if 'Unnamed' in column]
        # если есть хоть один Unnamed то просто заменяем названия колонок на Колонка №цифра
        if check_merge or check_dupl_length:
            df.columns = [f'Колонка №{i}' for i in range(1, df.shape[1] + 1)]
        # очищаем названия колонок от символов */\ []''
        # Создаем регулярное выражение
        pattern_symbols = re.compile(r"[/*'\[\]/\\]")
        clean_df_columns = [re.sub(pattern_symbols,'',column) for column in df.columns]
        df.columns = clean_df_columns

        # Добавляем столбец для облегчения подсчета по категориям
        df['Для подсчета'] = 1

        # Создаем листы
        for idx, name_column in enumerate(df.columns):
            # Делаем короткое название не более 30 символов
            wb.create_sheet(title=name_column[:30], index=idx)

        for idx, name_column in enumerate(df.columns):
            group_df = df.groupby([name_column]).agg({'Для подсчета': 'sum'})
            group_df.columns = ['Количество']

            # Сортируем по убыванию
            group_df.sort_values(by=['Количество'], inplace=True, ascending=False)

            for r in dataframe_to_rows(group_df, index=True, header=True):
                if len(r) != 1:
                    wb[name_column[:30]].append(r)
            wb[name_column[:30]].column_dimensions['A'].width = 50

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Удаляем листы
        del wb['Sheet']
        del wb['Для подсчета']
        # Сохраняем итоговый файл
        wb.save(
            f'{path_to_end_folder_groupby}/Подсчет частоты значений для всех колонок таблицы от {current_time}.xlsx')

    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')

    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')
    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30', 'Данные успешно обработаны')


def groupby_stat():
    """
    Функция для подсчета выбранной колонки по количественным показателям(сумма,среднее,медиана,мин,макс)
    :return:
    """

    try:
        df = pd.read_excel(name_file_data_groupby)
        # Делаем названия колонок строковыми
        df.columns = list(map(str, list(df.columns)))

        # Создаем шрифт которым будем выделять названия таблиц
        font_name_table = Font(name='Arial Black', size=15, italic=True)
        # Создаем файл excel
        wb = openpyxl.Workbook()

        # Проверяем наличие возможных дубликатов ,котороые могут получиться если обрезать по 30 символов
        lst_length_column = [column[:30] for column in df.columns]
        check_dupl_length = [k for k,v in Counter(lst_length_column).items() if v>1]

        # проверяем наличие объединенных ячеек
        check_merge = [column for column in df.columns if 'Unnamed' in column]
        # если есть хоть один Unnamed или дубликат то просто заменяем названия колонок на Колонка №цифра
        if check_merge or check_dupl_length:
            df.columns = [f'Колонка №{i}' for i in range(1, df.shape[1] + 1)]

        # очищаем названия колонок от символов */\ []''
        # Создаем регулярное выражение
        pattern_symbols = re.compile(r"[/*'\[\]/\\]")
        clean_df_columns = [re.sub(pattern_symbols,'',column) for column in df.columns]
        df.columns = clean_df_columns


        # Добавляем столбец для облегчения подсчета по категориям
        df['Итого'] = 1

        # Создаем листы
        for idx, name_column in enumerate(df.columns):
            # Делаем короткое название не более 30 символов
            wb.create_sheet(title=name_column[:30], index=idx)

        for idx, name_column in enumerate(df.columns):
            group_df = df[name_column].describe().to_frame()
            if group_df.shape[0] == 8:
                # подсчитаем сумму
                all_sum = df[name_column].sum()
                dct_row = {name_column: all_sum}
                row = pd.DataFrame(data=dct_row, index=['Сумма'])
                # Добавим в датафрейм
                group_df = pd.concat([group_df, row], axis=0)

                # Обновим названия индексов
                group_df.index = ['Количество значений', 'Среднее', 'Стандартное отклонение', 'Минимальное значение',
                                  '25%(Первый квартиль)', 'Медиана', '75%(Третий квартиль)', 'Максимальное значение',
                                  'Сумма']

            elif group_df.shape[0] == 4:
                group_df.index = ['Количество значений', 'Количество уникальных значений', 'Самое частое значение',
                                  'Количество повторений самого частого значения', ]
            for r in dataframe_to_rows(group_df, index=True, header=True):
                if len(r) != 1:
                    wb[name_column[:30]].append(r)
            wb[name_column[:30]].column_dimensions['A'].width = 50

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Удаляем лист
        del wb['Sheet']
        del wb['Итого']
        # Сохраняем итоговый файл
        wb.save(
            f'{path_to_end_folder_groupby}/Подсчет базовых статистик для всех колонок таблицы от {current_time}.xlsx')



    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')

    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')

    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30', 'Данные успешно обработаны')


def processing_comparison():
    """
    Функция для сравнения 2 колонок
    :return:
    """
    try:
        # Получаем значения текстовых полей
        first_sheet_name = str(entry_first_sheet_name.get())
        second_sheet_name = str(entry_second_sheet_name.get())
        # загружаем файлы
        first_df = pd.read_excel(name_first_file_comparison, sheet_name=first_sheet_name, dtype=str,
                                 keep_default_na=False)
        # получаем имя файла
        name_first_df = name_first_file_comparison.split('/')[-1]
        name_first_df = name_first_df.split('.xlsx')[0]

        second_df = pd.read_excel(name_second_file_comparison, sheet_name=second_sheet_name, dtype=str,
                                  keep_default_na=False)
        # получаем имя файла
        name_second_df = name_second_file_comparison.split('/')[-1]
        name_second_df = name_second_df.split('.xlsx')[0]

        params = pd.read_excel(file_params, header=None, keep_default_na=False)

        # Преврашаем каждую колонку в список
        params_first_columns = params[0].tolist()
        params_second_columns = params[1].tolist()

        # Конвертируем в инт заодно проверяя корректность введенных данных
        int_params_first_columns = convert_params_columns_to_int(params_first_columns)
        int_params_second_columns = convert_params_columns_to_int(params_second_columns)

        # Отнимаем 1 от каждого значения чтобы привести к питоновским индексам
        int_params_first_columns = list(map(lambda x: x - 1, int_params_first_columns))
        int_params_second_columns = list(map(lambda x: x - 1, int_params_second_columns))

        # Конвертируем нужные нам колонки в str
        convert_columns_to_str(first_df, int_params_first_columns)
        convert_columns_to_str(second_df, int_params_second_columns)

        # Проверяем наличие колонок с датами в списке колонок для объединения чтобы привести их в нормальный вид
        for number_column_params in int_params_first_columns:
            if 'дата' in first_df.columns[number_column_params].lower():
                first_df.iloc[:, number_column_params] = pd.to_datetime(first_df.iloc[:, number_column_params],
                                                                        errors='coerce', dayfirst=True)
                first_df.iloc[:, number_column_params] = first_df.iloc[:, number_column_params].apply(
                    create_doc_convert_date)

        for number_column_params in int_params_second_columns:
            if 'дата' in second_df.columns[number_column_params].lower():
                second_df.iloc[:, number_column_params] = pd.to_datetime(second_df.iloc[:, number_column_params],
                                                                         errors='coerce', dayfirst=True)
                second_df.iloc[:, number_column_params] = second_df.iloc[:, number_column_params].apply(
                    create_doc_convert_date)

        # в этом месте конвертируем даты в формат ДД.ММ.ГГГГ
        # processing_date_column(first_df, int_params_first_columns)
        # processing_date_column(second_df, int_params_second_columns)

        # Проверяем наличие колонки _merge
        if '_merge' in first_df.columns:
            first_df.drop(columns=['_merge'], inplace=True)
        if '_merge' in second_df.columns:
            second_df.drop(columns=['_merge'], inplace=True)
        # Проверяем наличие колонки ID
        if 'ID_объединения' in first_df.columns:
            first_df.drop(columns=['ID_объединения'], inplace=True)
        if 'ID_объединения' in second_df.columns:
            second_df.drop(columns=['ID_объединения'], inplace=True)

        # Создаем в каждом датафрейме колонку с айди путем склеивания всех нужных колонок в одну строку
        first_df['ID_объединения'] = first_df.iloc[:, int_params_first_columns].sum(axis=1)
        second_df['ID_объединения'] = second_df.iloc[:, int_params_second_columns].sum(axis=1)

        first_df['ID_объединения'] = first_df['ID_объединения'].apply(lambda x: x.replace(' ', ''))
        second_df['ID_объединения'] = second_df['ID_объединения'].apply(lambda x: x.replace(' ', ''))


        # В результат объединения попадают совпадающие по ключу записи обеих таблиц и все строки из этих двух таблиц, для которых пар не нашлось. Порядок таблиц в запросе не

        # Создаем документ
        wb = openpyxl.Workbook()
        # создаем листы
        ren_sheet = wb['Sheet']
        ren_sheet.title = 'Таблица 1'
        wb.create_sheet(title='Таблица 2', index=1)
        wb.create_sheet(title='Совпадающие данные', index=2)
        wb.create_sheet(title='Обновленная таблица', index=3)
        wb.create_sheet(title='Объединённая таблица', index=4)


        # Создаем переменные содержащие в себе количество колонок в базовых датареймах
        first_df_quantity_cols = len(first_df.columns)  # не забываем что там добавилась колонка ID

        # Проводим слияние
        itog_df = pd.merge(first_df, second_df, how='outer', left_on=['ID_объединения'], right_on=['ID_объединения'],
                           indicator=True)

        # копируем в отдельный датафрейм для создания таблицы с обновлениями
        update_df = itog_df.copy()

        # Записываем каждый датафрейм в соответсвующий лист
        # Левая таблица
        left_df = itog_df[itog_df['_merge'] == 'left_only']
        left_df.drop(['_merge'], axis=1, inplace=True)

        # Удаляем колонки второй таблицы чтобы не мешались
        left_df.drop(left_df.iloc[:, first_df_quantity_cols:], axis=1, inplace=True)

        # Переименовываем колонки у которых были совпадение во второй таблице, в таких колонках есть добавление _x
        clean_left_columns = list(map(lambda x: x[:-2] if '_x' in x else x, list(left_df.columns)))
        left_df.columns = clean_left_columns
        for r in dataframe_to_rows(left_df, index=False, header=True):
            wb['Таблица 1'].append(r)

        right_df = itog_df[itog_df['_merge'] == 'right_only']
        right_df.drop(['_merge'], axis=1, inplace=True)

        # Удаляем колонки первой таблицы таблицы чтобы не мешались
        right_df.drop(right_df.iloc[:, :first_df_quantity_cols - 1], axis=1, inplace=True)

        # Переименовываем колонки у которых были совпадение во второй таблице, в таких колонках есть добавление _x
        clean_right_columns = list(map(lambda x: x[:-2] if '_y' in x else x, list(right_df.columns)))
        right_df.columns = clean_right_columns

        for r in dataframe_to_rows(right_df, index=False, header=True):
            wb['Таблица 2'].append(r)

        both_df = itog_df[itog_df['_merge'] == 'both']
        both_df.drop(['_merge'], axis=1, inplace=True)
        # Очищаем от _x  и _y
        clean_both_columns = clean_ending_columns(list(both_df.columns), name_first_df, name_second_df)
        both_df.columns = clean_both_columns

        for r in dataframe_to_rows(both_df, index=False, header=True):
            wb['Совпадающие данные'].append(r)

        # Сохраняем общую таблицу
        # Заменяем названия индикаторов на более понятные
        itog_df['_merge'] = itog_df['_merge'].apply(lambda x: 'Данные из первой таблицы' if x == 'left_only' else
        ('Данные из второй таблицы' if x == 'right_only' else 'Совпадающие данные'))
        itog_df['_merge'] = itog_df['_merge'].astype(str)

        clean_itog_df = clean_ending_columns(list(itog_df.columns), name_first_df, name_second_df)
        itog_df.columns = clean_itog_df
        for r in dataframe_to_rows(itog_df, index=False, header=True):
            wb['Объединённая таблица'].append(r)

        # получаем список с совпадающими колонками первой таблицы
        first_df_columns = [column for column in list(update_df.columns) if str(column).endswith('_x')]
        # получаем список с совпадающими колонками второй таблицы
        second_df_columns = [column for column in list(update_df.columns) if str(column).endswith('_y')]
        # Создаем из списка совпадающих колонок второй таблицы словарь, чтобы было легче обрабатывать
        # да конечно можно было сделать в одном выражении но как я буду читать это через 2 недели?
        dct_second_columns = {column.split('_y')[0]: column for column in second_df_columns}

        for column in first_df_columns:
            # очищаем от _x
            name_column = column.split('_x')[0]
            # Обновляем значение в случае если в колонке _merge стоит both, иначе оставляем старое значение,
            # Чтобы обновить значение в ячейке, во второй таблице не должно быть пустого значения или пробела в аналогичной колонке

            update_df[column] = np.where(
                (update_df['_merge'] == 'both') & (update_df[dct_second_columns[name_column]]) & (
                            update_df[dct_second_columns[name_column]] != ' '),
                update_df[dct_second_columns[name_column]], update_df[column])

            # Удаляем колонки с _y
        update_df.drop(columns=[column for column in update_df.columns if column.endswith('_y')], inplace=True)

        # Переименовываем колонки с _x
        update_df.columns = list(map(lambda x: x[:-2] if x.endswith('_x') else x, update_df.columns))

        # удаляем строки с _merge == right_only
        update_df = update_df[update_df['_merge'] != 'right_only']

        # Удаляем служебные колонки
        update_df.drop(columns=['ID_объединения', '_merge'], inplace=True)

        # используем уже созданный датафрейм right_df Удаляем лишнюю колонку в right_df
        right_df.drop(columns=['ID_объединения'], inplace=True)

        # Добавляем нехватающие колонки
        new_right_df = right_df.reindex(columns=update_df.columns, fill_value=None)

        update_df = pd.concat([update_df, new_right_df])

        for r in dataframe_to_rows(update_df, index=False, header=True):
            wb['Обновленная таблица'].append(r)

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Сохраняем итоговый файл
        wb.save(f'{path_to_end_folder_comparison}/Результат слияния 2 таблиц от {current_time}.xlsx')
        # Сохраняем отдельно обновленную таблицу
        update_df.to_excel(
            f'{path_to_end_folder_comparison}/Таблица с обновленными данными и колонками от {current_time}.xlsx',
            index=False)

    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'В таблице нет такой колонки!\nПроверьте написание названия колонки')
        logging.exception('AN ERROR HAS OCCURRED')
    except ValueError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'В таблице нет листа с таким названием!\nПроверьте написание названия листа')
        logging.exception('AN ERROR HAS OCCURRED')
    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')
    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30', 'Данные успешно обработаны')

def clean_ending_columns(lst_columns:list,name_first_df,name_second_df):
    """
    Функция для очистки колонок таблицы с совпадающими данными от окончаний _x _y

    :param lst_columns:
    :param time_generate
    :param name_first_df
    :param name_second_df
    :return:
    """
    out_columns = [] # список для очищенных названий
    for name_column in lst_columns:
        if '_x' in name_column:
            # если они есть то проводим очистку и добавление времени
            cut_name_column = name_column[:-2] # обрезаем
            temp_name = f'{cut_name_column}_{name_first_df}' # соединяем
            out_columns.append(temp_name) # добавляем
        elif '_y' in name_column:
            cut_name_column = name_column[:-2]  # обрезаем
            temp_name = f'{cut_name_column}_{name_second_df}'  # соединяем
            out_columns.append(temp_name)  # добавляем
        else:
            out_columns.append(name_column)
    return out_columns


"""
Функции для склонения ФИО по падежам
"""
def select_data_decl_case():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_decl_case
    # Получаем путь к файлу
    data_decl_case = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_end_folder_decl_case():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_decl_case
    path_to_end_folder_decl_case = filedialog.askdirectory()


def capitalize_double_name(word):
    """
    Функция для того чтобы в двойных именах и фамилиях вторая часть была также с большой буквы
    """
    lst_word = word.split('-')  # сплитим по дефису
    if len(lst_word) == 1:  # если длина списка равна 1 то это не двойная фамилия и просто возвращаем слово

        return word
    elif len(lst_word) == 2:
        first_word = lst_word[0].capitalize()  # делаем первую букву слова заглавной а остальные строчными
        second_word = lst_word[1].capitalize()
        return f'{first_word}-{second_word}'
    else:
        return 'Не удалось просклонять'


def case_lastname(maker, lastname, gender, case: Case):
    """
    Функция для обработки и склонения фамилии. Это нужно для обработки случаев двойной фамилии
    """

    lst_lastname = lastname.split('-')  # сплитим по дефису

    if len(lst_lastname) == 1:  # если длина списка равна 1 то это не двойная фамилия и просто обрабатываем слово
        case_result_lastname = maker.make(NamePart.LASTNAME, gender, case, lastname)
        return case_result_lastname
    elif len(lst_lastname) == 2:
        first_lastname = lst_lastname[0].capitalize()  # делаем первую букву слова заглавной а остальные строчными
        second_lastname = lst_lastname[1].capitalize()
        # Склоняем по отдельности
        first_lastname = maker.make(NamePart.LASTNAME, gender, case, first_lastname)
        second_lastname = maker.make(NamePart.LASTNAME, gender, case, second_lastname)

        return f'{first_lastname}-{second_lastname}'


def detect_gender(detector, lastname, firstname, middlename):
    """
    Функция для определения гендера слова
    """
    #     detector = PetrovichGenderDetector() # создаем объект детектора
    try:
        gender_result = detector.detect(lastname=lastname, firstname=firstname, middlename=middlename)
        return gender_result
    except StopIteration:  # если не удалось определить то считаем что гендер андрогинный
        return Gender.ANDROGYNOUS


def decl_on_case(fio: str, case: Case) -> str:
    """
    Функция для склонения ФИО по падежам
    """
    fio = fio.strip()  # очищаем строку от пробельных символов с начала и конца
    part_fio = fio.split()  # разбиваем по пробелам создавая список где [0] это Фамилия,[1]-Имя,[2]-Отчество

    if len(part_fio) == 3:  # проверяем на длину и обрабатываем только те что имеют длину 3 во всех остальных случаях просим просклонять самостоятельно
        maker = PetrovichDeclinationMaker()  # создаем объект класса
        lastname = part_fio[0].capitalize()  # Фамилия
        firstname = part_fio[1].capitalize()  # Имя
        middlename = part_fio[2].capitalize()  # Отчество

        # Определяем гендер для корректного склонения
        detector = PetrovichGenderDetector()  # создаем объект детектора
        gender = detect_gender(detector, lastname, firstname, middlename)
        # Склоняем

        case_result_lastname = case_lastname(maker, lastname, gender, case)  # обрабатываем фамилию
        case_result_firstname = maker.make(NamePart.FIRSTNAME, gender, case, firstname)
        case_result_firstname = capitalize_double_name(case_result_firstname)  # обрабатываем случаи двойного имени
        case_result_middlename = maker.make(NamePart.MIDDLENAME, gender, case, middlename)
        # Возвращаем результат
        result_fio = f'{case_result_lastname} {case_result_firstname} {case_result_middlename}'
        return result_fio

    else:
        return 'Проверьте количество слов, должно быть 3 разделенных пробелами слова'

def create_initials(cell,checkbox,space):
    """
    Функция для создания инициалов
    """
    lst_fio = cell.split(' ') # сплитим по пробелу
    if len(lst_fio) == 3: # проверяем на стандартный размер в 3 слова иначе ничего не меняем
        if checkbox == 'ФИ':
            if space == 'без пробела':
                # возвращаем строку вида Иванов И.И.
                return f'{lst_fio[0]} {lst_fio[1][0].upper()}.{lst_fio[2][0].upper()}.'
            else:
                # возвращаем строку с пробелом после имени Иванов И. И.
                return f'{lst_fio[0]} {lst_fio[1][0].upper()}. {lst_fio[2][0].upper()}.'

        else:
            if space == 'без пробела':
                # И.И. Иванов
                return f'{lst_fio[1][0].upper()}.{lst_fio[2][0].upper()}. {lst_fio[0]}'
            else:
                # И. И. Иванов
                return f'{lst_fio[1][0].upper()}. {lst_fio[2][0].upper()}. {lst_fio[0]}'



    else:
        return cell

def process_decl_case():
    """
    Функция для проведения склонения ФИО по падежам
    :return:
    """
    try:
        fio_column = decl_case_entry_fio.get()

        df = pd.read_excel(data_decl_case, dtype={fio_column: str})

        temp_df = pd.DataFrame()  # временный датафрейм для хранения колонок просклоненных по падежам

        # Получаем номер колонки с фио которые нужно обработать
        lst_columns = list(df.columns)  # Превращаем в список
        index_fio_column = lst_columns.index(fio_column)  # получаем индекс

        # Обрабатываем nan значения и те которые обозначены пробелом
        df[fio_column].fillna('Не заполнено', inplace=True)
        df[fio_column] = df[fio_column].apply(lambda x: x.strip())
        df[fio_column] = df[fio_column].apply(
            lambda x: x if x else 'Не заполнено')  # Если пустая строка то заменяем на значение Не заполнено

        temp_df['Родительный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.GENITIVE))
        temp_df['Дательный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.DATIVE))
        temp_df['Винительный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.ACCUSATIVE))
        temp_df['Творительный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.INSTRUMENTAL))
        temp_df['Предложный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.PREPOSITIONAL))
        temp_df['Фамилия_инициалы'] = df[fio_column].apply(lambda x: create_initials(x,'ФИ','без пробела'))
        temp_df['Инициалы_фамилия'] = df[fio_column].apply(lambda x: create_initials(x,'ИФ','без пробела'))
        temp_df['Фамилия_инициалы_пробел'] = df[fio_column].apply(lambda x: create_initials(x,'ФИ','пробел'))
        temp_df['Инициалы_фамилия_пробел'] = df[fio_column].apply(lambda x: create_initials(x,'ИФ','пробел'))

        # Создаем колонки для склонения фамилий с иницалами родительный падеж
        temp_df['Фамилия_инициалы_род_падеж'] = temp_df['Родительный_падеж'].apply(lambda x: create_initials(x,'ФИ','без пробела'))
        temp_df['Фамилия_инициалы_род_падеж_пробел'] = temp_df['Родительный_падеж'].apply(lambda x: create_initials(x, 'ФИ', 'пробел'))
        temp_df['Инициалы_фамилия_род_падеж'] = temp_df['Родительный_падеж'].apply(lambda x: create_initials(x,'ИФ','без пробела'))
        temp_df['Инициалы_фамилия_род_падеж_пробел'] = temp_df['Родительный_падеж'].apply(lambda x: create_initials(x, 'ИФ', 'пробел'))

        # Создаем колонки для склонения фамилий с иницалами дательный падеж
        temp_df['Фамилия_инициалы_дат_падеж'] = temp_df['Дательный_падеж'].apply(lambda x: create_initials(x,'ФИ','без пробела'))
        temp_df['Фамилия_инициалы_дат_падеж_пробел'] = temp_df['Дательный_падеж'].apply(lambda x: create_initials(x, 'ФИ', 'пробел'))
        temp_df['Инициалы_фамилия_дат_падеж'] = temp_df['Дательный_падеж'].apply(lambda x: create_initials(x,'ИФ','без пробела'))
        temp_df['Инициалы_фамилия_дат_падеж_пробел'] = temp_df['Дательный_падеж'].apply(lambda x: create_initials(x, 'ИФ', 'пробел'))

        # Создаем колонки для склонения фамилий с иницалами винительный падеж
        temp_df['Фамилия_инициалы_вин_падеж'] = temp_df['Винительный_падеж'].apply(lambda x: create_initials(x,'ФИ','без пробела'))
        temp_df['Фамилия_инициалы_вин_падеж_пробел'] = temp_df['Винительный_падеж'].apply(lambda x: create_initials(x, 'ФИ', 'пробел'))
        temp_df['Инициалы_фамилия_вин_падеж'] = temp_df['Винительный_падеж'].apply(lambda x: create_initials(x,'ИФ','без пробела'))
        temp_df['Инициалы_фамилия_вин_падеж_пробел'] = temp_df['Винительный_падеж'].apply(lambda x: create_initials(x, 'ИФ', 'пробел'))

        # Создаем колонки для склонения фамилий с иницалами творительный падеж
        temp_df['Фамилия_инициалы_твор_падеж'] = temp_df['Творительный_падеж'].apply(lambda x: create_initials(x,'ФИ','без пробела'))
        temp_df['Фамилия_инициалы_твор_падеж_пробел'] = temp_df['Творительный_падеж'].apply(lambda x: create_initials(x, 'ФИ', 'пробел'))
        temp_df['Инициалы_фамилия_твор_падеж'] = temp_df['Творительный_падеж'].apply(lambda x: create_initials(x,'ИФ','без пробела'))
        temp_df['Инициалы_фамилия_твор_падеж_пробел'] = temp_df['Творительный_падеж'].apply(lambda x: create_initials(x, 'ИФ', 'пробел'))
        # Создаем колонки для склонения фамилий с иницалами предложный падеж
        temp_df['Фамилия_инициалы_пред_падеж'] = temp_df['Предложный_падеж'].apply(lambda x: create_initials(x,'ФИ','без пробела'))
        temp_df['Фамилия_инициалы_пред_падеж_пробел'] = temp_df['Предложный_падеж'].apply(lambda x: create_initials(x, 'ФИ', 'пробел'))
        temp_df['Инициалы_фамилия_пред_падеж'] = temp_df['Предложный_падеж'].apply(lambda x: create_initials(x,'ИФ','без пробела'))
        temp_df['Инициалы_фамилия_пред_падеж_пробел'] = temp_df['Предложный_падеж'].apply(lambda x: create_initials(x, 'ИФ', 'пробел'))

        # Вставляем получившиеся колонки после базовой колонки с фио
        df.insert(index_fio_column + 1, 'Родительный_падеж', temp_df['Родительный_падеж'])
        df.insert(index_fio_column + 2, 'Дательный_падеж', temp_df['Дательный_падеж'])
        df.insert(index_fio_column + 3, 'Винительный_падеж', temp_df['Винительный_падеж'])
        df.insert(index_fio_column + 4, 'Творительный_падеж', temp_df['Творительный_падеж'])
        df.insert(index_fio_column + 5, 'Предложный_падеж', temp_df['Предложный_падеж'])
        df.insert(index_fio_column + 6, 'Фамилия_инициалы', temp_df['Фамилия_инициалы'])
        df.insert(index_fio_column + 7, 'Инициалы_фамилия', temp_df['Инициалы_фамилия'])
        df.insert(index_fio_column + 8, 'Фамилия_инициалы_пробел', temp_df['Фамилия_инициалы_пробел'])
        df.insert(index_fio_column + 9, 'Инициалы_фамилия_пробел', temp_df['Инициалы_фамилия_пробел'])
        # Добавляем колонки с склонениями инициалов родительный падеж
        df.insert(index_fio_column + 10, 'Фамилия_инициалы_род_падеж', temp_df['Фамилия_инициалы_род_падеж'])
        df.insert(index_fio_column + 11, 'Фамилия_инициалы_род_падеж_пробел', temp_df['Фамилия_инициалы_род_падеж_пробел'])
        df.insert(index_fio_column + 12, 'Инициалы_фамилия_род_падеж', temp_df['Инициалы_фамилия_род_падеж'])
        df.insert(index_fio_column + 13, 'Инициалы_фамилия_род_падеж_пробел', temp_df['Инициалы_фамилия_род_падеж_пробел'])
        # Добавляем колонки с склонениями инициалов дательный падеж
        df.insert(index_fio_column + 14, 'Фамилия_инициалы_дат_падеж', temp_df['Фамилия_инициалы_дат_падеж'])
        df.insert(index_fio_column + 15, 'Фамилия_инициалы_дат_падеж_пробел', temp_df['Фамилия_инициалы_дат_падеж_пробел'])
        df.insert(index_fio_column + 16, 'Инициалы_фамилия_дат_падеж', temp_df['Инициалы_фамилия_дат_падеж'])
        df.insert(index_fio_column + 17, 'Инициалы_фамилия_дат_падеж_пробел', temp_df['Инициалы_фамилия_дат_падеж_пробел'])
        # Добавляем колонки с склонениями инициалов винительный падеж
        df.insert(index_fio_column + 18, 'Фамилия_инициалы_вин_падеж', temp_df['Фамилия_инициалы_вин_падеж'])
        df.insert(index_fio_column + 19, 'Фамилия_инициалы_вин_падеж_пробел', temp_df['Фамилия_инициалы_вин_падеж_пробел'])
        df.insert(index_fio_column + 20, 'Инициалы_фамилия_вин_падеж', temp_df['Инициалы_фамилия_вин_падеж'])
        df.insert(index_fio_column + 21, 'Инициалы_фамилия_вин_падеж_пробел', temp_df['Инициалы_фамилия_вин_падеж_пробел'])
        # Добавляем колонки с склонениями инициалов творительный падеж
        df.insert(index_fio_column + 22, 'Фамилия_инициалы_твор_падеж', temp_df['Фамилия_инициалы_твор_падеж'])
        df.insert(index_fio_column + 23, 'Фамилия_инициалы_твор_падеж_пробел', temp_df['Фамилия_инициалы_твор_падеж_пробел'])
        df.insert(index_fio_column + 24, 'Инициалы_фамилия_твор_падеж', temp_df['Инициалы_фамилия_твор_падеж'])
        df.insert(index_fio_column + 25, 'Инициалы_фамилия_твор_падеж_пробел', temp_df['Инициалы_фамилия_твор_падеж_пробел'])
        # Добавляем колонки с склонениями инициалов предложный падеж
        df.insert(index_fio_column + 26, 'Фамилия_инициалы_пред_падеж', temp_df['Фамилия_инициалы_пред_падеж'])
        df.insert(index_fio_column + 27, 'Фамилия_инициалы_пред_падеж_пробел', temp_df['Фамилия_инициалы_пред_падеж_пробел'])
        df.insert(index_fio_column + 28, 'Инициалы_фамилия_пред_падеж', temp_df['Инициалы_фамилия_пред_падеж'])
        df.insert(index_fio_column + 29, 'Инициалы_фамилия_пред_падеж_пробел', temp_df['Инициалы_фамилия_пред_падеж_пробел'])



        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        df.to_excel(f'{path_to_end_folder_decl_case}/ФИО по падежам от {current_time}.xlsx', index=False)
    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'В таблице не найдена указанная колонка {e.args}')
    except ValueError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'В таблице нет колонки с таким названием!\nПроверьте написание названия колонки')
        logging.exception('AN ERROR HAS OCCURRED')
    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    # except:
    #     logging.exception('AN ERROR HAS OCCURRED')
    #     messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.30',
    #                          'Возникла ошибка!!! Подробности ошибки в файле error.log')
    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.30', 'Данные успешно обработаны')

"""
Функции для создания контекстного меню(Копировать,вставить,вырезать)
"""
def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")

def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))

def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)



if __name__ == '__main__':
    window = Tk()
    window.title('Веста Обработка таблиц и создание документов ver 1.30')
    window.geometry('774x860+700+100')
    window.resizable(False, False)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку создания документов по шаблону
    tab_create_doc = ttk.Frame(tab_control)
    tab_control.add(tab_create_doc, text='Создание\nдокументов')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Создание документов
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_create_doc,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nГенерация документов по шаблону'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки'
                           '\nДанные обрабатываются только с первого листа файла Excel!!!')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img = PhotoImage(file=path_to_img)
    Label(tab_create_doc,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_doc = LabelFrame(tab_create_doc, text='Подготовка')
    frame_data_for_doc.grid(column=0, row=2, padx=10)

    # Создаем кнопку Выбрать шаблон
    btn_template_doc = Button(frame_data_for_doc, text='1) Выберите шаблон документа', font=('Arial Bold', 15),
                              command=select_file_template_doc
                              )
    btn_template_doc.grid(column=0, row=3, padx=10, pady=10)
    #
    # Создаем кнопку Выбрать файл с данными
    btn_data_doc = Button(frame_data_for_doc, text='2) Выберите файл с данными', font=('Arial Bold', 15),
                          command=select_file_data_doc
                          )
    btn_data_doc.grid(column=0, row=4, padx=10, pady=10)
    #
    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    # Определяем текстовую переменную
    entry_name_column_data = StringVar()
    # Описание поля
    label_name_column_data = Label(frame_data_for_doc,
                                   text='3) Введите название колонки в таблице\n по которой будут создаваться имена файлов')
    label_name_column_data.grid(column=0, row=5, padx=10, pady=5)
    # поле ввода
    data_column_entry = Entry(frame_data_for_doc, textvariable=entry_name_column_data, width=30)
    data_column_entry.grid(column=0, row=6, padx=5, pady=5, ipadx=30, ipady=4)

    # Поле для ввода названия генериуемых документов
    # Определяем текстовую переменную
    entry_type_file = StringVar()
    # Описание поля
    label_name_column_type_file = Label(frame_data_for_doc, text='4) Введите название создаваемых документов')
    label_name_column_type_file.grid(column=0, row=7, padx=10, pady=5)
    # поле ввода
    type_file_column_entry = Entry(frame_data_for_doc, textvariable=entry_type_file, width=30)
    type_file_column_entry.grid(column=0, row=8, padx=5, pady=5, ipadx=30, ipady=4)

    btn_choose_end_folder_doc = Button(frame_data_for_doc, text='5) Выберите конечную папку', font=('Arial Bold', 15),
                                       command=select_end_folder_doc
                                       )
    btn_choose_end_folder_doc.grid(column=0, row=9, padx=10, pady=10)

    # Создаем область для того чтобы поместить туда опции
    frame_data_for_options = LabelFrame(tab_create_doc, text='Дополнительные опции')
    frame_data_for_options.grid(column=0, row=10, padx=10)

    # Создаем переменную для хранения результа переключения чекбокса
    mode_combine_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_combine_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_calculate = Checkbutton(frame_data_for_options,
                                       text='Поставьте галочку, если вам нужно чтобы все файлы были объединены в один',
                                       variable=mode_combine_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_calculate.grid(column=0, row=11, padx=10, pady=5)

    # создаем чекбокс для единичного документа

    # Создаем переменную для хранения результа переключения чекбокса
    mode_group_doc = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_group_doc.set('No')
    # Создаем чекбокс для выбора режима подсчета
    chbox_mode_group = Checkbutton(frame_data_for_options,
                                       text='Поставьте галочку, если вам нужно создать один документ\nдля конкретного значения (например для определенного ФИО)',
                                       variable=mode_group_doc,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_group.grid(column=0, row=12, padx=10, pady=5)
    # Создаем поле для ввода значения по которому будет создаваться единичный документ
    # Определяем текстовую переменную
    entry_value_column = StringVar()
    # Описание поля
    label_name_column_group = Label(frame_data_for_options, text='Введите значение из колонки\nуказанной на шаге 3 для которого нужно создать один документ,\nнапример конкретное ФИО')
    label_name_column_group.grid(column=0, row=13, padx=10, pady=5)
    # поле ввода
    type_file_group_entry = Entry(frame_data_for_options, textvariable=entry_value_column, width=30)
    type_file_group_entry.grid(column=0, row=14, padx=5, pady=5, ipadx=30, ipady=4)



    # Создаем кнопку для создания документов из таблиц с произвольной структурой
    btn_create_files_other = Button(tab_create_doc, text='6) Создать документ(ы)',
                                    font=('Arial Bold', 15),
                                    command=generate_docs_other
                                    )
    btn_create_files_other.grid(column=0, row=14, padx=10, pady=10)

    # Создаем вклдаку для обработки дат рождения

    tab_calculate_date = ttk.Frame(tab_control)
    tab_control.add(tab_calculate_date, text='Обработка\nдат рождения')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Обработка дат рождения
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_calculate_date,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nПодсчет по категориям,выделение месяца,года,подсчет текущего возраста'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки'
                           '\nДанные обрабатываются только с первого листа файла Excel!!!')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img_date = PhotoImage(file=path_to_img)
    Label(tab_calculate_date,
          image=img_date
          ).grid(column=1, row=0, padx=10, pady=25)

    # Определяем текстовую переменную которая будет хранить дату
    entry_date = StringVar()
    # Описание поля
    label_name_date_field = Label(tab_calculate_date,
                                  text='Введите  дату в формате XX.XX.XXXX\n относительно, которой нужно подсчитать текущий возраст')
    label_name_date_field.grid(column=0, row=2, padx=10, pady=10)
    # поле ввода
    date_field = Entry(tab_calculate_date, textvariable=entry_date, width=30)
    date_field.grid(column=0, row=3, padx=5, pady=5, ipadx=30, ipady=15)

    # Создаем кнопку Выбрать файл с данными
    btn_data_date = Button(tab_calculate_date, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                           command=select_file_data_date)
    btn_data_date.grid(column=0, row=4, padx=10, pady=10)

    btn_choose_end_folder_date = Button(tab_calculate_date, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                        command=select_end_folder_date
                                        )
    btn_choose_end_folder_date.grid(column=0, row=5, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_name_column = StringVar()
    # Описание поля
    label_name_column = Label(tab_calculate_date,
                              text='3) Введите название колонки с датами рождения,\nкоторые нужно обработать ')
    label_name_column.grid(column=0, row=6, padx=10, pady=10)
    # поле ввода
    column_entry = Entry(tab_calculate_date, textvariable=entry_name_column, width=30)
    column_entry.grid(column=0, row=7, padx=7, pady=5, ipadx=30, ipady=15)

    btn_calculate_date = Button(tab_calculate_date, text='4) Обработать', font=('Arial Bold', 20),
                                command=calculate_date)
    btn_calculate_date.grid(column=0, row=8, padx=10, pady=10)

    # Создаем вкладку для подсчета данных по категориям
    tab_groupby_data = ttk.Frame(tab_control)
    tab_control.add(tab_groupby_data, text='Подсчет\nданных')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Подсчет данных  по категориям
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_groupby_data,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nПодсчет данных'
                           '\nДанные обрабатываются только с первого листа файла Excel!!!')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img_groupby = PhotoImage(file=path_to_img)
    Label(tab_groupby_data,
          image=img_groupby
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_groupby = LabelFrame(tab_groupby_data, text='Подготовка')
    frame_data_for_groupby.grid(column=0, row=2, padx=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_groupby = Button(frame_data_for_groupby, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                              command=select_file_data_groupby
                              )
    btn_data_groupby.grid(column=0, row=3, padx=10, pady=10)

    btn_choose_end_folder_groupby = Button(frame_data_for_groupby, text='2) Выберите конечную папку',
                                           font=('Arial Bold', 20),
                                           command=select_end_folder_groupby
                                           )
    btn_choose_end_folder_groupby.grid(column=0, row=4, padx=10, pady=10)

       # Создаем кнопки подсчета

    btn_groupby_category = Button(tab_groupby_data, text='Подсчитать количество по категориям\nдля всех колонок',
                                  font=('Arial Bold', 20),
                                  command=groupby_category)
    btn_groupby_category.grid(column=0, row=5, padx=10, pady=10)

    btn_groupby_stat = Button(tab_groupby_data, text='Подсчитать базовую статистику\nдля всех колонок',
                              font=('Arial Bold', 20),
                              command=groupby_stat)
    btn_groupby_stat.grid(column=0, row=6, padx=10, pady=10)

    # Создаем вкладку для сравнения 2 столбцов

    tab_comparison = ttk.Frame(tab_control)
    tab_control.add(tab_comparison, text='Слияние\n2 таблиц')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Создание документов
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_comparison,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_com = resource_path('logo.png')
    img_comparison = PhotoImage(file=path_com)
    Label(tab_comparison,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_comparison = LabelFrame(tab_comparison, text='Подготовка')
    frame_data_for_comparison.grid(column=0, row=2, padx=10)

    # Создаем кнопку выбрать файл с параметрами
    btn_columns_params = Button(frame_data_for_comparison, text='1) Выберите файл с параметрами слияния',
                                font=('Arial Bold', 10),
                                command=select_file_params_comparsion)
    btn_columns_params.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_comparison = Button(frame_data_for_comparison, text='2) Выберите первый файл с данными',
                                       font=('Arial Bold', 10),
                                       command=select_first_comparison
                                       )
    btn_data_first_comparison.grid(column=0, row=4, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_first_sheet_name = StringVar()
    # Описание поля
    label_first_sheet_name = Label(frame_data_for_comparison,
                                   text='3) Введите название листа в первом файле')
    label_first_sheet_name.grid(column=0, row=5, padx=10, pady=10)
    # поле ввода имени листа
    first_sheet_name_entry = Entry(frame_data_for_comparison, textvariable=entry_first_sheet_name, width=30)
    first_sheet_name_entry.grid(column=0, row=6, padx=5, pady=5, ipadx=15, ipady=10)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_comparison = Button(frame_data_for_comparison, text='4) Выберите второй файл с данными',
                                        font=('Arial Bold', 10),
                                        command=select_second_comparison
                                        )
    btn_data_second_comparison.grid(column=0, row=7, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_sheet_name = StringVar()
    # Описание поля
    label_second_sheet_name = Label(frame_data_for_comparison,
                                    text='5) Введите название листа во втором файле')
    label_second_sheet_name.grid(column=0, row=8, padx=10, pady=10)
    # поле ввода
    second__sheet_name_entry = Entry(frame_data_for_comparison, textvariable=entry_second_sheet_name, width=30)
    second__sheet_name_entry.grid(column=0, row=9, padx=5, pady=5, ipadx=15, ipady=10)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_comparison = Button(frame_data_for_comparison, text='6) Выберите конечную папку',
                                       font=('Arial Bold', 10),
                                       command=select_end_folder_comparison
                                       )
    btn_select_end_comparison.grid(column=0, row=10, padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_comparison = Button(tab_comparison, text='7) Произвести слияние\nтаблиц', font=('Arial Bold', 20),
                                    command=processing_comparison
                                    )
    btn_data_do_comparison.grid(column=0, row=11, padx=10, pady=10)

    # Создаем вкладку для обработки таблиц excel  с одинаковой структурой
    tab_calculate_data = ttk.Frame(tab_control)
    tab_control.add(tab_calculate_data, text='Извлечение\nданных')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вклдаку Обработки данных
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_calculate_data,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nИзвлечение данных из файлов Excel\nс одинаковой структурой')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_calculate = resource_path('logo.png')
    img_calculate = PhotoImage(file=path_to_img)
    Label(tab_calculate_data,
          image=img_calculate
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с параметрами
    btn_select_file_params = Button(tab_calculate_data, text='1) Выбрать файл с параметрами', font=('Arial Bold', 20),
                                    command=select_file_params_calculate_data
                                    )
    btn_select_file_params.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_select_files_data = Button(tab_calculate_data, text='2) Выбрать файлы с данными', font=('Arial Bold', 20),
                                   command=select_files_data_calculate_data
                                   )
    btn_select_files_data.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_calculate_data, text='3) Выбрать конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder_calculate_data
                                   )
    btn_choose_end_folder.grid(column=0, row=4, padx=10, pady=10)

    # Создаем переменную для хранения результа переключения чекбокса
    mode_text_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_text_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_calculate = Checkbutton(tab_calculate_data,
                                       text='Поставьте галочку, если вам нужно подсчитать текстовые данные ',
                                       variable=mode_text_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_calculate.grid(column=0, row=5, padx=10, pady=10)

    # Создаем кнопку для запуска подсчета файлов

    btn_calculate = Button(tab_calculate_data, text='4) Подсчитать', font=('Arial Bold', 20),
                           command=calculate_data
                           )
    btn_calculate.grid(column=0, row=6, padx=10, pady=10)

    """
    Создание вкладки для объединения таблиц в одну большую
    """
    # Создаем вкладку для подсчета данных по категориям
    tab_merger_tables = ttk.Frame(tab_control)
    tab_control.add(tab_merger_tables, text='Слияние\nфайлов')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Подсчет данных  по категориям
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_merger_tables,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nСлияние файлов Excel с одинаковой структурой'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки'
                      )
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img_merger = PhotoImage(file=path_to_img)
    Label(tab_merger_tables,
          image=img_merger
          ).grid(column=1, row=0, padx=10, pady=25)

    # Переключатель:вариант слияния файлов
    # Создаем переключатель
    group_rb_type_harvest = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_harvest = LabelFrame(tab_merger_tables, text='1) Выберите вариант слияния')
    frame_rb_type_harvest.grid(column=0, row=1, padx=10)
    #
    Radiobutton(frame_rb_type_harvest, text='А) Простое слияние по названию листов', variable=group_rb_type_harvest, value=0).pack()
    Radiobutton(frame_rb_type_harvest, text='Б) Слияние по порядку листов', variable=group_rb_type_harvest, value=1).pack()
    Radiobutton(frame_rb_type_harvest, text='В) Сложное слияние по названию листов', variable=group_rb_type_harvest, value=2).pack()

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_merger = LabelFrame(tab_merger_tables, text='Подготовка')
    frame_data_for_merger.grid(column=0, row=2, padx=10)

    #Создаем кнопку Выбрать папку с данными

    btn_data_merger = Button(frame_data_for_merger, text='2) Выберите папку с данными', font=('Arial Bold', 14),
                             command=select_folder_data_merger
                             )
    btn_data_merger.grid(column=0, row=3, padx=5, pady=5)

    # Создаем кнопку Выбрать эталонный файл

    btn_example_merger = Button(frame_data_for_merger, text='3) Выберите эталонный файл', font=('Arial Bold', 14),
                                command=select_standard_file_merger)
    btn_example_merger.grid(column=0, row=4, padx=5, pady=5)

    btn_choose_end_folder_merger = Button(frame_data_for_merger, text='4) Выберите конечную папку',
                                          font=('Arial Bold', 14),
                                          command=select_end_folder_merger
                                          )
    btn_choose_end_folder_merger.grid(column=0, row=5, padx=5, pady=5)


    # Определяем переменную в которой будем хранить количество пропускаемых строк
    merger_entry_skip_rows = StringVar()
    # Описание поля
    merger_label_skip_rows = Label(frame_data_for_merger,
                                   text='5) Введите количество строк\nв листах,чтобы пропустить\nзаголовок\n'
                                        'ТОЛЬКО для вариантов слияния А и Б ')
    merger_label_skip_rows.grid(column=0, row=8, padx=10, pady=10)
    # поле ввода
    merger_number_skip_rows = Entry(frame_data_for_merger, textvariable=merger_entry_skip_rows, width=5)
    merger_number_skip_rows.grid(column=0, row=9, padx=5, pady=5, ipadx=10, ipady=7)

    # Создаем кнопку выбора файла с параметрами
    btn_params_merger = Button(frame_data_for_merger, text='Выберите файл с параметрами слияния\n'
                                                           'ТОЛЬКО для варианта В', font=('Arial Bold', 14),
                                command=select_params_file_merger)
    btn_params_merger.grid(column=0, row=10, padx=5, pady=5)
     # Создаем кнопку слияния

    btn_merger_process = Button(tab_merger_tables, text='6) Произвести слияние \nфайлов',
                                font=('Arial Bold', 20),
                                command=merge_tables)
    btn_merger_process.grid(column=0, row=11, padx=10, pady=10)


    """
    Создание вкладки для склонения ФИО по падежам
    """
    # Создаем вкладку для подсчета данных по категориям
    tab_decl_by_cases = ttk.Frame(tab_control)
    tab_control.add(tab_decl_by_cases, text='Склонение ФИО\nпо падежам')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Подсчет данных  по категориям
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_decl_by_cases,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nСклонение ФИО по падежам и создание инициалов'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки'
                      )
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img_decl_by_cases = PhotoImage(file=path_to_img)
    Label(tab_decl_by_cases,
          image=img_decl_by_cases
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_decl_case = LabelFrame(tab_decl_by_cases, text='Подготовка')
    frame_data_for_decl_case.grid(column=0, row=2, padx=10)

   # выбрать файл с данными
    btn_data_decl_case = Button(frame_data_for_decl_case, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                                command=select_data_decl_case)
    btn_data_decl_case.grid(column=0, row=3, padx=10, pady=10)

    # Ввести название колонки с ФИО
    # # Определяем переменную
    decl_case_fio_col = StringVar()
    # Описание поля ввода
    decl_case_label_fio = Label(frame_data_for_decl_case,
                                    text='2) Введите название колонки\n с ФИО в им.падеже')
    decl_case_label_fio.grid(column=0, row=4, padx=10, pady=10)
    # поле ввода
    decl_case_entry_fio = Entry(frame_data_for_decl_case, textvariable=decl_case_fio_col, width=25)
    decl_case_entry_fio.grid(column=0, row=5, padx=5, pady=5, ipadx=10, ipady=7)
    #
    btn_choose_end_folder_decl_case = Button(frame_data_for_decl_case, text='3) Выберите конечную папку',
                                          font=('Arial Bold', 20),
                                          command=select_end_folder_decl_case
                                          )
    btn_choose_end_folder_decl_case.grid(column=0, row=6, padx=10, pady=10)

    # Создаем кнопку склонения по падежам

    btn_decl_case_process = Button(tab_decl_by_cases, text='4) Произвести склонение \nпо падежам',
                                font=('Arial Bold', 20),
                                command=process_decl_case)
    btn_decl_case_process.grid(column=0, row=7, padx=10, pady=10)
window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
window.bind_class("Entry", "<Control-a>", callback_select_all)
window.mainloop()
