"""
Скрипт для разделения списка по значениям выбранной колонки.Результаты сохраняются либо в листы одного файла либо
в отдельные файлы. Например разделить большой список по полу или по группам
"""
import pandas as pd
import os
from tkinter import messagebox
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import time
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
from jinja2 import exceptions
import logging
logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)

def split_table(file_data_split:str,name_sheet:str,number_column:int,checkbox_split:int,path_to_end_folder):
    """
    Функция для разделения таблицы по значениям в определенном листе и колонке. Разделение по файлам и листам с сохранением названий

    :param file_data_split: файл с таблицей
    :param name_sheet: имя листа с таблицей
    :param number_column:порядковый номер колонки , прибавляется 1 чтобы соответстовать экселю
    :param checkbox_split: вариант разделения
    :param path_to_end_folder: путь к итоговой папке
    :return: один файл в котором много листов либо много файлов в зависимости от режима
    """
    pass



