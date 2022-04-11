import pandas as pd
import os
from docxtpl import DocxTemplate
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import time
import datetime
from datetime import date
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import BarChart, Reference, PieChart, PieChart3D, Series
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.options.mode.chained_assignment = None
import sys
import locale


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


def generate_docs_other():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных ЦОПП Бурятия)
    :return:
    """
    try:
        name_column = entry_name_column_data.get()
        name_type_file = entry_type_file.get()

        # Считываем данные
        # Добавил параметр dtype =str чтобы данные не преобразовались а использовались так как в таблице
        df = pd.read_excel(name_file_data_doc, dtype=str)

        # Обрабатываем колонки с датами, чтобы они отображались корректно
        # for column in df.columns:
        #     if df[column].dtype == 'datetime64[ns]':
        #         df[column] = df[column].apply(convert_date)

        # Конвертируем датафрейм в список словарей
        data = df.to_dict('records')

        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_doc)
            context = row
            print(context)
            doc.render(context)
            # Сохраняенм файл
            doc.save(f'{path_to_end_folder_doc}/{name_type_file} {row[name_column]}.docx')

    except NameError as e:
        messagebox.showinfo('ЦОПП Бурятия', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
    else:
        messagebox.showinfo('ЦОПП Бурятия', 'Создание документов завершено!')


def set_rus_locale():
    """
    Функция чтобы можно было извлечь русские названия месяцев
    """
    locale.setlocale(
        locale.LC_ALL,
        'rus_rus' if sys.platform == 'win32' else 'ru_RU.UTF-8')


def calculate_age(born):
    """
    Функция для расчета текущего возраста взято с https://stackoverflow.com/questions/2217488/age-from-birthdate-in-python/9754466#9754466
    :param born: дата рождения
    :return: возраст
    """

    try:
        today = date.today()
        return today.year - born.year - ((today.month, today.day) < (born.month, born.day))
    except:
        print(born)
        messagebox.showerror('ЦОПП Бурятия', 'Отсутствует или некорректная дата рождения слушателя\nПроверьте файл!')
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
        messagebox.showerror('ЦОПП Бурятия', 'Проверьте правильность заполнения ячеек с датой!!!')
        quit()


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
        messagebox.showerror('ЦОПП Бурятия', f'Выберите файл с данными и папку куда будет генерироваться файл')
    except KeyError:
        messagebox.showerror('ЦОПП Бурятия', f'В таблице нет такой колонки!\nПроверьте написание названия колонки')
    else:
        messagebox.showinfo('ЦОПП Бурятия', 'Данные успешно обработаны')


def groupby_category():
    """
    Функция для подсчета выбранной колонки по категориям
    :return:
    """
    name_column = groupby_entry_name_column.get()
    try:

        print(f'Обрабатываемая колонка {name_column}')
        # Считываем файл
        df = pd.read_excel(name_file_data_groupby)
        print(f'Колонки в таблице {df.columns}')
        # Добавляем столбец для облегчения подсчета по категориям
        df['Итого'] = 1
        # Создаем шрифт которым будем выделять названия таблиц
        font_name_table = Font(name='Arial Black', size=15, italic=True)
        # Создаем файл excel
        wb = openpyxl.Workbook()
        # Переименовываем лист
        ren_sheet = wb['Sheet']
        ren_sheet.title = 'Подсчет по категориям'

        # Проводим группировку
        group_df = df.groupby([name_column]).agg({'Итого': 'sum'})
        for r in dataframe_to_rows(group_df, index=True, header=True):
            wb['Подсчет по категориям'].append(r)
        wb['Подсчет по категориям'].column_dimensions['A'].width = 30

        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Сохраняем итоговый файл
        wb.save(f'{path_to_end_folder_groupby}/Подсчет по категориям для колонки {name_column} {current_time}.xlsx')


    except NameError:
        messagebox.showerror('ЦОПП Бурятия', f'Выберите файл с данными и папку куда будет генерироваться файл')
    except KeyError:
        messagebox.showerror('ЦОПП Бурятия', f'В таблице нет такой колонки!\nПроверьте написание названия колонки')
    except TypeError:
        messagebox.showerror('ЦОПП Бурятия',
                             f'В колонке {name_column}\nПрисутствуют некорректные данные!\nДанные должны быть однотипными')
    else:
        messagebox.showinfo('ЦОПП Бурятия', 'Данные успешно обработаны')


def groupby_stat():
    """
    Функция для подсчета выбранной колонки по количественным показателям(сумма,среднее,медиана,мин,макс)
    :return:
    """
    name_column = groupby_entry_name_column.get()
    try:

        print(f'Обрабатываемая колонка {name_column}')
        # Считываем файл
        df = pd.read_excel(name_file_data_groupby)
        print(f'Колонки в таблице {df.columns}')
        # Добавляем столбец для облегчения подсчета по категориям
        df['Итого'] = 1
        # Создаем шрифт которым будем выделять названия таблиц
        font_name_table = Font(name='Arial Black', size=15, italic=True)
        # Создаем файл excel
        wb = openpyxl.Workbook()
        wb = openpyxl.Workbook()
        ren_sheet = wb['Sheet']
        ren_sheet.title = 'Подсчет статистик'

        group_df = df[name_column].describe().to_frame()

        if group_df.shape[0] == 8:
            # подсчитаем сумму
            all_sum = df[name_column].sum()
            dct_row = {name_column: all_sum}
            row = pd.DataFrame(data=dct_row, index=['Сумма'])
            # Добавим в датафрейм
            group_df = pd.concat([group_df, row], axis=0)
            # group_df = group_df.append({name_column:all_sum},ignore_index=True)
            # Обновим названия индексов
            group_df.index = ['Количество значений', 'Среднее', 'Стандартное отклонение', 'Минимальное значение',
                              '25%(Первый квартиль)', 'Медиана', '75%(Третий квартиль)', 'Максимальное значение',
                              'Сумма']



        elif group_df.shape[0] == 4:
            group_df.index = ['Количество значений', 'Количество уникальных значений', 'Самое частое значение',
                              'Количество повторений самого частого значения', ]
        else:
            messagebox.showerror('ЦОПП Бурятия', 'Возникла проблема при обработке. Проверьте значения в колонке')
        for r in dataframe_to_rows(group_df, index=True, header=True):
            wb['Подсчет статистик'].append(r)
        wb['Подсчет статистик'].column_dimensions['A'].width = 30

        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Сохраняем итоговый файл
        wb.save(f'{path_to_end_folder_groupby}/Подсчет статистик по колонке{name_column} {current_time}.xlsx')


    except NameError:
        messagebox.showerror('ЦОПП Бурятия', f'Выберите файл с данными и папку куда будет генерироваться файл')
    except KeyError:
        messagebox.showerror('ЦОПП Бурятия', f'В таблице нет такой колонки!\nПроверьте написание названия колонки')
    except TypeError:
        messagebox.showerror('ЦОПП Бурятия',
                             f'В колонке {name_column}\nПрисутствуют некорректные данные!\nДанные должны быть однотипными')
    else:
        messagebox.showinfo('ЦОПП Бурятия', 'Данные успешно обработаны')


def processing_comparison():
    """
    Функция для сравнения 2 колонок
    :return:
    """
    try:
        # Получаем значения текстовых полей
        first_column = entry_first_name_column.get()
        second_column = entry_second_name_column.get()

        # загружаем файлы
        df_frist = pd.read_excel(name_first_file_comparison, dtype=str)
        df_second = pd.read_excel(name_second_file_comparison, dtype=str)
        # Создаем переменную для типа создаваемого документа
        status_rb_type_doc = group_rb_type_doc.get()
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Сохраняем итоговый файл

        # В зависимости от значения проводим merge
        if status_rb_type_doc == 0:
            itog_df = pd.merge(df_frist, df_second, how='inner', left_on=first_column, right_on=second_column)
            # Сохраняем результат
            itog_df.to_excel(f'{path_to_end_folder_comparison}/Совпадающие значения  от {current_time}.xlsx', index=False)
        elif status_rb_type_doc == 1:
            itog_df = pd.merge(df_frist, df_second, how='left', left_on=first_column, right_on=second_column)
            # Сохраняем результат
            itog_df.to_excel(f'{path_to_end_folder_comparison}/Left Результат обработки от {current_time}.xlsx', index=False)
            #В результат попадают совпадающие по ключу данные обеих таблиц и все записи из левой таблицы, для которых не нашлось пары в правой.
        elif status_rb_type_doc == 2:
            itog_df = pd.merge(df_frist, df_second, how='right', left_on=first_column, right_on=second_column)
            #В результат объединения попадают совпадающие по ключу записи обеих таблиц и все данные из правой таблицы, для которых не нашлось пары в левой.
            # Сохраняем результат
            itog_df.to_excel(f'{path_to_end_folder_comparison}/Right Результат обработки от {current_time}.xlsx', index=False)
        elif status_rb_type_doc == 3:
            itog_df = pd.merge(df_frist, df_second, how='outer', left_on=first_column, right_on=second_column)
            # Сохраняем результат
            itog_df.to_excel(f'{path_to_end_folder_comparison}/Outer Результат обработки от {current_time}.xlsx', index=False)
            #В результат объединения попадают совпадающие по ключу записи обеих таблиц и все строки из этих двух таблиц, для которых пар не нашлось. Порядок таблиц в запросе не важен.
        elif status_rb_type_doc == 4:
            # Создаем документ
            wb = openpyxl.Workbook()
            # создаем листы
            ren_sheet = wb['Sheet']
            ren_sheet.title = 'Первая таблица'
            wb.create_sheet(title='Вторая таблица',index=1)
            wb.create_sheet(title='Общие данные',index=2)
            # Создаем датафрейм
            itog_df = pd.merge(df_frist, df_second, how='outer', left_on=first_column, right_on=second_column,indicator=True)

            # Записываем каждый датафрейм в соответсвующий лист
            left_df = itog_df[itog_df['_merge'] == 'left_only']
            left_df.drop(['_merge'],axis=1,inplace=True)
            for r in dataframe_to_rows(left_df, index=False, header=True):
                wb['Первая таблица'].append(r)

            right_df = itog_df[itog_df['_merge'] == 'right_only']
            right_df.drop(['_merge'], axis=1,inplace=True)
            for r in dataframe_to_rows(right_df, index=False, header=True):
                wb['Вторая таблица'].append(r)

            both_df = itog_df[itog_df['_merge'] == 'both']
            both_df.drop(['_merge'],axis=1, inplace=True)
            for r in dataframe_to_rows(both_df, index=False, header=True):
                wb['Общие данные'].append(r)

            # Сохраняем
            t = time.localtime()
            current_time = time.strftime('%H_%M_%S', t)
            # Сохраняем итоговый файл
            wb.save(f'{path_to_end_folder_comparison}/Уникальные данные от {current_time}.xlsx')
    except NameError:
        messagebox.showerror('ЦОПП Бурятия', f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except KeyError:
        messagebox.showerror('ЦОПП Бурятия', f'В таблице нет такой колонки!\nПроверьте написание названия колонки')
    else:
        messagebox.showinfo('ЦОПП Бурятия', 'Данные успешно обработаны')


if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('700x860')
    window.resizable(False, False)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку создания документов по шаблону
    tab_create_doc = ttk.Frame(tab_control)
    tab_control.add(tab_create_doc, text='Создание документов')
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
    btn_template_doc = Button(frame_data_for_doc, text='1) Выберите шаблон документа', font=('Arial Bold', 20),
                              command=select_file_template_doc
                              )
    btn_template_doc.grid(column=0, row=3, padx=10, pady=10)
    #
    # Создаем кнопку Выбрать файл с данными
    btn_data_doc = Button(frame_data_for_doc, text='2) Выберите файл с данными', font=('Arial Bold', 20),
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
    label_name_column_data.grid(column=0, row=5, padx=10, pady=10)
    # поле ввода
    data_column_entry = Entry(frame_data_for_doc, textvariable=entry_name_column_data, width=30)
    data_column_entry.grid(column=0, row=6, padx=5, pady=5, ipadx=30, ipady=15)

    # Поле для ввода названия генериуемых документов
    # Определяем текстовую переменную
    entry_type_file = StringVar()
    # Описание поля
    label_name_column_type_file = Label(frame_data_for_doc, text='4) Введите название создаваемых документов')
    label_name_column_type_file.grid(column=0, row=7, padx=10, pady=10)
    # поле ввода
    type_file_column_entry = Entry(frame_data_for_doc, textvariable=entry_type_file, width=30)
    type_file_column_entry.grid(column=0, row=8, padx=5, pady=5, ipadx=30, ipady=15)

    btn_choose_end_folder_doc = Button(frame_data_for_doc, text='5) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_doc
                                       )
    btn_choose_end_folder_doc.grid(column=0, row=9, padx=10, pady=10)

    # Создаем кнопку для создания документов из таблиц с произвольной структурой
    btn_create_files_other = Button(tab_create_doc, text='Создать документы',
                                    font=('Arial Bold', 20),
                                    command=generate_docs_other
                                    )
    btn_create_files_other.grid(column=0, row=10, padx=10, pady=10)

    tab_calculate_date = ttk.Frame(tab_control)
    tab_control.add(tab_calculate_date, text='Обработка дат рождения')
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

    # Создаем кнопку Выбрать файл с данными
    btn_data_date = Button(tab_calculate_date, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                           command=select_file_data_date)
    btn_data_date.grid(column=0, row=1, padx=10, pady=10)

    btn_choose_end_folder_date = Button(tab_calculate_date, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                        command=select_end_folder_date
                                        )
    btn_choose_end_folder_date.grid(column=0, row=2, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_name_column = StringVar()
    # Описание поля
    label_name_column = Label(tab_calculate_date,
                              text='3) Введите название колонки с датами рождения,\nкоторые нужно обработать ')
    label_name_column.grid(column=0, row=3, padx=10, pady=10)
    # поле ввода
    column_entry = Entry(tab_calculate_date, textvariable=entry_name_column, width=30)
    column_entry.grid(column=0, row=4, padx=5, pady=5, ipadx=30, ipady=15)

    btn_calculate_date = Button(tab_calculate_date, text='4) Обработать', font=('Arial Bold', 20),
                                command=calculate_date)
    btn_calculate_date.grid(column=0, row=5, padx=10, pady=10)

    # Создаем вкладку для подсчета данных по категориям
    tab_groupby_data = ttk.Frame(tab_control)
    tab_control.add(tab_groupby_data, text='Подсчет данных')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Подсчет данных  по категориям
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_groupby_data,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nПодсчет данных'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки'
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

    # Определяем текстовую переменную
    groupby_entry_name_column = StringVar()
    # Описание поля
    groupby_label_name_column = Label(frame_data_for_groupby,
                                      text='3) Введите название колонки, которую нужно обработать')
    groupby_label_name_column.grid(column=0, row=5, padx=10, pady=10)
    # поле ввода
    groupby_column_entry = Entry(frame_data_for_groupby, textvariable=groupby_entry_name_column, width=30)
    groupby_column_entry.grid(column=0, row=6, padx=5, pady=5, ipadx=30, ipady=15)

    # Создаем кнопки подсчета

    btn_groupby_category = Button(tab_groupby_data, text='Подсчитать количество\n по категориям',
                                  font=('Arial Bold', 20),
                                  command=groupby_category)
    btn_groupby_category.grid(column=0, row=7, padx=10, pady=10)

    btn_groupby_stat = Button(tab_groupby_data, text='Подсчитать базовую статистику\nпо колонке',
                              font=('Arial Bold', 20),
                              command=groupby_stat)
    btn_groupby_stat.grid(column=0, row=8, padx=10, pady=10)

    # Создаем вкладку для сравнения 2 столбцов

    tab_comparison = ttk.Frame(tab_control)
    tab_control.add(tab_comparison, text='Сравнение 2 колонок')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Создание документов
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_comparison,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Получение совпадающих значений из 2 колонок'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки'
                           '\nДанные обрабатываются только с первого листа файла Excel!!!')
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

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_comparison = Button(frame_data_for_comparison, text='1) Выберите первый файл с данными',
                                       font=('Arial Bold', 10),
                                       command=select_first_comparison
                                       )
    btn_data_first_comparison.grid(column=0, row=3, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_first_name_column = StringVar()
    # Описание поля
    label_first_name_column = Label(frame_data_for_comparison,
                                    text='2) Введите название колонки в первом файле,\nкоторую нужно сравнить')
    label_first_name_column.grid(column=0, row=4, padx=10, pady=10)
    # поле ввода
    column_first_entry = Entry(frame_data_for_comparison, textvariable=entry_first_name_column, width=30)
    column_first_entry.grid(column=0, row=5, padx=5, pady=5, ipadx=15, ipady=10)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_comparison = Button(frame_data_for_comparison, text='3) Выберите второй файл с данными',
                                        font=('Arial Bold', 10),
                                        command=select_second_comparison
                                        )
    btn_data_second_comparison.grid(column=0, row=6, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_name_column = StringVar()
    # Описание поля
    label_second_name_column = Label(frame_data_for_comparison,
                                     text='4) Введите название колонки во втором файле,\nкоторую нужно сравнить')
    label_second_name_column.grid(column=0, row=7, padx=10, pady=10)
    # поле ввода
    column_second_entry = Entry(frame_data_for_comparison, textvariable=entry_second_name_column, width=30)
    column_second_entry.grid(column=0, row=8, padx=5, pady=5, ipadx=15, ipady=10)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_comparison = Button(frame_data_for_comparison, text='5) Выберите конечную папку',
                                       font=('Arial Bold', 10),
                                       command=select_end_folder_comparison
                                       )
    btn_select_end_comparison.grid(column=0, row=9, padx=10, pady=10)

    # Создаем переменную хранящую тип документа, в зависимости от значения будет использоваться та или иная функция
    group_rb_type_doc = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_doc = LabelFrame(tab_comparison, text='6) Выберите тип сравнения')
    frame_rb_type_doc.grid(column=0, row=10, padx=10)
    #
    Radiobutton(frame_rb_type_doc, text='Общие данные для обеих колонок (пересечение)', variable=group_rb_type_doc,
                value=0).pack()
    Radiobutton(frame_rb_type_doc, text='Left Join', variable=group_rb_type_doc, value=1).pack()
    Radiobutton(frame_rb_type_doc, text='Right Join', variable=group_rb_type_doc, value=2).pack()
    Radiobutton(frame_rb_type_doc, text='Объединить таблицы', variable=group_rb_type_doc, value=3).pack()
    Radiobutton(frame_rb_type_doc, text='Outer Minus Join ', variable=group_rb_type_doc, value=4).pack()

    # Создаем кнопку Обработать данные
    btn_data_do_comparison = Button(tab_comparison, text='7) Обработать данные', font=('Arial Bold', 20),
                                    command=processing_comparison
                                    )
    btn_data_do_comparison.grid(column=0, row=11, padx=10, pady=10)

window.mainloop()
