"""
Скрипт для отработки генерации вордовских файлов в один файл
"""
from docxtpl import DocxTemplate
import pandas as pd
import openpyxl
import time
from docxcompose.composer import Composer
from docx import Document
import tempfile
import os


def combine_all_docx(filename_master, files_lst):
    """
    Функция для объединения файлов взято отсюда
    https://stackoverflow.com/questions/24872527/combine-word-document-using-python-docx
    :param filename_master: базовый файл
    :param files_list: список с созданными файлами
    :return: итоговый файл
    """
    #Получаем текущее время
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
    composer.save(f"{path_to_end_folder_date}/Объединеный файл от {current_time}.docx")


df = pd.read_excel('Данные 200.xlsx')
name_file_template_doc = 'Шаблон согласия для объединенного файла.docx'
path_to_end_folder_date = 'data/'
data = df.to_dict('records')

# Список с созданными файлами
files_lst = []



#Создаем временную папку
with tempfile.TemporaryDirectory() as tmpdirname:
    print('created temporary directory', tmpdirname)
    # Создаем и сохраняем во временную папку созданные документы Word
    for row in data:
        doc = DocxTemplate(name_file_template_doc)
        context = row
        doc.render(context)
        # Сохраняем файл
        doc.save(f'{tmpdirname}/{row["ФИО"]}.docx')
        # Добавляем путь к файлу в список
        files_lst.append(f'{tmpdirname}/{row["ФИО"]}.docx')

    # Получаем базовый файл
    main_doc = files_lst.pop(0)
    # Запускаем функцию
    combine_all_docx(main_doc,files_lst)



