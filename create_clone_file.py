import openpyxl
from openpyxl import load_workbook

# file_name = 'data/union/Ингушетия Приложение_№_1 (2).xlsx'
file_name = 'data/test1/Свод  за 1 квартал 2022 года.xlsx'
base_wb = load_workbook(file_name)
path_dir = 'data/clone_file'
# Создаем 100 файлов
for i in range(1,101):
    # base_wb.save(f'{path_dir}/Тестовый файл №{i}.xlsx')
    base_wb.save(f'{path_dir}/Сложный вариант таблицы №{i}.xlsx')


