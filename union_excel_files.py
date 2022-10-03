import pandas as pd
import openpyxl
import os

path_to_files ='data/temp2/'

skip_rows = 0
name_sheet = 'Раздел 1.5'
# standard_df = pd.read_excel('data/temp2/Логины и пароли ЭО-22.xlsx')
standard_df = pd.read_excel('data/temp2/БФИИТ.xlsx',skiprows=skip_rows,sheet_name=name_sheet)

#  Создаем базовый датафрейм
base_df = pd.DataFrame(columns=standard_df.columns)
base_df.insert(0,'Имя файла',None)
for dirpath, dirnames, filenames in os.walk(path_to_files):
    for filename in filenames:
        if filename.endswith('.xlsx'):
            # Получаем название файла без расширения
            name_file = filename.split('.xlsx')[0]
            temp_df= pd.read_excel(f'{dirpath}{filename}',skiprows=skip_rows,sheet_name=name_sheet)
            # Проверяем соответствие колонок
            if list(standard_df.columns) == list(temp_df.columns):
                temp_df.insert(0, 'Имя файла', None)
                temp_df['Имя файла'] = name_file
                base_df = pd.concat([base_df,temp_df],axis=0,ignore_index=True)

base_df.to_excel('Тест.xlsx',index=False)



