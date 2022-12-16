import os
import pandas as pd

df = pd.read_excel('data/Список регионов РФ.xlsx')
lst_reg = df['Наименование'].tolist()
path_to_end = 'data/folder'
for reg in lst_reg:
    if not os.path.isdir(f'{path_to_end}/{reg}'):
        os.mkdir(f'{path_to_end}/{reg}')