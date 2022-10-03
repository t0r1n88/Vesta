import pandas as pd

# Октлючаем предупреждение о цепном присваивании
pd.options.mode.chained_assignment = None  # default='warn'
df = pd.read_excel('data/Сгенерированный массив данных.xlsx')
# Получаем уникальные значения групп
lst_group = df['Направление'].unique()
# перебираем названия групп, фильтруем датафрейм по итерируемому названию,сохраняем в эксель файл полученные значения
for group_name in lst_group:
    # Фильтруем по названию
    temp_df = df.query('Направление==@group_name')
    # Создаем столбец Имя Фамилия
    # Удаляем лишние столбцы
    temp_df.to_excel(f'Список {group_name}.xlsx',index=False)
