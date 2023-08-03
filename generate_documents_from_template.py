





def generate_docs_other():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных Веста Обработка таблиц и создание документов ver 1.35)
    :return:
    """
    try:
        name_column = entry_name_column_data.get()
        name_type_file = entry_type_file.get()
        name_value_column = entry_value_column.get()

        # получаем состояние чекбокса создания pdf
        mode_pdf = mode_pdf_value.get()

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
                    if mode_pdf == 'Yes':
                        convert(f'{path_to_end_folder_doc}/{name_file}.docx',f'{path_to_end_folder_doc}/{name_file}.pdf',keep_active=True)
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
                        if mode_pdf == 'Yes':
                            convert(f'{path_to_end_folder_doc}/{name_file}.docx',f'{path_to_end_folder_doc}/{name_file}.pdf',keep_active=True)
                elif len(single_data) > 1:
                    for idx,row in enumerate(single_data):
                        doc = DocxTemplate(name_file_template_doc)
                        doc.render(row)
                        # Сохраняем файл

                        doc.save(f'{path_to_end_folder_doc}/{name_file}_{idx}.docx')
                        if mode_pdf == 'Yes':
                            convert(f'{path_to_end_folder_doc}/{name_file}_{idx}.docx',f'{path_to_end_folder_doc}/{name_file}_{idx}.pdf',keep_active=True)
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
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             f'В таблице не найдена указанная колонка {e.args}')
    except PermissionError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             f'Закройте все файлы Word созданные Вестой')
        logging.exception('AN ERROR HAS OCCURRED')
    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except CheckBoxException:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             f'Уберите галочку из чекбокса Поставьте галочку, если вам нужно создать один документ\nдля конкретного значения (например для определенного ФИО)'
                             )
    except NotFoundValue:
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             f'Указанное значение не найдено в выбранной колонке\nПроверьте наличие такого значения в таблице'
                             )
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')

    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов ver 1.35', 'Создание документов завершено!')