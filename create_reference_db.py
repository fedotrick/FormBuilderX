import sqlite3

def create_reference_database():
    # Добавить обработку ошибок:
    try:
        # Подключаемся к базе данных (создаст файл, если его нет)
        conn = sqlite3.connect('справочник.db')
        cursor = conn.cursor()

        # Данные для заполнения таблиц
        data = {
            'Контролеры_Сборки': [('Елхова',), ('Малых',), ('Романцева',), ('Шестункина',)],
            'Сборщики': [('Буцик',), ('Минакова',), ('Ротарь',), ('Чернова',), ('Чупахина',)],
            'ЛГМ_Отливки': [
                ('ЛСКМ.03.01.102-Л1', 'Держатель ригеля'),
                ('ЛСКМ.03.51.102-Л', 'Держатель ригеля OPTIMA'),
                ('ЛСКМ.04.00.004-Л', 'Держатель диагонали'),
                ('ЛСКМ.04.00.002-Л', 'Держатель диагонали'),
                ('ОКР-2410.02.000', 'Вороток'),
                ('ЛСКМ.00.03.202-Л', 'Вороток'),
                ('ЛСКМ.98.53.101-Л', 'Соединитель угловой'),
                ('ЛСКМ.98.14.001-Л', 'Адаптер'),
                ('ЛСКМ.00.06.001-Л U1', 'Накладка для резьбы домкрата'),
                ('АМ.3509030-130 3Г', 'Блок-картер 2 цилиндра'),
                ('5У.01.001-V15m1', 'ПНГ'),
                ('Н2А.03М.01.01-Л', 'Полухомут верхний'),
                ('Н2А.05М.02.01-Л', 'Полухомут нижний')
            ],
            'ЛПД_Отливки': [
                ('8450090064-Л', 'Корпус шкива опорного'),
                ('2123-1011371-Л', 'Фиксатор шестерни привода масляного насоса'),
                ('21214-1011371-Л', 'Фиксатор шестерни привода масляного насоса'),
                ('11189-1041034-Л', 'Кронштейн генератора'),
                ('11189-1041034-10-Л', 'Кронштейн генератора'),
                ('21082-3701652-Л', 'Кронштейн генератора нижний'),
                ('21214-3701652-Л', 'Кронштейн генератора нижний'),
                ('8450036497-Л', 'Кронштейн крепления передней защитной крышки'),
                ('850120774-Л', 'Кронштейн вспомогательных агрегатов')
            ],
            'Прочие_Отливки': [
                ('Чугун',), ('Колесо РИТМ',), ('Скоба',), ('Лопасть',),
                ('Корпус',), ('Изложница',), ('Защита',), ('Кольцо вкладыш',)
            ],
            'Типы_Эксперимента': [('Бумага',), ('Волокно',)],
            'Старшие_Смены_Плавки': [
                ('Белков',), ('Валиулин',), ('Ермаков',), ('Карасев',)
            ],
            'Участники_Плавки': [
                ('Беляев',), ('Волков',), ('Исмаилов',), ('Кокшин',),
                ('Левин',), ('Политов',), ('Рабинович',), ('Семенов',), ('Терентьев',)
            ],
            'Специалисты_Термообработка': [('Аюбов',), ('Эгамов',)],
            'Специалисты_Дробемет': [('Аюбов',), ('Эгамов',)],
            'Специалисты_Резка': [
                ('Абдухакимов',), ('Ахмаджонов',), ('Исмаилов',), ('Косимов',),
                ('Косимов-2',), ('Машрапов',), ('Отаназаров',), ('Самиев',),
                ('Туичиев',), ('Эргашев',)
            ],
            'Специалисты_Зачистка': [
                ('Абдуллаев',), ('Бурхонов',), ('Матесаев',), ('Отаназаров',), ('Самиев',)
            ],
            'Контролеры': [('Елхова',), ('Лабуткина',), ('Рябова',), ('Улитина',)]
        }

        # Создаем таблицы
        tables = {
            'Контролеры_Сборки': ['ФИО'],
            'Сборщики': ['ФИО'],
            'ЛГМ_Отливки': ['Номер', 'Наименование'],
            'ЛПД_Отливки': ['Номер', 'Наименование'],
            'Прочие_Отливки': ['Наименование'],
            'Типы_Эксперимента': ['Наименование'],
            'Старшие_Смены_Плавки': ['ФИО'],
            'Участники_Плавки': ['ФИО'],
            'Специалисты_Термообработка': ['ФИО'],
            'Специалисты_Дробемет': ['ФИО'],
            'Специалисты_Резка': ['ФИО'],
            'Специалисты_Зачистка': ['ФИО'],
            'Контролеры': ['ФИО']
        }

        # Создаем таблицы, если их нет
        for table_name, columns in tables.items():
            columns_sql = ', '.join([f'"{col}" TEXT' for col in columns])
            cursor.execute(f'CREATE TABLE IF NOT EXISTS "{table_name}" ({columns_sql})')
            
            # Проверяем, есть ли данные в таблице
            cursor.execute(f'SELECT COUNT(*) FROM "{table_name}"')
            count = cursor.fetchone()[0]
            
            # Заполняем данными только если таблица пустая
            if count == 0:
                rows = data[table_name]
                placeholders = ','.join(['?' for _ in range(len(rows[0]))])
                cursor.executemany(
                    f'INSERT INTO "{table_name}" VALUES ({placeholders})',
                    rows
                )
                print(f'Таблица "{table_name}" создана и заполнена.')
            else:
                print(f'Таблица "{table_name}" уже содержит данные, пропускаем заполнение.')

        # Сохраняем изменения и закрываем соединение
        conn.commit()
    except Exception as e:
        print(f"Ошибка при создании базы данных: {e}")
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    create_reference_database()
    print("База данных успешно создана и заполнена.") 