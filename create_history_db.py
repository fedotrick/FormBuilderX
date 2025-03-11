import sqlite3
import re
from datetime import datetime

def create_history_database():
    # Добавить обработку ошибок:
    try:
        conn = sqlite3.connect('история_форм.db')
        cursor = conn.cursor()

        # Создаем таблицу для хранения истории форм
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS "Маршрутные_Карты" (
                "ИД" INTEGER PRIMARY KEY AUTOINCREMENT,
                "Номер_Кластера" TEXT NOT NULL,
                "Номер_Отливки" TEXT,
                "Наименование_Отливки" TEXT,
                "Отливка" TEXT,
                "Дата_Склейки" TEXT,
                "Исполнитель_Склейки" TEXT,
                "Количество_Склейки" TEXT,
                "Примечание_Склейки" TEXT,
                "Дата_Контроля" TEXT,
                "Время_Контроля" TEXT,
                "Исполнитель_Контроля" TEXT,
                "Количество_Контроля" TEXT,
                "Примечание_Контроля" TEXT,
                "Дата_Создания" TEXT
            )
        ''')

        # Создаем уникальный индекс для номера кластера
        cursor.execute('''
            CREATE UNIQUE INDEX IF NOT EXISTS "idx_номер_кластера" 
            ON "Маршрутные_Карты" ("Номер_Кластера")
        ''')

        conn.commit()
    except Exception as e:
        print(f"Ошибка при создании базы данных: {e}")
        raise
    finally:
        conn.close()

def validate_cluster_number(number):
    """Проверяет формат номера кластера"""
    pattern = r'^К\d{2}/\d{2}-\d{3}$'
    if not re.match(pattern, number):
        raise ValueError(
            "Неверный формат номера кластера. "
            "Должен быть в формате: КГГ/ММ-NNN (например: К25/03-001)"
        )
    
    # Проверяем корректность месяца
    month = int(number[4:6])
    if month < 1 or month > 12:
        raise ValueError("Неверный номер месяца в номере кластера (должен быть от 01 до 12)")

def save_form_data(data):
    conn = sqlite3.connect('история_форм.db')
    cursor = conn.cursor()
    
    # Определяем тип отливки и формируем строку отливки
    ref_conn = sqlite3.connect('справочник.db')
    ref_cursor = ref_conn.cursor()
    
    cast_string = ""
    
    # Проверяем в ЛГМ
    ref_cursor.execute('SELECT "Наименование" FROM "ЛГМ_Отливки" WHERE "Номер" = ?', (data['cast_number'],))
    result = ref_cursor.fetchone()
    if result:
        cast_string = f"{result[0]} {data['cast_number']}"
    else:
        # Проверяем в ЛПД
        ref_cursor.execute('SELECT "Наименование" FROM "ЛПД_Отливки" WHERE "Номер" = ?', (data['cast_number'],))
        result = ref_cursor.fetchone()
        if result:
            cast_string = f"{data['cast_number']} {result[0]}"
        else:
            # Если не найдено в ЛГМ и ЛПД, значит это из Прочие
            cast_string = data['cast_name']
    
    ref_conn.close()
    
    # Добавляем новую запись
    cursor.execute('''
        INSERT INTO "Маршрутные_Карты" (
            "Номер_Кластера",
            "Номер_Отливки",
            "Наименование_Отливки",
            "Отливка",
            "Дата_Склейки",
            "Исполнитель_Склейки",
            "Количество_Склейки",
            "Примечание_Склейки",
            "Дата_Контроля",
            "Время_Контроля",
            "Исполнитель_Контроля",
            "Количество_Контроля",
            "Примечание_Контроля",
            "Дата_Создания"
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data['cluster_number'],
        data['cast_number'],
        data['cast_name'],
        cast_string,
        data['gluing_date'],
        data['gluing_executor'],
        data['gluing_quantity'],
        data['gluing_notes'],
        data['control_date'],
        data['control_time'],
        data['control_executor'],
        data['control_quantity'],
        data['control_notes'],
        datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    ))
    
    conn.commit()
    conn.close()

def get_next_cluster_number(date_str):
    """
    Генерирует следующий номер кластера на основе даты склейки
    date_str: строка в формате dd.MM.yyyy
    """
    # Добавить проверку входных данных:
    if not date_str or len(date_str.split('.')) != 3:
        raise ValueError("Неверный формат даты. Ожидается dd.MM.yyyy")
    
    # Разбираем дату
    day, month, year = map(int, date_str.split('.'))
    year_short = str(year)[-2:]  # Берем последние 2 цифры года
    
    conn = sqlite3.connect('история_форм.db')
    cursor = conn.cursor()
    
    try:
        # Ищем последний номер для этого месяца и года
        cursor.execute('''
            SELECT "Номер_Кластера"
            FROM "Маршрутные_Карты"
            WHERE "Номер_Кластера" LIKE ?
            ORDER BY "Номер_Кластера" DESC
            LIMIT 1
        ''', (f'К{year_short}/{month:02d}-%',))
        
        result = cursor.fetchone()
        
        if result:
            # Если есть номера для этого месяца, увеличиваем последний на 1
            last_number = int(result[0][-3:])
            next_number = last_number + 1
            if next_number > 999:
                raise ValueError("Превышено максимальное количество кластеров для этого месяца (999)")
        else:
            # Если нет номеров для этого месяца, начинаем с 001
            next_number = 1
        
        # Формируем новый номер кластера
        new_cluster_number = f'К{year_short}/{month:02d}-{next_number:03d}'
        return new_cluster_number
        
    finally:
        conn.close()

if __name__ == "__main__":
    create_history_database()
    print("База данных истории форм успешно создана.") 