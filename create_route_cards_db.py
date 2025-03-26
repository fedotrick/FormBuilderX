import sqlite3

def create_route_cards_database():
    try:
        conn = sqlite3.connect('маршрутные_карты.db')
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS маршрутные_карты (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                номер_маршрутной_карты TEXT UNIQUE NOT NULL,
                номер_кластера TEXT
            )
        ''')

        conn.commit()
    except Exception as e:
        print(f"Ошибка при создании базы данных: {e}")
        raise
    finally:
        conn.close()

def update_cluster_number(route_card_number, cluster_number):
    """Обновляет номер кластера для выбранной маршрутной карты"""
    conn = sqlite3.connect('маршрутные_карты.db')
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
            UPDATE маршрутные_карты 
            SET "Номер_кластера" = ? 
            WHERE "Номер_бланка" = ?
        ''', (cluster_number, route_card_number))
        conn.commit()
    finally:
        conn.close() 