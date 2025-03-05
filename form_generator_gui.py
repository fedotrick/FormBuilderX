from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QFileDialog, QFormLayout, QMessageBox)
from pptx import Presentation
import qrcode
import os
from io import BytesIO
import sys
from datetime import datetime

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Генератор маршрутных карт")
        self.setMinimumWidth(600)

        # Создаем центральный виджет и основной layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Создаем форму для ввода данных
        form_layout = QFormLayout()

        # Создаем поля ввода
        self.fields = {
            'cast_number': QLineEdit(),
            'cast_name': QLineEdit(),
            'cluster_number': QLineEdit(),
            
            'gluing_date': QLineEdit(),
            'gluing_executor': QLineEdit(),
            'gluing_quantity': QLineEdit(),
            'gluing_notes': QLineEdit(),
            
            'control_date': QLineEdit(),
            'control_time': QLineEdit(),
            'control_executor': QLineEdit(),
            'control_quantity': QLineEdit(),
            'control_notes': QLineEdit(),
        }

        # Добавляем поля в форму
        form_layout.addRow("Номер отливки (модели):", self.fields['cast_number'])
        form_layout.addRow("Наименование отливки (модели):", self.fields['cast_name'])
        form_layout.addRow("Номер кластера:", self.fields['cluster_number'])
        
        form_layout.addRow("Склейка элементов п/м - Дата:", self.fields['gluing_date'])
        form_layout.addRow("Склейка элементов п/м - Исполнитель:", self.fields['gluing_executor'])
        form_layout.addRow("Склейка элементов п/м - Количество:", self.fields['gluing_quantity'])
        form_layout.addRow("Склейка элементов п/м - Примечание:", self.fields['gluing_notes'])
        
        form_layout.addRow("Контроль сборки - Дата:", self.fields['control_date'])
        form_layout.addRow("Контроль сборки - Время:", self.fields['control_time'])
        form_layout.addRow("Контроль сборки - Исполнитель:", self.fields['control_executor'])
        form_layout.addRow("Контроль сборки - Количество:", self.fields['control_quantity'])
        form_layout.addRow("Контроль сборки - Примечание:", self.fields['control_notes'])

        # Добавляем кнопки
        button_layout = QHBoxLayout()
        self.generate_btn = QPushButton("Сгенерировать")
        self.generate_btn.clicked.connect(self.generate_form)
        button_layout.addWidget(self.generate_btn)

        # Добавляем layouts на форму
        layout.addLayout(form_layout)
        layout.addLayout(button_layout)

        # Устанавливаем текущую дату и время в соответствующие поля
        current_datetime = datetime.now()
        self.fields['gluing_date'].setText(current_datetime.strftime("%d.%m.%Y"))
        self.fields['control_date'].setText(current_datetime.strftime("%d.%m.%Y"))
        self.fields['control_time'].setText(current_datetime.strftime("%H:%M"))

    def generate_form(self):
        try:
            # Проверяем наличие файла шаблона
            if not os.path.exists("ШАБЛОН.pptx"):
                raise FileNotFoundError("Файл ШАБЛОН.pptx не найден в текущей директории")
            
            # Получаем данные из полей
            data = {field: widget.text() for field, widget in self.fields.items()}
            
            # Проверяем, что номер кластера не пустой
            if not data['cluster_number']:
                raise ValueError("Номер кластера не может быть пустым")
            
            # Добавляем отладочную информацию
            print(f"Номер кластера для QR-кода: {data['cluster_number']}")
            
            # Генерируем форму
            self.generate_pptx_with_data("ШАБЛОН.pptx", data)
            
            QMessageBox.information(self, "Успех", "Форма успешно сгенерирована!")
        except FileNotFoundError as e:
            QMessageBox.critical(self, "Ошибка", str(e))
        except ValueError as e:
            QMessageBox.critical(self, "Ошибка", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Неожиданная ошибка: {str(e)}\n\n"
                               f"Тип ошибки: {type(e).__name__}")

    def generate_pptx_with_data(self, template_path, data):
        # Открываем шаблон
        prs = Presentation(template_path)
        
        # Проверяем наличие слайдов
        if not prs.slides:
            raise ValueError("Презентация не содержит слайдов")
            
        slide = prs.slides[0]
        
        # Создаем QR-код с обработкой ошибок
        try:
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=1,
            )
            print(f"Создание QR-кода для значения: {data['cluster_number']}")
            qr.add_data(str(data['cluster_number']))  # Явно преобразуем в строку
            qr.make(fit=True)
            qr_image = qr.make_image(fill_color="black", back_color="white")
        except Exception as e:
            print(f"Ошибка при создании QR-кода: {str(e)}")
            raise

        # Сохраняем QR-код во временный буфер
        image_stream = BytesIO()
        qr_image.save(image_stream, format='PNG')
        image_stream.seek(0)

        # Добавляем QR-код рядом с "МАРШРУТНАЯ КАРТА"
        for shape in slide.shapes:
            if hasattr(shape, "text") and "МАРШРУТНАЯ КАРТА" in shape.text:
                qr_width = 400000
                qr_height = 400000
                left = shape.left + shape.width + 200000
                top = shape.top + (shape.height - qr_height) / 2
                slide.shapes.add_picture(
                    image_stream,
                    left,
                    top,
                    width=qr_width,
                    height=qr_height
                )

        # Работаем с таблицами
        tables_found = False
        for shape in slide.shapes:
            if shape.has_table:
                tables_found = True
                table = shape.table
                print(f"Найдена таблица: {len(table.rows)} строк, {len(table.columns)} столбцов")
                
                # Проверяем первую ячейку таблицы, чтобы определить тип таблицы
                table_header = table.cell(0, 0).text.strip()
                print(f"Заголовок таблицы: '{table_header}'")

                # Если это таблица с номером отливки
                if "Номер" in table_header:
                    print("Обрабатываем таблицу с номером отливки")
                    try:
                        # Заполняем только данные отливки
                        table.cell(1, 0).text = data['cast_number']
                        table.cell(1, 1).text = data['cast_name']
                        table.cell(1, 2).text = data['cluster_number']
                    except Exception as e:
                        print(f"Ошибка при заполнении данных отливки: {str(e)}")
                    continue  # Переходим к следующей таблице

                # Если это таблица операций
                try:
                    # Сначала найдем индексы нужных столбцов
                    column_indices = {}
                    print("\nЗаголовки столбцов:")
                    for col in range(len(table.columns)):
                        header = table.cell(0, col).text.strip()
                        print(f"Столбец {col}: '{header}'")
                        if header == "Дата":
                            column_indices['date'] = col
                        elif header == "Время":
                            column_indices['time'] = col
                        elif header == "Исполнитель":
                            column_indices['executor'] = col
                        elif header == "Количество":
                            column_indices['quantity'] = col
                        elif header == "Примечание":
                            column_indices['notes'] = col
                    
                    print("\nНайденные индексы столбцов:", column_indices)

                    # Теперь найдем строки операций и заполним данные
                    print("\nСодержимое первого столбца:")
                    for row in range(len(table.rows)):
                        try:
                            operation = table.cell(row, 0).text.strip()
                            print(f"Строка {row}: '{operation}'")
                            
                            # Заполняем данные для Склейки
                            if "Склейка элементов п/м" in operation:
                                print("Найдена строка Склейки")
                                if 'date' in column_indices:
                                    table.cell(row, column_indices['date']).text = data['gluing_date']
                                if 'executor' in column_indices:
                                    table.cell(row, column_indices['executor']).text = data['gluing_executor']
                                if 'quantity' in column_indices:
                                    table.cell(row, column_indices['quantity']).text = data['gluing_quantity']
                                if 'notes' in column_indices:
                                    table.cell(row, column_indices['notes']).text = data['gluing_notes']
                            
                            # Заполняем данные для Контроля
                            if "Контроль сборки кластера" in operation:
                                print("Найдена строка Контроля")
                                if 'date' in column_indices:
                                    # Записываем дату
                                    table.cell(row, column_indices['date']).text = data['control_date']
                                    
                                    # Ищем ячейку со значением "Время:"
                                    for r in range(len(table.rows)):
                                        for c in range(len(table.columns)):
                                            cell_text = table.cell(r, c).text.strip()
                                            if cell_text == "Время:":
                                                print("Найдена ячейка 'Время:'")
                                                # Записываем время в ячейку справа
                                                if c + 1 < len(table.columns):
                                                    table.cell(r, c + 1).text = data['control_time']
                                
                                if 'executor' in column_indices:
                                    table.cell(row, column_indices['executor']).text = data['control_executor']
                                    table.cell(row, column_indices['quantity']).text = data['control_quantity']
                                    table.cell(row, column_indices['notes']).text = data['control_notes']
                                
                        except Exception as e:
                            print(f"Ошибка при обработке строки {row}: {str(e)}")
                            
                except Exception as e:
                    print(f"Ошибка при обработке таблицы: {str(e)}")
        
        if not tables_found:
            print("На слайде не найдено ни одной таблицы!")

        # Формируем имя выходного файла, заменяя недопустимые символы
        cluster_number = data['cluster_number'].replace('/', '_').replace('\\', '_')
        output_path = f"маршрутная_карта_{cluster_number}.pptx"
        counter = 1
        while os.path.exists(output_path):
            output_path = f"маршрутная_карта_{cluster_number}_{counter}.pptx"
            counter += 1

        # Сохраняем результат
        prs.save(output_path)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 