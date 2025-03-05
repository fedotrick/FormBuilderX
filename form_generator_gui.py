from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QFileDialog, QFormLayout, QMessageBox, QGroupBox,
                             QDateEdit, QTimeEdit)
from PySide6.QtCore import Qt, QDate, QTime, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont
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
        self.setMinimumWidth(800)
        self.setMinimumHeight(900)
        
        # Добавляем атрибут для хранения анимации
        self.fade_animation = None
        
        # Устанавливаем общий стиль приложения
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-family: 'Segoe UI', Arial;
                font-size: 14px;
                font-weight: bold;
                border: 2px solid #dcdcdc;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 10px;
                color: #2c3e50;
                background-color: #f5f5f5;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
                font-family: 'Segoe UI', Arial;
                font-size: 12px;
                min-height: 20px;
            }
            QLineEdit:focus {
                border: 2px solid #3498db;
            }
            QDateEdit, QTimeEdit {
                padding: 8px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
                font-family: 'Segoe UI', Arial;
                font-size: 12px;
                min-height: 20px;
            }
            QDateEdit:focus, QTimeEdit:focus {
                border: 2px solid #3498db;
            }
            QLabel {
                font-family: 'Segoe UI', Arial;
                font-size: 12px;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-family: 'Segoe UI', Arial;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #2472a4;
            }
        """)

        # Создаем центральный виджет и основной layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Создаем группы полей
        self.create_cast_group()
        self.create_gluing_group()
        self.create_control_group()

        # Добавляем группы в основной layout
        main_layout.addWidget(self.cast_group)
        main_layout.addWidget(self.gluing_group)
        main_layout.addWidget(self.control_group)

        # Добавляем кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.generate_btn = QPushButton("Сгенерировать")
        self.generate_btn.setFixedWidth(200)
        self.generate_btn.setFixedHeight(40)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.generate_btn.clicked.connect(self.generate_form)
        button_layout.addWidget(self.generate_btn)

        main_layout.addLayout(button_layout)
        
        # Устанавливаем текущую дату и время
        self.set_current_datetime()

    def create_cast_group(self):
        self.cast_group = QGroupBox("Информация об отливке")
        layout = QFormLayout()
        layout.setSpacing(15)  # Увеличиваем отступы между полями
        layout.setContentsMargins(20, 20, 20, 20)  # Добавляем отступы внутри группы

        self.fields = {}
        # Создаем поля для отливки
        self.fields['cast_number'] = QLineEdit()
        self.fields['cast_name'] = QLineEdit()
        self.fields['cluster_number'] = QLineEdit()

        # Добавляем подсказки
        self.fields['cast_number'].setPlaceholderText("Введите номер отливки")
        self.fields['cast_name'].setPlaceholderText("Введите наименование отливки")
        self.fields['cluster_number'].setPlaceholderText("Введите номер кластера")

        layout.addRow(self.create_label("Номер отливки (модели):"), self.fields['cast_number'])
        layout.addRow(self.create_label("Наименование отливки (модели):"), self.fields['cast_name'])
        layout.addRow(self.create_label("Номер кластера:"), self.fields['cluster_number'])

        self.cast_group.setLayout(layout)

    def create_gluing_group(self):
        self.gluing_group = QGroupBox("Склейка элементов")
        layout = QFormLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Создаем поля для склейки
        self.fields['gluing_date'] = QDateEdit()
        self.fields['gluing_date'].setCalendarPopup(True)  # Включаем выпадающий календарь
        self.fields['gluing_date'].setDisplayFormat("dd.MM.yyyy")  # Формат отображения даты
        self.fields['gluing_executor'] = QLineEdit()
        self.fields['gluing_quantity'] = QLineEdit()
        self.fields['gluing_notes'] = QLineEdit()

        # Добавляем подсказки
        self.fields['gluing_executor'].setPlaceholderText("Введите ФИО исполнителя")
        self.fields['gluing_quantity'].setPlaceholderText("Введите количество")
        self.fields['gluing_notes'].setPlaceholderText("Введите примечания")

        layout.addRow(self.create_label("Дата:"), self.fields['gluing_date'])
        layout.addRow(self.create_label("Исполнитель:"), self.fields['gluing_executor'])
        layout.addRow(self.create_label("Количество:"), self.fields['gluing_quantity'])
        layout.addRow(self.create_label("Примечание:"), self.fields['gluing_notes'])

        self.gluing_group.setLayout(layout)

    def create_control_group(self):
        self.control_group = QGroupBox("Контроль сборки")
        layout = QFormLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Создаем поля для контроля
        self.fields['control_date'] = QDateEdit()
        self.fields['control_date'].setCalendarPopup(True)  # Включаем выпадающий календарь
        self.fields['control_date'].setDisplayFormat("dd.MM.yyyy")  # Формат отображения даты
        
        self.fields['control_time'] = QTimeEdit()
        self.fields['control_time'].setDisplayFormat("HH:mm")  # Формат отображения времени
        
        self.fields['control_executor'] = QLineEdit()
        self.fields['control_quantity'] = QLineEdit()
        self.fields['control_notes'] = QLineEdit()

        # Добавляем подсказки
        self.fields['control_executor'].setPlaceholderText("Введите ФИО исполнителя")
        self.fields['control_quantity'].setPlaceholderText("Введите количество")
        self.fields['control_notes'].setPlaceholderText("Введите примечания")

        layout.addRow(self.create_label("Дата:"), self.fields['control_date'])
        layout.addRow(self.create_label("Время:"), self.fields['control_time'])
        layout.addRow(self.create_label("Исполнитель:"), self.fields['control_executor'])
        layout.addRow(self.create_label("Количество:"), self.fields['control_quantity'])
        layout.addRow(self.create_label("Примечание:"), self.fields['control_notes'])

        self.control_group.setLayout(layout)

    def create_label(self, text):
        label = QLabel(text)
        # Убираем установку шрифта здесь, так как она теперь в общих стилях
        return label

    def set_current_datetime(self):
        current_date = QDate.currentDate()
        current_time = QTime.currentTime()
        
        self.fields['gluing_date'].setDate(current_date)
        self.fields['control_date'].setDate(current_date)
        self.fields['control_time'].setTime(current_time)

    def validate_fields(self):
        required_fields = ['cluster_number']
        for field in required_fields:
            if not self.fields[field].text().strip():
                return False, f"Поле '{field}' обязательно для заполнения"
        return True, ""

    def clear_fields(self):
        # Очищаем все поля, кроме дат и времени
        for field_name, widget in self.fields.items():
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, (QDateEdit, QTimeEdit)):
                # Для полей даты и времени устанавливаем текущие значения
                continue
        
        # Обновляем текущую дату и время
        self.set_current_datetime()

    def generate_form(self):
        try:
            # Валидация полей
            is_valid, error_message = self.validate_fields()
            if not is_valid:
                QMessageBox.warning(self, "Предупреждение", error_message)
                return

            # Проверяем наличие файла шаблона
            if not os.path.exists("ШАБЛОН.pptx"):
                raise FileNotFoundError("Файл ШАБЛОН.pptx не найден в текущей директории")
            
            # Получаем данные из полей
            data = {}
            for field, widget in self.fields.items():
                if isinstance(widget, QDateEdit):
                    data[field] = widget.date().toString("dd.MM.yyyy")
                elif isinstance(widget, QTimeEdit):
                    data[field] = widget.time().toString("HH:mm")
                else:
                    data[field] = widget.text()
            
            # Проверяем, что номер кластера не пустой
            if not data['cluster_number']:
                raise ValueError("Номер кластера не может быть пустым")
            
            # Добавляем отладочную информацию
            print(f"Номер кластера для QR-кода: {data['cluster_number']}")
            
            # Генерируем форму
            self.generate_pptx_with_data("ШАБЛОН.pptx", data)
            
            QMessageBox.information(self, "Успех", "Форма успешно сгенерирована!")
            
            # Очищаем поля после успешной генерации
            self.clear_fields()
            
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

    def show(self):
        self.setWindowOpacity(1.0)  # Устанавливаем непрозрачность сразу
        super().show()  # Показываем окно

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 