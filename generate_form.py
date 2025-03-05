from PIL import Image
import qrcode
import os

def generate_form_with_qr(template_path, output_path, qr_data):
    # Открываем шаблон
    template = Image.open(template_path)
    
    # Создаем QR-код
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=1,  # Еще немного уменьшаем отступы
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    qr_image = qr.make_image(fill_color="black", back_color="white")
    
    # Уменьшаем размер QR-кода до 50 пикселей
    qr_size = 50
    qr_image = qr_image.resize((qr_size, qr_size))
    
    # Вычисляем позицию для QR-кода в правом верхнем углу
    # Оставляем отступ в 20 пикселей от правого края и 10 пикселей сверху
    margin_right = 20
    margin_top = 10  # Уменьшаем верхний отступ с 20 до 10 пикселей
    x_position = template.width - qr_size - margin_right
    y_position = margin_top
    
    # Создаем новое изображение с тем же размером, что и шаблон
    result = template.copy()
    
    # Вставляем QR-код
    result.paste(qr_image, (x_position, y_position))
    
    # Сохраняем результат
    result.save(output_path)

def main():
    template_path = "ШАБЛОН.png"
    
    while True:
        # Получаем данные от пользователя
        print("\nГенератор QR-кодов для маршрутных карт")
        print("-" * 40)
        qr_data = input("Введите данные для QR-кода (или 'выход' для завершения): ")
        
        if qr_data.lower() in ['выход', 'exit', 'quit', 'q']:
            print("Программа завершена.")
            break
        
        # Формируем имя выходного файла
        output_path = f"маршрутная_карта_{len(qr_data)}.png"
        counter = 1
        while os.path.exists(output_path):
            output_path = f"маршрутная_карта_{len(qr_data)}_{counter}.png"
            counter += 1
        
        try:
            generate_form_with_qr(template_path, output_path, qr_data)
            print(f"Файл успешно создан: {output_path}")
        except Exception as e:
            print(f"Произошла ошибка: {str(e)}")

if __name__ == "__main__":
    main() 