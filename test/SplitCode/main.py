import sys
import os
from PySide6.QtWidgets import QApplication
from interface import JournalApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Устанавливаем стиль приложения
    app.setStyle('Fusion')
    
    filename = "Тетрадь_ППС_2025_2026_каф_NN_Фамилия_ИО_оч_заоч.xltx"
    
    # Проверяем существование файла
    if not os.path.exists(filename):
        # Пробуем найти файл с другим расширением
        possible_extensions = ['.xltx', '.xlsx', '.xls']
        found = False
        for ext in possible_extensions:
            alt_filename = filename.replace('.xltx', ext)
            if os.path.exists(alt_filename):
                filename = alt_filename
                found = True
                break
        
        if not found:
            print(f"Файл не найден: {filename}")
            print(f"Текущая директория: {os.getcwd()}")
            print(f"Файлы в директории: {os.listdir('.')}")
            sys.exit(1)
    
    print(f"Загружаем файл: {filename}")
    
    window = JournalApp(filename)
    window.show()
    
    sys.exit(app.exec())