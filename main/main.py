import sys
from PySide6.QtWidgets import QApplication, QMessageBox
from journal_ui import JournalApp
from journal_logic import JournalLogic

def main():
    """Основная функция запуска приложения"""
    try:
        app = QApplication(sys.argv)
        
        # Установка стиля приложения
        app.setStyle('Fusion')
        
        # Создаем обработчик логики
        logic_handler = JournalLogic()
        
        # Создаем UI и передаем ему обработчик логики
        window = JournalApp(logic_handler)
        window.show()
        
        sys.exit(app.exec())
    except Exception as e:
        print(f"Не удалось запустить приложение: {e}")
        QMessageBox.critical(None, "Ошибка запуска", f"Не удалось запустить приложение: {e}")

if __name__ == "__main__":
    main()