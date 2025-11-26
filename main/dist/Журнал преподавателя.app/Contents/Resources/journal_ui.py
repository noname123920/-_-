from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, 
                              QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
                              QHeaderView, QGroupBox, QMessageBox, QFileDialog, QListWidget,
                              QListWidgetItem, QAbstractItemView, QRadioButton,
                              QButtonGroup, QDateEdit, QSplitter, QDialog, QDialogButtonBox, QTextBrowser, QSizePolicy)
from PySide6.QtCore import Qt, QDate, QTimer
from PySide6.QtGui import QPalette, QColor, QFont, QMovie, QAction
import os

class JournalApp(QMainWindow):
    def __init__(self, logic_handler):
        super().__init__()
        self.logic_handler = logic_handler
        self.logic_handler.set_ui(self)
        self.setup_ui()
        
    def setup_ui(self):
        """Настраивает пользовательский интерфейс с приоритетом для просмотра"""
        self.setWindowTitle("Журнал преподавателя")
        self.setGeometry(100, 100, 1400, 900)
        
        # Настройка цветовой схемы
        self.setup_colors()
        
        # Создание центрального виджета
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(5)
        main_layout.setContentsMargins(8, 8, 8, 8)
        
        # Создание виджетов в порядке важности
        self.create_menu_bar()
        self.create_file_section(main_layout)
        self.create_period_section(main_layout)
        self.create_input_section(main_layout)
        
        # Секция просмотра получает максимальное пространство
        self.create_view_section(main_layout)
        main_layout.setStretchFactor(main_layout.itemAt(main_layout.count()-1).widget(), 1)
        
        # Если был сохранен последний файл, пытаемся загрузить его
        if self.logic_handler.filename and os.path.exists(self.logic_handler.filename):
            QTimer.singleShot(100, lambda: self.logic_handler.load_workbook(self.logic_handler.filename))
    
    def setup_colors(self):
        """Настраивает цветовую схему приложения"""
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(255, 230, 242))
        palette.setColor(QPalette.WindowText, QColor(139, 0, 139))
        palette.setColor(QPalette.Base, QColor(255, 240, 245))
        palette.setColor(QPalette.AlternateBase, QColor(255, 182, 193))
        palette.setColor(QPalette.Button, QColor(255, 102, 178))
        palette.setColor(QPalette.ButtonText, QColor(0, 0, 0))
        palette.setColor(QPalette.Highlight, QColor(152, 251, 152))
        palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))
        palette.setColor(QPalette.Text, QColor(0, 0, 0))
        palette.setColor(QPalette.BrightText, QColor(0, 0, 0))
        self.setPalette(palette)
    
    def create_menu_bar(self):
        """Создает меню приложения"""
        menubar = self.menuBar()
        
        # Меню Файл
        file_menu = menubar.addMenu("Файл")
        open_action = QAction("Открыть файл Excel", self)
        open_action.triggered.connect(self.logic_handler.open_file)
        file_menu.addAction(open_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Меню Помощь
        help_menu = menubar.addMenu("Помощь")
        instructions_action = QAction("Инструкция", self)
        instructions_action.triggered.connect(self.logic_handler.show_instructions)
        help_menu.addAction(instructions_action)
        
        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.logic_handler.show_about)
        help_menu.addAction(about_action)
    
    def create_file_section(self, parent_layout):
        """Создает секцию выбора файла"""
        file_group = QGroupBox("Рабочий файл")
        file_layout = QHBoxLayout(file_group)
        
        # Поле пути к файлу
        self.file_path_label = QLabel("Файл не выбран")
        self.file_path_label.setStyleSheet("QLabel { background-color: #fff0f5; padding: 5px; border: 1px solid #ffb6c1; }")
        self.file_path_label.setMinimumHeight(30)
        file_layout.addWidget(self.file_path_label)
        
        # Кнопки управления файлом
        btn_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("Выбрать файл")
        self.select_file_btn.setStyleSheet(self.get_action_button_style())
        self.select_file_btn.clicked.connect(self.logic_handler.open_file)
        btn_layout.addWidget(self.select_file_btn)
        
        file_layout.addLayout(btn_layout)
        
        parent_layout.addWidget(file_group)
    
    def create_period_section(self, parent_layout):
        """Создает секцию выбора периода"""
        period_group = QGroupBox("Выбор периода")
        period_layout = QGridLayout(period_group)
        
        # Даты
        period_layout.addWidget(QLabel("Начало периода:"), 0, 0)
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate().addMonths(-3))
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("dd.MM.yyyy")
        period_layout.addWidget(self.start_date, 0, 1)
        
        period_layout.addWidget(QLabel("Конец периода:"), 0, 2)
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("dd.MM.yyyy")
        period_layout.addWidget(self.end_date, 0, 3)
        
        # Тип недели
        period_layout.addWidget(QLabel("Тип недели:"), 0, 4)
        
        self.week_type_group = QButtonGroup(self)
        self.numerator_radio = QRadioButton("Числитель")
        self.denominator_radio = QRadioButton("Знаменатель")
        self.both_weeks_radio = QRadioButton("Обе недели")
        
        self.week_type_group.addButton(self.numerator_radio)
        self.week_type_group.addButton(self.denominator_radio)
        self.week_type_group.addButton(self.both_weeks_radio)
        
        self.numerator_radio.setChecked(True)
        
        radio_layout = QHBoxLayout()
        radio_layout.addWidget(self.numerator_radio)
        radio_layout.addWidget(self.denominator_radio)
        radio_layout.addWidget(self.both_weeks_radio)
        period_layout.addLayout(radio_layout, 0, 5)
        
        # Кнопки
        btn_layout = QHBoxLayout()
        self.generate_dates_btn = QPushButton("Сгенерировать даты по периоду")
        self.generate_dates_btn.setStyleSheet(self.get_action_button_style())
        self.generate_dates_btn.clicked.connect(self.logic_handler.generate_dates_by_period)
        btn_layout.addWidget(self.generate_dates_btn)
        
        self.clear_dates_btn = QPushButton("Очистить даты")
        self.clear_dates_btn.setStyleSheet(self.get_action_button_style())
        self.clear_dates_btn.clicked.connect(self.logic_handler.clear_dates)
        btn_layout.addWidget(self.clear_dates_btn)
        
        period_layout.addLayout(btn_layout, 1, 0, 1, 6)
        
        # Информация о датах
        self.dates_info_label = QLabel("Выбрано дат: 0")
        self.dates_info_label.setStyleSheet("QLabel { color: blue; font-weight: bold; }")
        period_layout.addWidget(self.dates_info_label, 2, 0, 1, 6)
        
        parent_layout.addWidget(period_group)
    
    def create_input_section(self, parent_layout):
        """Создает секцию ввода данных"""
        input_splitter = QSplitter(Qt.Horizontal)
        
        # Колонка 1: Управление датами
        dates_group = QGroupBox("Управление датами")
        dates_layout = QVBoxLayout(dates_group)
        
        # Добавление даты
        add_date_layout = QHBoxLayout()
        add_date_layout.addWidget(QLabel("Добавить дату:"))
        self.single_date = QDateEdit()
        self.single_date.setDate(QDate.currentDate())
        self.single_date.setCalendarPopup(True)
        self.single_date.setDisplayFormat("dd.MM.yyyy")
        add_date_layout.addWidget(self.single_date)
        
        self.add_date_btn = QPushButton("Добавить дату")
        self.add_date_btn.setStyleSheet(self.get_action_button_style())
        self.add_date_btn.clicked.connect(self.logic_handler.add_single_date)
        add_date_layout.addWidget(self.add_date_btn)
        
        dates_layout.addLayout(add_date_layout)
        
        # Удаление даты
        remove_date_layout = QHBoxLayout()
        remove_date_layout.addWidget(QLabel("Удалить дату:"))
        self.remove_date_combo = QComboBox()
        self.remove_date_combo.setMinimumWidth(150)
        remove_date_layout.addWidget(self.remove_date_combo)
        
        self.remove_date_btn = QPushButton("Удалить дату")
        self.remove_date_btn.setStyleSheet(self.get_action_button_style())
        self.remove_date_btn.clicked.connect(self.logic_handler.remove_selected_date)
        remove_date_layout.addWidget(self.remove_date_btn)
        
        dates_layout.addLayout(remove_date_layout)
        
        # Список выбранных дат
        dates_layout.addWidget(QLabel("Выбранные даты:"))
        self.dates_listbox = QListWidget()
        self.dates_listbox.setMinimumHeight(150)
        dates_layout.addWidget(self.dates_listbox)
        
        input_splitter.addWidget(dates_group)
        
        # Колонка 2: Поля ввода
        data_group = QGroupBox("Данные для заполнения")
        data_layout = QGridLayout(data_group)
        
        fields = [
            ("Дисциплина:", "discipline", "combobox"),
            ("Группа:", "group", "entry"),
            ("Вид нагрузки:", "load_type", "combobox"),
            ("Лекции:", "lecture", "entry"),
            ("Практические:", "practice", "entry"),
            ("Лабораторные:", "lab", "entry")
        ]
        
        self.entries = {}
        for i, (label, field, field_type) in enumerate(fields):
            data_layout.addWidget(QLabel(label), i, 0)
            
            if field_type == "combobox":
                if field == "discipline":
                    self.entries[field] = QComboBox()
                    self.entries[field].setEditable(True)
                else:
                    self.entries[field] = QComboBox()
                    self.entries[field].addItems(self.logic_handler.LOAD_TYPES)
            else:
                self.entries[field] = QLineEdit()
            
            data_layout.addWidget(self.entries[field], i, 1)
        
        input_splitter.addWidget(data_group)
        
        # Колонка 3: GIF/Анимация
        gif_group = QGroupBox("")
        gif_layout = QVBoxLayout(gif_group)
        
        # Загрузка и отображение GIF
        self.gif_label = QLabel("Журнал\nпреподавателя")
        self.gif_label.setAlignment(Qt.AlignCenter)
        self.gif_label.setStyleSheet("QLabel { font-size: 16px; font-weight: bold; color: #8b008b; }")
        self.gif_label.setMinimumSize(200, 200)
        gif_layout.addWidget(self.gif_label)
        
        # Пытаемся загрузить GIF
        self.logic_handler.load_and_display_gif()
        
        input_splitter.addWidget(gif_group)
        
        # Установка пропорций
        input_splitter.setSizes([300, 400, 200])
        
        parent_layout.addWidget(input_splitter)
        
        # Кнопка добавления записей
        self.add_entries_btn = QPushButton("Добавить записи")
        self.add_entries_btn.setStyleSheet(self.get_action_button_style())
        self.add_entries_btn.clicked.connect(self.logic_handler.add_entries)
        self.add_entries_btn.setEnabled(False)
        parent_layout.addWidget(self.add_entries_btn)
    
    def create_view_section(self, parent_layout):
        """Создает секцию просмотра данных"""
        view_group = QGroupBox("Просмотр и управление данными")
        view_layout = QVBoxLayout(view_group)
        
        # Панель управления просмотром
        control_layout = QHBoxLayout()
        
        control_layout.addWidget(QLabel("Лист:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.logic_handler.show_data)
        control_layout.addWidget(self.sheet_combo)
        
        # Кнопки управления
        self.refresh_btn = QPushButton("Обновить данные")
        self.refresh_btn.setStyleSheet(self.get_action_button_style())
        self.refresh_btn.clicked.connect(self.logic_handler.show_data)
        control_layout.addWidget(self.refresh_btn)
        
        self.delete_btn = QPushButton("Удалить выбранные")
        self.delete_btn.setStyleSheet(self.get_danger_button_style())
        self.delete_btn.clicked.connect(self.logic_handler.delete_selected_entries)
        control_layout.addWidget(self.delete_btn)
        
        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.setStyleSheet(self.get_action_button_style())
        self.select_all_btn.clicked.connect(self.logic_handler.select_all_entries)
        control_layout.addWidget(self.select_all_btn)
        
        self.deselect_btn = QPushButton("Снять выделение")
        self.deselect_btn.setStyleSheet(self.get_action_button_style())
        self.deselect_btn.clicked.connect(self.logic_handler.deselect_all_entries)
        control_layout.addWidget(self.deselect_btn)
        
        control_layout.addStretch()
        
        view_layout.addLayout(control_layout)
        
        # Информация о выборе
        self.selection_info = QLabel("Выбрано записей: 0")
        self.selection_info.setStyleSheet("QLabel { color: green; font-weight: bold; }")
        view_layout.addWidget(self.selection_info)
        
        # Таблица данных
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(7)
        self.table_widget.setHorizontalHeaderLabels(["Число", "Дисциплина", "Группа", "Нагрузка", "Лекции", "Практические", "Лабораторные"])
        
        # Настройка таблицы
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        self.table_widget.setAlternatingRowColors(True)
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        self.table_widget.verticalHeader().setVisible(False)
        
        # Установка ширины колонок
        column_widths = [80, 250, 120, 100, 80, 100, 100]
        for i, width in enumerate(column_widths):
            self.table_widget.setColumnWidth(i, width)
        
        view_layout.addWidget(self.table_widget)
        
        parent_layout.addWidget(view_group)
    
    def get_action_button_style(self):
        """Возвращает стиль для кнопок действий"""
        return """
            QPushButton {
                background-color: #98fb98;
                color: #006400;
                border: 2px solid #90ee90;
                border-radius: 5px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #90ee90;
            }
            QPushButton:pressed {
                background-color: #7cfc00;
            }
        """
    
    def get_danger_button_style(self):
        """Возвращает стиль для опасных кнопок"""
        return """
            QPushButton {
                background-color: #ff6b6b;
                color: #8b0000;
                border: 2px solid #ff5252;
                border-radius: 5px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #ff5252;
            }
            QPushButton:pressed {
                background-color: #ff3838;
            }
        """
    
    def show_info_dialog(self, title, text):
        """Создает диалоговое окно с информацией"""
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setModal(True)
        dialog.resize(600, 500)
        
        layout = QVBoxLayout(dialog)
        
        text_edit = QTextBrowser()
        text_edit.setPlainText(text)
        layout.addWidget(text_edit)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)
        
        dialog.exec()
    
    def closeEvent(self, event):
        """Обработчик закрытия приложения"""
        self.logic_handler.close_workbook()
        self.logic_handler.save_config()
        event.accept()