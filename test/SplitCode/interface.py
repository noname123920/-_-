from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, 
                               QTableWidgetItem, QHeaderView, QMessageBox, QGroupBox, 
                               QGridLayout, QDateEdit, QRadioButton, QButtonGroup, 
                               QListWidget)
from PySide6.QtCore import Qt, QDate
from logic import JournalLogic


class JournalApp(QMainWindow):
    def __init__(self, filename):
        super().__init__()
        self.logic = JournalLogic(filename)
        self.setup_ui()
        self.load_workbook()
    
    def setup_ui(self):
        self.setWindowTitle("Журнал преподавателя")
        self.setGeometry(100, 100, 1200, 900)
        
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Период
        period_group = self.create_period_group()
        main_layout.addWidget(period_group)
        
        # Даты
        dates_group = self.create_dates_group()
        main_layout.addWidget(dates_group)
        
        # Поля ввода
        input_group = self.create_input_group()
        main_layout.addWidget(input_group)
        
        # Кнопка добавления
        add_button = QPushButton("Добавить записи")
        add_button.clicked.connect(self.add_entries)
        add_button.setMinimumHeight(35)
        add_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        main_layout.addWidget(add_button)
        
        # Просмотр данных
        view_group = self.create_view_group()
        main_layout.addWidget(view_group, 1)
    
    def create_period_group(self):
        group = QGroupBox("Выбор периода")
        layout = QGridLayout(group)
        
        # Начало периода
        layout.addWidget(QLabel("Начало периода:"), 0, 0)
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate(datetime.now().year, 9, 1))
        self.start_date.setDisplayFormat("dd.MM.yyyy")
        self.start_date.setCalendarPopup(True)
        layout.addWidget(self.start_date, 0, 1)
        
        # Конец периода
        layout.addWidget(QLabel("Конец периода:"), 0, 2)
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate(datetime.now().year, 12, 31))
        self.end_date.setDisplayFormat("dd.MM.yyyy")
        self.end_date.setCalendarPopup(True)
        layout.addWidget(self.end_date, 0, 3)
        
        # Тип недели
        layout.addWidget(QLabel("Тип недели:"), 0, 4)
        
        self.period_week_type = QButtonGroup(self)
        week_type_layout = QHBoxLayout()
        
        for text, value in [("Числитель", "числитель"), ("Знаменатель", "знаменатель"), ("Обе недели", "обе недели")]:
            radio = QRadioButton(text)
            radio.setProperty("value", value)
            self.period_week_type.addButton(radio)
            week_type_layout.addWidget(radio)
        
        self.period_week_type.buttons()[0].setChecked(True)
        layout.addLayout(week_type_layout, 0, 5, 1, 2)
        
        # Кнопки
        button_layout = QHBoxLayout()
        generate_btn = QPushButton("Сгенерировать даты по периоду")
        generate_btn.clicked.connect(self.generate_dates_by_period)
        button_layout.addWidget(generate_btn)
        
        clear_btn = QPushButton("Очистить даты")
        clear_btn.clicked.connect(self.clear_dates)
        button_layout.addWidget(clear_btn)
        
        layout.addLayout(button_layout, 1, 0, 1, 6)
        
        # Информация о датах
        self.dates_info_label = QLabel("Выбрано дат: 0")
        self.dates_info_label.setStyleSheet("color: blue; font-weight: bold;")
        layout.addWidget(self.dates_info_label, 2, 0, 1, 6)
        
        return group
    
    def create_dates_group(self):
        group = QGroupBox("Управление датами")
        layout = QHBoxLayout(group)
        
        # Левая часть - управление датами
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # Добавление даты
        add_date_layout = QHBoxLayout()
        add_date_layout.addWidget(QLabel("Добавить дату:"))
        self.single_date = QDateEdit()
        self.single_date.setDate(QDate.currentDate())
        self.single_date.setDisplayFormat("dd.MM.yyyy")
        self.single_date.setCalendarPopup(True)
        add_date_layout.addWidget(self.single_date)
        
        add_btn = QPushButton("Добавить дату")
        add_btn.clicked.connect(self.add_single_date)
        add_date_layout.addWidget(add_btn)
        add_date_layout.addStretch()
        
        left_layout.addLayout(add_date_layout)
        
        # Удаление даты
        remove_date_layout = QHBoxLayout()
        remove_date_layout.addWidget(QLabel("Удалить дату:"))
        self.remove_date_combo = QComboBox()
        remove_date_layout.addWidget(self.remove_date_combo)
        
        remove_btn = QPushButton("Удалить дату")
        remove_btn.clicked.connect(self.remove_selected_date)
        remove_date_layout.addWidget(remove_btn)
        remove_date_layout.addStretch()
        
        left_layout.addLayout(remove_date_layout)
        
        # Список дат
        left_layout.addWidget(QLabel("Выбранные даты:"))
        self.dates_listbox = QListWidget()
        self.dates_listbox.setMaximumHeight(120)
        left_layout.addWidget(self.dates_listbox)
        
        # Кнопки управления
        manage_buttons_layout = QHBoxLayout()
        clear_all_btn = QPushButton("Очистить все даты")
        clear_all_btn.clicked.connect(self.clear_dates)
        manage_buttons_layout.addWidget(clear_all_btn)
        
        refresh_btn = QPushButton("Обновить список")
        refresh_btn.clicked.connect(self.update_dates_display)
        manage_buttons_layout.addWidget(refresh_btn)
        manage_buttons_layout.addStretch()
        
        left_layout.addLayout(manage_buttons_layout)
        
        layout.addWidget(left_widget, 1)
        
        return group
    
    def create_input_group(self):
        group = QGroupBox("Данные для добавления")
        layout = QGridLayout(group)
        
        self.entries = {}
        fields = [
            ("Дисциплина:", "discipline", "entry"), 
            ("Группа:", "group", "entry"),
            ("Вид нагрузки:", "load_type", "combobox"),
            ("Лекции:", "lecture", "entry"),
            ("Практические:", "practice", "entry"),
            ("Лабораторные:", "lab", "entry")
        ]
        
        for i, (label, field, field_type) in enumerate(fields):
            layout.addWidget(QLabel(label), i, 0)
            
            if field_type == "combobox":
                widget = QComboBox()
                widget.addItems(self.logic.LOAD_TYPES)
                self.entries[field] = widget
            else:
                widget = QLineEdit()
                self.entries[field] = widget
            
            layout.addWidget(widget, i, 1)
        
        return group
    
    def create_view_group(self):
        group = QGroupBox("Просмотр данных")
        layout = QVBoxLayout(group)
        
        # Панель управления
        control_layout = QHBoxLayout()
        
        control_layout.addWidget(QLabel("Лист:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.show_data)
        control_layout.addWidget(self.sheet_combo)
        
        # Кнопки управления
        self.delete_btn = QPushButton("Удалить выбранные записи")
        self.delete_btn.clicked.connect(self.delete_selected_entries)
        self.delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 3px;
                padding: 5px 10px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        control_layout.addWidget(self.delete_btn)
        
        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.clicked.connect(self.select_all_entries)
        control_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("Снять выделение")
        self.deselect_all_btn.clicked.connect(self.deselect_all_entries)
        control_layout.addWidget(self.deselect_all_btn)
        
        control_layout.addStretch()
        
        layout.addLayout(control_layout)
        
        # Информация о выборе
        self.selection_info = QLabel("Выбрано записей: 0")
        self.selection_info.setStyleSheet("color: green; font-weight: bold;")
        layout.addWidget(self.selection_info)
        
        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Число", "Дисциплина", "Группа", "Нагрузка", "Лекции", "Практические", "Лабораторные"
        ])
        
        # Настройка таблицы
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.MultiSelection)
        self.table.itemSelectionChanged.connect(self.on_table_selection_changed)
        
        layout.addWidget(self.table, 1)
        
        return group
    
    def on_table_selection_changed(self):
        """Обновляет информацию о количестве выбранных записей"""
        selected_count = len(self.table.selectedItems()) // self.table.columnCount()
        self.selection_info.setText(f"Выбрано записей: {selected_count}")
    
    def select_all_entries(self):
        """Выбирает все записи в таблице"""
        self.table.selectAll()
    
    def deselect_all_entries(self):
        """Снимает выделение со всех записей"""
        self.table.clearSelection()
    
    def generate_dates_by_period(self):
        """Генерирует даты по периоду"""
        # Получаем выбранный тип недели
        target_week_type = None
        for btn in self.period_week_type.buttons():
            if btn.isChecked():
                target_week_type = btn.property("value")
                break
        
        if not target_week_type:
            QMessageBox.critical(self, "Ошибка", "Не выбран тип недели")
            return
        
        success, message = self.logic.generate_dates_by_period(
            self.start_date.date(), 
            self.end_date.date(), 
            target_week_type
        )
        
        if success:
            QMessageBox.information(self, "Успех", message)
        else:
            QMessageBox.critical(self, "Ошибка", message)
        
        self.update_dates_display()
    
    def add_single_date(self):
        """Добавляет одиночную дату"""
        success, message = self.logic.add_single_date(self.single_date.date())
        
        if success:
            QMessageBox.information(self, "Успех", message)
        else:
            QMessageBox.critical(self, "Ошибка", message)
        
        self.update_dates_display()
    
    def remove_selected_date(self):
        """Удаляет выбранную дату"""
        selected_index = self.remove_date_combo.currentIndex()
        success, message = self.logic.remove_selected_date(selected_index)
        
        if success:
            QMessageBox.information(self, "Успех", message)
        else:
            QMessageBox.critical(self, "Ошибка", message)
        
        self.update_dates_display()
    
    def clear_dates(self):
        """Очищает все даты"""
        success, message = self.logic.clear_dates()
        
        if success:
            QMessageBox.information(self, "Успех", message)
        else:
            QMessageBox.critical(self, "Ошибка", message)
        
        self.update_dates_display()
    
    def update_dates_display(self):
        """Обновляет отображение дат"""
        # Обновляем информацию о датах
        self.dates_info_label.setText(self.logic.get_dates_info())
        
        # Обновляем список дат
        display_texts, date_values = self.logic.get_dates_for_display()
        
        self.dates_listbox.clear()
        self.dates_listbox.addItems(display_texts)
        
        self.remove_date_combo.clear()
        self.remove_date_combo.addItems(date_values)
        if date_values:
            self.remove_date_combo.setCurrentIndex(0)
    
    def load_workbook(self):
        """Загружает рабочую книгу"""
        success, message = self.logic.load_workbook()
        
        if success:
            # Обновляем список листов
            sheets = self.logic.get_sheets()
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheets)
            if sheets:
                self.sheet_combo.setCurrentIndex(0)
            
            self.show_data()
        else:
            QMessageBox.critical(self, "Ошибка", message)
    
    def show_data(self):
        """Показывает данные из выбранного листа"""
        sheet_name = self.sheet_combo.currentText()
        if not sheet_name:
            return
        
        data_rows = self.logic.get_sheet_data(sheet_name)
        
        self.table.setRowCount(len(data_rows))
        for i, row_data in enumerate(data_rows):
            for j, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(i, j, item)
        
        self.selection_info.setText("Выбрано записей: 0")
    
    def delete_selected_entries(self):
        """Удаляет выбранные записи"""
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "Внимание", "Выберите записи для удаления")
            return
        
        # Получаем индексы выбранных строк
        selected_rows = set()
        for range in selected_ranges:
            for row in range.topRow(), range.bottomRow() + 1:
                selected_rows.add(row)
        
        selected_rows = sorted(selected_rows)
        if not selected_rows:
            return
        
        entries_to_delete = []
        for row in selected_rows:
            if row < self.table.rowCount():
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                entries_to_delete.append(row_data)
        
        confirm = QMessageBox.question(
            self, 
            "Подтверждение удаления", 
            f"Вы действительно хотите удалить {len(entries_to_delete)} записей?\n"
            f"Это действие нельзя отменить.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if confirm != QMessageBox.Yes:
            return
        
        sheet_name = self.sheet_combo.currentText()
        success, message = self.logic.delete_entries(sheet_name, entries_to_delete)
        
        if success:
            QMessageBox.information(self, "Успех", message)
            self.show_data()
        else:
            QMessageBox.critical(self, "Ошибка", message)
    
    def add_entries(self):
        """Добавляет записи в журнал"""
        # Собираем данные из полей ввода
        data = {
            'discipline': self.entries['discipline'].text(),
            'group': self.entries['group'].text(),
            'load_type': self.entries['load_type'].currentText(),
        }
        
        # Обрабатываем числовые поля
        try:
            lecture = self.entries['lecture'].text()
            practice = self.entries['practice'].text()
            lab = self.entries['lab'].text()
            
            data['lecture'] = float(lecture) if lecture else 0.0
            data['practice'] = float(practice) if practice else 0.0
            data['lab'] = float(lab) if lab else 0.0
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Проверьте числовые поля (Лекции, Практические, Лабораторные) - они должны содержать только числа")
            return
        
        success, message = self.logic.add_entries(data)
        
        if success:
            QMessageBox.information(self, "Успех", message)
            # Очищаем поля ввода
            for field in ['discipline', 'group', 'lecture', 'practice', 'lab']:
                if field in self.entries and isinstance(self.entries[field], QLineEdit):
                    self.entries[field].clear()
            
            if 'load_type' in self.entries:
                self.entries['load_type'].setCurrentIndex(0)
            
            self.show_data()
        else:
            QMessageBox.critical(self, "Ошибка", message)
    
    def closeEvent(self, event):
        """Обработчик закрытия окна"""
        self.logic.close_workbook()
        event.accept()


# Импорт для datetime
from datetime import datetime