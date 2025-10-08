import openpyxl

class JournalFiller:
    def __init__(self, filename):
        self.filename = filename
        self.wb = None
        self.START_ROW = 7
        self.HOURS_COLS = {'lecture': 12, 'practice': 13, 'lab': 14}
        self.load_workbook()
        
    def load_workbook(self):
        try:
            self.wb = openpyxl.load_workbook(self.filename)
            print(f"Файл '{self.filename}' загружен")
            return True
        except FileNotFoundError:
            print(f"Файл '{self.filename}' не найден")
            return False
    
    def get_existing_days(self, sheet):
        days_data = []
        row = self.START_ROW
        
        while sheet[f'E{row}'].value is not None:
            day_value = sheet[f'E{row}'].value
            if isinstance(day_value, (int, float)):
                days_data.append({
                    'row': row,
                    'day': int(day_value),
                    'discipline': sheet[f'F{row}'].value,
                    'group': sheet[f'G{row}'].value,
                    'load_type': sheet[f'H{row}'].value,
                    **{key: sheet.cell(row=row, column=col).value 
                       for key, col in self.HOURS_COLS.items()}
                })
            row += 1
        
        days_data.sort(key=lambda x: x['day'])
        return days_data
    
    def find_insert_position(self, sheet, new_day):
        existing_days = self.get_existing_days(sheet)
        
        if not existing_days:
            return self.START_ROW
        
        for day_data in existing_days:
            if day_data['day'] == new_day:
                return day_data['row']
        
        for i, day_data in enumerate(existing_days):
            if new_day < day_data['day']:
                return self.START_ROW if i == 0 else existing_days[i-1]['row'] + 1
        
        return existing_days[-1]['row'] + 1
    
    def insert_sorted_entry(self, sheet, insert_row, data):
        # Сохраняем существующие данные
        existing_data = []
        row = insert_row
        
        while sheet[f'E{row}'].value is not None:
            cell_data = {
                'day': sheet[f'E{row}'].value,
                'discipline': sheet[f'F{row}'].value,
                'group': sheet[f'G{row}'].value,
                'load_type': sheet[f'H{row}'].value
            }
            for key, col in self.HOURS_COLS.items():
                cell_data[key] = sheet.cell(row=row, column=col).value
            
            existing_data.append(cell_data)
            
            # Очищаем ячейки
            for col in ['E', 'F', 'G', 'H']:
                sheet[f'{col}{row}'] = None
            for col in self.HOURS_COLS.values():
                sheet.cell(row=row, column=col).value = None
            
            row += 1
        
        # Вставляем новую запись
        sheet[f'E{insert_row}'] = data['day_number']
        sheet[f'F{insert_row}'] = data['discipline']
        sheet[f'G{insert_row}'] = data['group']
        sheet[f'H{insert_row}'] = data['load_type']
        
        for key, col in self.HOURS_COLS.items():
            sheet.cell(row=insert_row, column=col).value = data.get(key)
        
        # Восстанавливаем данные
        for i, cell_data in enumerate(existing_data):
            new_row = insert_row + 1 + i
            sheet[f'E{new_row}'] = cell_data['day']
            sheet[f'F{new_row}'] = cell_data['discipline']
            sheet[f'G{new_row}'] = cell_data['group']
            sheet[f'H{new_row}'] = cell_data['load_type']
            for key, col in self.HOURS_COLS.items():
                sheet.cell(row=new_row, column=col).value = cell_data[key]
        
        return insert_row
    
    def add_entry(self, month, data):
        if not self.wb or month not in self.wb.sheetnames:
            print(f"Лист '{month}' не найден!")
            return False
        
        sheet = self.wb[month]
        
        try:
            insert_row = self.find_insert_position(sheet, data['day_number'])
            result_row = self.insert_sorted_entry(sheet, insert_row, data)
            
            print(f"Запись добавлена на лист '{month}' (строка {result_row}):")
            print(f"  Число: {data['day_number']}, Дисциплина: {data['discipline']}")
            print(f"  Группа: {data['group']}, Нагрузка: {data['load_type']}")
            for key in self.HOURS_COLS:
                if data.get(key):
                    print(f"  {key}: {data[key]}")
            return True
            
        except Exception as e:
            print(f"Ошибка: {e}")
            return False
    
    def save(self):
        if not self.wb:
            return False
        try:
            self.wb.save(self.filename)
            print("Файл сохранен")
            return True
        except Exception as e:
            print(f"Ошибка сохранения: {e}")
            return False
    
    def show_data(self, month):
        if not self.wb or month not in self.wb.sheetnames:
            print(f"Лист '{month}' не найден!")
            return
        
        existing_days = self.get_existing_days(self.wb[month])
        
        print(f"\nДанные на листе '{month}':")
        print("-" * 100)
        print("Строка | Число | Дисциплина          | Группа              | Нагрузка | Лекции | Практ. | Лаб.")
        print("-" * 100)
        
        for day in existing_days:
            print(f"{day['row']:6} | {day['day']:5} | {day['discipline']:20} | {day['group']:20} | "
                  f"{day['load_type']:8} | {day['lecture'] or '':6} | {day['practice'] or '':6} | {day['lab'] or '':4}")

def input_float(prompt, default=None):
    value = input(prompt).strip()
    if not value:
        return default
    try:
        return float(value)
    except ValueError:
        print("Ошибка: введите число")
        return input_float(prompt, default)

def validate_day(day):
    try:
        day_int = int(day)
        return (True, day_int) if 1 <= day_int <= 31 else (False, "Число должно быть от 1 до 31")
    except ValueError:
        return (False, "Введите корректное число")

def main():
    journal = JournalFiller("ПРимер.xlsx")
    
    if not journal.wb:
        return
    
    while True:
        print(f"\n{'='*50}")
        print("ЖУРНАЛ ПРЕПОДАВАТЕЛЯ")
        print(f"{'='*50}")
        print("Доступные листы:", ", ".join(journal.wb.sheetnames))
        print("\n1. Добавить запись")
        print("2. Показать данные")
        print("3. Сохранить и выйти")
        print("4. Выйти без сохранения")
        
        choice = input("\nВыберите действие: ").strip()
        
        if choice == '1':
            month = input("Месяц: ").strip()
            
            day_input = input("Число месяца (1-31): ").strip()
            valid, day_result = validate_day(day_input)
            if not valid:
                print(day_result)
                continue
            
            data = {
                'day_number': day_result,
                'discipline': input("Дисциплина: ").strip(),
                'group': input("Группа: ").strip(),
                'load_type': input("Нагрузка: ").strip(),
                'lecture': input_float("Лекции: "),
                'practice': input_float("Практические: "),
                'lab': input_float("Лабораторные: ")
            }
            
            if journal.add_entry(month, data):
                if input("Сохранить? (да/нет): ").lower() == 'да':
                    journal.save()
        
        elif choice == '2':
            month = input("Лист для просмотра: ").strip()
            journal.show_data(month)
        
        elif choice == '3':
            journal.save()
            break
        
        elif choice == '4':
            break

if __name__ == "__main__":
    main()