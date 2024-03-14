import os
import pandas as pd
import datetime
import shutil
import re

class ExcelReformatter:
    def __init__(self, download_path, results_dir, archive_dir):
        self.download_path = download_path
        self.results_dir = results_dir
        self.archive_dir = archive_dir

    def reform_excel(self):
        current_datetime = datetime.datetime.now()
        current_datetime_str = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")

        # Функция для поиска файла с заданным текстом в имени
        def find_file(directory, text):
            for filename in os.listdir(directory):
                if text in filename:
                    return os.path.join(directory, filename)
            return None

        # Нахождение файла
        file_name = find_file(self.download_path, 'ЧАСТЬ НАЗВАНИЯ ФАЙЛА') # Здесь указать хотя бы часть названия файла, который будет обрабатываться после скачивания

        if file_name:
            df = pd.read_excel(file_name)

            df.rename(columns={'ID': 'region'}, inplace=True)
            df['region'] = 'РЕГИОН' # Указать ваш регион
            df.drop(columns=['Время создания'], inplace=True)

            # Замена всех символов, кроме цифр, из столбца 'Телефон (в формате 9xxxxxxxxx)' на пустую строку
            df['Телефон (в формате 9xxxxxxxxx)'] = df['Телефон (в формате 9xxxxxxxxx)'].apply(lambda x: re.sub(r'\D', '', str(x)))
            # Замена пустых строк на NaN
            df['Телефон (в формате 9xxxxxxxxx)'] = df['Телефон (в формате 9xxxxxxxxx)'].replace('', pd.NA)
            # Конвертация строковых значений в целочисленный формат (NaN останется NaN)
            df['mob_phone'] = df['Телефон (в формате 9xxxxxxxxx)'].apply(lambda x: int(x) if pd.notna(x) else x)
            df.drop(columns=['Телефон (в формате 9xxxxxxxxx)'], inplace=True)

            df['voice'] = 1

            # Сохранение файла
            output_file = os.path.join(self.results_dir, f'НАЗВАНИЕ ФАЙЛА{current_datetime_str}.csv') # Здесь указать как файл должен будет называться
            df.to_csv(output_file, index=False, encoding='utf-8')
            print(f"Результат сохранен в {output_file}")

            # Перемещение исходного файла в архив
            shutil.move(file_name, os.path.join(self.archive_dir, os.path.basename(file_name)))
            
        else:
            print("Файл не найден.")

