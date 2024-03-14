import time
from datetime import datetime
from ya_forms import YaFormsScraper
from excel_reform import ExcelReformatter
import os

# Путь к драйверу браузера
chrome_driver_path = 'C:/Users/User/Downloads/chromedriver-win64/chromedriver.exe'

# Путь к папке с профилем Chrome
user_data_dir = 'C:/Users/User/AppData/Local/Google/Chrome/User Data'

# URL страницы
form_url = 'https://forms.yandex.ru/cloud/admin/'

# Создаем экземпляр класса YaFormsScraper и запускаем скрапинг
def scrape_and_reformat():
    ya_scraper = YaFormsScraper(chrome_driver_path, user_data_dir)
    ya_scraper.scrape_form_data(form_url)

    # Указание путей для ExcelReformatter
    download_path = 'C:/Users/User/Downloads' # ВОТ ЗДЕСЬ НАДО ПОМЕНЯТЬ ТВОЙ ПУТЬ КУДА СОХРАНЯЮТСЯ ЭКСЕЛИ
    current_dir  = os.getcwd()
    results_dir = os.path.join(current_dir, 'results')
    archive_dir = os.path.join(current_dir, 'archive')

    # Создаем экземпляр класса ExcelReformatter и запускаем реформатирование
    excel_reformatter = ExcelReformatter(download_path, results_dir, archive_dir)
    excel_reformatter.reform_excel()

# Указание времени для запуска
schedules = {
    '2024-03-14': ['07:00', '11:00', '15:00', '19:00', '23:00'],
    '2024-03-16': ['03:00', '07:00', '11:00', '15:00', '19:00', '23:00'],
    '2024-03-17': ['03:00', '07:00', '11:00', '15:00', '19:00', '21:00'],
    '2024-03-18': ['01:00']
}

# Функция для запуска задач в зависимости от дня и времени
def run_schedule():
    current_time = datetime.now()
    current_date = current_time.strftime('%Y-%m-%d')
    current_hour = current_time.strftime('%H:%M')
    
    if current_date in schedules:
        if current_hour in schedules[current_date]:
            scrape_and_reformat()

# Запуск бесконечного цикла для проверки и запуска заданий в расписании
while True:
    run_schedule()
    time.sleep(60)  # Пауза 60 секунд перед следующей проверкой
