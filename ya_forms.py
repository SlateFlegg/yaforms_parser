from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

class YaFormsScraper:
    def __init__(self, chrome_driver_path, user_data_dir):
        self.chrome_driver_path = chrome_driver_path
        self.user_data_dir = user_data_dir

    def scrape_form_data(self, form_url):
        # Создание сервиса ChromeDriver
        service = Service(self.chrome_driver_path)

        # Настройки для использования профиля
        options = webdriver.ChromeOptions()
        options.add_argument(f'--user-data-dir={self.user_data_dir}')

        # Запуск веб-драйвера Chrome с использованием сервиса и настройками
        driver = webdriver.Chrome(service=service, options=options)

        # Открытие страницы
        driver.get(form_url)

        # Ожидание загрузки страницы
        time.sleep(5)

        # Нахождение кнопки и нажатие на нее
        download_button = driver.find_element(By.XPATH, "//button[contains(@class, 'yc-button') and span[text()='Скачать XLSX']]")
        download_button.click()

        # Ожидание загрузки файла
        time.sleep(10)

        # Закрываем браузер
        driver.quit()
