import config
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
import pandas as pd
from datetime import datetime
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import warnings
warnings.filterwarnings("ignore")

def download(chrome_driver_path:str, user_data_dir:str, form_url:str) -> None:
    service = Service(chrome_driver_path)

    # Настройки для использования профиля
    options = webdriver.ChromeOptions()
    options.add_argument(f'--user-data-dir={user_data_dir}')

    # Запуск веб-драйвера Chrome с использованием сервиса и настройками
    driver = webdriver.Chrome(options=options)

    # Открытие страницы
    driver.get(form_url)

    # Ожидание загрузки страницы
    time.sleep(10)

    # Нахождение кнопки и нажатие на нее
    download_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div/div[1]/div[2]/article/main/div[1]/div/button[2]")
    download_button.click()

    # Ожидание загрузки файла
    time.sleep(20)
    driver.quit()

def find_file(directory:str, title:str) -> str:
            for filename in os.listdir(directory):
                if title in filename:
                    return os.path.join(directory, filename)
            return None

def send_email(file_path:str) -> None:
    smtp_server = "smtp.yandex.ru"
    smtp_port = 587
    sender_email =  config.sender_email
    sender_password = config.sender_password
    receivers = config.receivers

    current_date = datetime.now().strftime("%d.%m.%y %H:%M")

    # Создание письма
    subject = f"Проход без карты {current_date}"
    body = config.body

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(receivers)#[0] #receiv  er_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    # Добавление вложения
    attachment = open(file_path, "rb")  # Открываем файл в бинарном режиме

    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)  # Кодируем файл в base64

    # Указываем заголовок для вложения
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(file_path)}",
    )

    message.attach(part)  # Прикрепляем файл к письму
    attachment.close()  # Закрываем файл

    # Подключение к SMTP-серверу и отправка письма
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()
            server.starttls()  # Включаем шифрование TLS
            server.login(sender_email, sender_password)  # Авторизация
            server.sendmail(sender_email, receivers, message.as_string())  # Отправка
            print(f"Письмо с вложением успешно отправлено! {current_date}")
    except Exception as e:
        print(f"Ошибка при отправке письма: {e}")

def excel_proccessing(download_path:str, results_dir:str, file_name:str) -> None:
    path = find_file(download_path, file_name)
    df = pd.read_excel(path)
    df['Время создания'] = pd.to_datetime(df['Время создания'])
    today = datetime.today().date()
    today_df = df[df['Время создания'].dt.date == today]
    today_df['klass'] = today_df['Параллель'].str.replace('класс', '') + today_df['Буква класса']
    today_df['name'] = today_df['Найдите своё имя:']
    today_df['surname'] = today_df['Найдите фамилию (или введите вручную ниже):']
    for cl in today_df.columns:
        if "имя" in cl:
            today_df['name'] = today_df['name'].fillna(df[cl])
        if "фамил" in cl:
            today_df['surname'] = today_df['surname'].fillna(df[cl])

    today_df['Час выгрузки'] = today_df['Время создания'].apply(lambda x: "9:00" if x.time() < datetime.strptime('09:00:00', '%H:%M:%S').time() else "11:00" )

    styled_df = today_df[['name', 'surname','Введите фамилию:', 'klass', 'Время создания', 'Час выгрузки']]
    st_col = styled_df.columns
    ren_list = ['Имя', "Фамилия", "Ручной ввод фамилии", "Класс", "Отметка времени", 'Час выгрузки']
    renamed_columns_dict = dict(zip(st_col, ren_list))
    styled_df = styled_df.rename(columns=renamed_columns_dict)

    file_path = results_dir+file_name+'.xlsx'
    styled_df.to_excel(file_path, index=False)
    send_email(file_path)
    os.remove(path)

def scrape_and_reformat(download_path, save_dir,chrome_driver_path, user_data_dir, form_url):
    
    download(chrome_driver_path, user_data_dir, form_url)

    today = datetime.today().date()
    is_workday = (datetime.today().weekday() != 5 and datetime.today().weekday() != 6)

    file_name = f'{today.year}-{today.month:02d}-{today.day:02d}'

    download_path = config.download_path
    results_dir = os.path.join(save_dir, 'results')
    file_name = f'{today.year}-{today.month:02d}-{today.day:02d} Prokh'

    excel_proccessing(download_path, results_dir, file_name)

def run_schedule(schedules, save_dir,  download_path, chrome_driver_path, user_data_dir, form_url):
    current_hour = datetime.now().strftime('%H:%M')

    if current_hour in schedules:
        if is_workday:
            scrape_and_reformat(download_path, save_dir, chrome_driver_path, user_data_dir, form_url)



schedules = config.schedules
download_path = config.download_path
# Путь к драйверу браузера
chrome_driver_path = config.chrome_driver_path
# Путь к папке с профилем Chrome
user_data_dir = config.user_data_dir
# URL страницы
form_url = config.form_url
save_dir = config.save_dir

#Запуск бесконечного цикла для проверки и запуска заданий в расписании
while True:
    run_schedule(schedules, save_dir, download_path, chrome_driver_path, user_data_dir, form_url)
    time.sleep(60)  # Пауза 60 секунд перед следующей проверкой