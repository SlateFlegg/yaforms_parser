import telebot
import os
import time

# Задайте токен вашего бота
TOKEN = "Token"
user_id = 000 # id Пользователя, кому отправлять файл

# Создаем экземпляр бота
bot = telebot.TeleBot(TOKEN)

# Функция для отправки файла пользователю
def send_file(file_path):
    print(file_path)
    # Отправляем файл пользователю
    if os.path.exists(file_path):
        bot.send_document(user_id, open(file_path, 'rb'))
    else:
        print("Файл не найден.")

# Функция для проверки наличия новых файлов и их отправки
def check_and_send_files():
    # Путь к директории, где бот будет искать новые файлы
    current_dir  = os.getcwd()
    directory = os.path.join(current_dir, 'results')
    
    while True:
        # Периодическая проверка наличия новых файлов
        files = os.listdir(directory)
        for file_name in files:
            if file_name.startswith('НАЗВАНИЕ ФАЙЛА'):  # Проверяем имя файла
                file_path = os.path.join(directory, file_name)
                send_file(file_path)  # Отправляем файл
                os.remove(file_path)  # Удаляем файл из директории
        time.sleep(10)  # Пауза между проверками

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "Привет! Я бот для отправки файлов.")

check_and_send_files()

# Запускаем бота
bot.polling()

