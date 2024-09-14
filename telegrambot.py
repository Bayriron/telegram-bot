import os
import json
import pandas as pd
import openpyxl  # Явный импорт openpyxl
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes
import telegram.ext.filters as filters
import logging
from flask import Flask, request
import subprocess
from dotenv import load_dotenv

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Загрузка переменных окружения
load_dotenv()
TOKEN = os.getenv('BOT_TOKEN')

# Директория для хранения статистики
STATS_DIR = 'stats'
ANSWER_FILE = os.path.join(STATS_DIR, 'answers.json')
STATS_FILE = os.path.join(STATS_DIR, 'stats.json')
EXCEL_FILE = os.path.join(STATS_DIR, 'student_stats.xlsx')

# Создание директории для статистики, если она не существует
if not os.path.exists(STATS_DIR):
    os.makedirs(STATS_DIR)

# Словарь для отслеживания состояния пользователей
user_states = {}
cached_stats = None  # Переменная для хранения закэшированной статистики

def load_answer_key():
    """Загружает ключи ответов из файла answers.json с валидацией данных."""
    try:
        with open(ANSWER_FILE, 'r', encoding='utf-8') as file:
            data = json.load(file)
        answers = data.get('answers', [])
        
        # Проверка на корректность данных
        if not isinstance(answers, list) or not all(isinstance(answer, str) for answer in answers):
            raise ValueError("Некорректный формат данных в answers.json.")
        return answers
    except FileNotFoundError:
        logger.error(f"Файл {ANSWER_FILE} не найден.")
        return []
    except json.JSONDecodeError:
        logger.error(f"Ошибка декодирования JSON в файле {ANSWER_FILE}.")
        return []
    except ValueError as e:
        logger.error(e)
        return []

ANSWER_KEY = load_answer_key()

def normalize_answer(answer):
    """Нормализует ответ, удаляя пробелы и приводя к нижнему регистру."""
    return ''.join(answer.split()).lower()

def load_stats():
    """Загружает статистику из файла stats.json. Используется кэширование для улучшения производительности."""
    global cached_stats
    if cached_stats is None:
        if os.path.exists(STATS_FILE):
            try:
                with open(STATS_FILE, 'r', encoding='utf-8') as file:
                    cached_stats = json.load(file).get('users', {})
            except json.JSONDecodeError:
                logger.error("Ошибка декодирования JSON. Проверьте формат файла stats.json.")
                cached_stats = {}
        else:
            cached_stats = {}
    return cached_stats

def save_stats(stats):
    """Сохраняет статистику в файл stats.json."""
    global cached_stats
    cached_stats = stats  # Обновляем кэш
    data = {
        "users": stats
    }
    try:
        with open(STATS_FILE, 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False)
    except IOError as e:
        logger.error(f"Ошибка записи в файл {STATS_FILE}: {e}")

def save_stats_to_excel(stats):
    """Сохраняет статистику в Excel файл."""
    data = []
    for user_id, user_data in stats.items():
        total_correct = sum(user_data['scores'])
        data.append({
            'Имя': user_data['first_name'],
            'Фамилия': user_data['last_name'],
            'Правильных ответов': total_correct
        })
    df = pd.DataFrame(data)
    try:
        df.to_excel(EXCEL_FILE, index=False)
    except IOError as e:
        logger.error(f"Ошибка записи в файл {EXCEL_FILE}: {e}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обрабатывает команду /start и отправляет приветственное сообщение."""
    logger.info("Команда /start вызвана")
    keyboard = [
        ['/get_test'],
        ['/show_stats']
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard)
    await update.message.reply_text('Привет! Я ваш бот для проверки тестов.', reply_markup=reply_markup)

async def send_test(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Отправляет тестовые изображения пользователю."""
    logger.info("Команда /get_test вызвана")
    chat_id = update.message.chat_id
    user_id = update.message.from_user.id
    user_states[user_id] = 'test_sent'  # Обновляем состояние пользователя
    file_paths = [
        'C:\\Users\\User\\Desktop\\phyton\\Снимок экрана 2024-09-11 111528.png',
        'C:\\Users\\User\\Desktop\\phyton\\Снимок экрана 2024-09-11 111539.png',
        # Добавьте пути к другим изображениям
    ]
    for file_path in file_paths:
        try:
            with open(file_path, 'rb') as image:
                await context.bot.send_photo(chat_id=chat_id, photo=image)
        except FileNotFoundError:
            await update.message.reply_text(f"Файл {file_path} не найден. Пожалуйста, проверьте путь к файлу.")
            logger.error(f"Файл {file_path} не найден.")
    
    recommendations = (
        "Для решения тестов советуется отводить время 1 час.\n"
        "Отправляйте ответы в нижнем регистре без пробелов.\n"
        "Пример: abcdeabcdeabcdeabcdeabcdeabcde"
    )
    await update.message.reply_text(recommendations)

async def submit_answers(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обрабатывает ответы пользователя и отправляет результат."""
    logger.info("Команда /submit_answers вызвана")
    user_id = update.message.from_user.id
    user_first_name = update.message.from_user.first_name
    user_last_name = update.message.from_user.last_name
    user_answers = list(update.message.text.strip())  # Преобразуем строку в список символов
    correct_answers = 0
    incorrect_answers = []

    # Валидация входных данных
    if len(user_answers) != len(ANSWER_KEY):
        await update.message.reply_text("Неправильное количество ответов. Пожалуйста, проверьте и отправьте снова.")
        logger.warning(f"Неправильное количество ответов от пользователя {user_id}.")
        return

    for i, answer in enumerate(user_answers):
        if i < len(ANSWER_KEY) and normalize_answer(answer) == normalize_answer(ANSWER_KEY[i]):
            correct_answers += 1
        else:
            incorrect_answers.append((i + 1, answer))

    percentage = calculate_percentage(correct_answers, len(ANSWER_KEY))
    response = f'Правильных ответов: {correct_answers}/{len(ANSWER_KEY)} ({percentage}%)\nОшибки: {incorrect_answers}'
    await update.message.reply_text(response)

    # Сохранение статистики только для первой попытки
    stats = load_stats()
    if user_id not in stats or len(stats[user_id]['scores']) == 0:
        stats[user_id] = {
            'first_name': user_first_name,
            'last_name': user_last_name,
            'scores': [correct_answers]
        }
    else:
        logger.info(f"Пользователь {user_id} уже отправлял ответы. Повторная попытка не будет добавлена в статистику.")

    save_stats(stats)
    save_stats_to_excel(stats)

    # Обновляем состояние пользователя
    user_states[user_id] = 'answers_submitted'

def calculate_percentage(correct_answers, total_questions):
    """Вычисляет процент правильных ответов."""
    if total_questions == 0:
        return 0
    return round((correct_answers / total_questions) * 100)

async def show_stats(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Отображает статистику пользователей."""
    logger.info("Команда /show_stats вызвана")
    stats = load_stats()
    sorted_stats = sorted(stats.items(), key=lambda x: sum(x[1]['scores']), reverse=True)
    response = 'Статистика:\n'
    total_questions = len(ANSWER_KEY)
    for user_id, data in sorted_stats:
        total_correct = sum(data['scores'])
        percentage = calculate_percentage(total_correct, total_questions)
        response += f'{data["first_name"]} {data["last_name"]}: {total_correct} правильных ответов ({percentage}%)\n'
    await update.message.reply_text(response)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обрабатывает текстовые сообщения пользователей."""
    try:
        user_id = update.message.from_user.id
        if user_states.get(user_id) == 'answers_submitted':
            await update.message.reply_text('Вы уже отправили свои ответы. Пожалуйста, используйте команды.')
        else:
            await submit_answers(update, context)
    except Exception as e:
        logger.error(f"Ошибка при обработке сообщения: {e}")
        await update.message.reply_text('Произошла ошибка при обработке вашего сообщения.')

# Flask сервер для обработки вебхуков
app = Flask(__name__)

@app.route('/@/Nodirtest_bot', methods=['POST'])
def webhook():
    update = request.get_json()
    if update:
        application.update_queue.put(update)
        return 'OK'
    else:
        logger.error("Получено пустое обновление")
        return 'Bad Request', 400

def set_webhook():
    """Устанавливает вебхук для бота."""
    import requests
    TOKEN = os.getenv('BOT_TOKEN')
    WEBHOOK_URL = 'https://0634-213-230-86-246.ngrok-free.app/@/Nodirtest_bot'
    url = f'https://api.telegram.org/bot{TOKEN}/setWebhook'
    response = requests.post(url, data={'url': WEBHOOK_URL})
    if response.status_code == 200:
        logger.info(f"Вебхук успешно установлен: {WEBHOOK_URL}")
    else:
        logger.error(f"Ошибка установки вебхука: {response.text}")

def start_ngrok():
    """Запускает ngrok и возвращает публичный URL."""
    process = subprocess.Popen(['ngrok', 'http', '5000'], stdout=subprocess.PIPE)
    for line in process.stdout:
        if b'url=' in line:
            url = line.decode('utf-8').split('url=')[1].strip()
            logger.info(f"ngrok URL: {url}")
            return url
    logger.error("Не удалось получить URL ngrok")
    return None

if __name__ == '__main__':
    # Запуск ngrok и получение публичного URL
    public_url = start_ngrok()
    WEBHOOK_URL = f'{public_url}/@/Nodirtest_bot'
    
    set_webhook()
    
    application = ApplicationBuilder().token(TOKEN).build()
    start_handler = CommandHandler('start', start)
    send_test_handler = CommandHandler('get_test', send_test)
    submit_answers_handler = MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)
    show_stats_handler = CommandHandler('show_stats', show_stats)
    
    application.add_handler(start_handler)
    application.add_handler(send_test_handler)
    application.add_handler(submit_answers_handler)
    application.add_handler(show_stats_handler)
    
    app.run(port=5000)
