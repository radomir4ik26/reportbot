import os
import logging 
import asyncio
from fastapi import FastAPI
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from docx import Document
import uvicorn

# Отримання токену та порту з змінних оточення
API_TOKEN = os.environ.get('API_TOKEN', '8035646713:AAGaYfc6NcmAHR0iseSNu7Vcs2N6tOodlXI')
PORT = int(os.environ.get('PORT', 10000))

# Налаштування бота
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()

# Решта коду залишається без змін до останніх функцій...

# Додаємо підтримку веб-сервера для Render
def create_web_app():
    """Функція для створення FastAPI додатку"""
    app = FastAPI()
    
    @app.get("/")
    async def root():
        return {"message": "Telegram Bot is running"}
    
    @app.get("/health")
    async def health_check():
        return {"status": "healthy"}
    
    return app

# Головна функція для запуску
async def main():
    try:
        # Перевіряємо наявність шаблону документа
        if not os.path.exists(template_path):
            logging.error(f"Шаблон документа '{template_path}' не знайдено!")
            print(f"Помилка: файл шаблону '{template_path}' не існує. Перевірте шлях до файлу.")
            return
        
        # Перевіряємо наявність файлів зі списками
        if not os.path.exists(names_path):
            logging.warning(f"Файл зі списком курсантів '{names_path}' не знайдено!")
            print(f"Попередження: файл '{names_path}' не існує. Буде використано порожній список.")
        
        if not os.path.exists(reasons_path):
            logging.warning(f"Файл зі списком причин '{reasons_path}' не знайдено!")
            print(f"Попередження: файл '{reasons_path}' не існує. Буде використано порожній список.")
        
        # Додаємо WebApp для health-check на Render
        web_app = create_web_app()
        
        # Створення конфігурації для uvicorn
        config = uvicorn.Config(
            web_app, 
            host="0.0.0.0", 
            port=PORT, 
            log_level="info"
        )
        server = uvicorn.Server(config)
        
        # Паралельний запуск бота та веб-сервера
        server_task = asyncio.create_task(server.serve())
        bot_task = asyncio.create_task(dp.start_polling(bot))
        
        # Очікуємо завершення обох задач
        await asyncio.gather(server_task, bot_task)
        
    except Exception as e:
        logging.error(f"Помилка при запуску бота: {e}")

# Точка входу
if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    asyncio.run(main())
