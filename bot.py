import os
import asyncio
import logging
from dotenv import load_dotenv
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, CallbackQuery

# Загружаем переменные окружения
load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

# Импорты из ваших файлов
from handlers import *
from database import Database
from states import TextOrderState  # Убедитесь, что этот импорт есть

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

async def main():
    # Создаем папки
    os.makedirs("data", exist_ok=True)
    
    # Инициализация бота для aiogram 3.x
    bot = Bot(token=TOKEN)
    storage = MemoryStorage()
    dp = Dispatcher(storage=storage)
    
    # Регистрация обработчиков для aiogram 3.x
    dp.message.register(cmd_start, Command("start"))
    dp.message.register(start_order, F.text == "📝 Новый заказ")
    dp.message.register(process_instructor, TextOrderState.waiting_instructor)
    dp.message.register(process_quantity, TextOrderState.waiting_quantity)
    dp.callback_query.register(confirm_order, F.data.in_(["confirm_yes", "confirm_no", "cancel"]))
    dp.message.register(show_my_orders, F.text == "📋 Мои заказы")
    
    if ADMIN_ID:
        dp.message.register(export_to_excel, F.text == "📊 Выгрузить Excel")
    
    print(f"🚀 Бот запущен")
    print(f"👑 Админ ID: {ADMIN_ID}")
    
    # Запуск polling
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())