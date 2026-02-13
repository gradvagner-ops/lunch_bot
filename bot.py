import os
import asyncio
import logging
from dotenv import load_dotenv
from aiogram import Bot, Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage

# Загружаем переменные окружения
load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

from handlers import *
from database import Database

logging.basicConfig(level=logging.INFO)

async def main():
    # Создаем папки
    os.makedirs("data", exist_ok=True)
    
    bot = Bot(token=TOKEN)
    storage = MemoryStorage()
    dp = Dispatcher(bot, storage=storage)
    
    # Регистрация обработчиков
    dp.register_message_handler(cmd_start, commands=["start"])
    dp.register_message_handler(start_order, text="📝 Новый заказ")
    dp.register_message_handler(process_instructor, state=TextOrderState.waiting_instructor)
    dp.register_message_handler(process_quantity, state=TextOrderState.waiting_quantity)
    dp.register_callback_query_handler(confirm_order, text=["confirm_yes", "confirm_no", "cancel"])
    dp.register_message_handler(show_my_orders, text="📋 Мои заказы")
    
    if ADMIN_ID:
        dp.register_message_handler(export_to_excel, text="📊 Выгрузить Excel")
    
    print(f"🚀 Бот запущен!")
    print(f"👑 Админ ID: {ADMIN_ID}")
    
    await dp.start_polling()

if __name__ == "__main__":
    asyncio.run(main())