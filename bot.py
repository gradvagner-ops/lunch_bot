import os
import asyncio
import logging
from dotenv import load_dotenv
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.fsm.storage.memory import MemoryStorage

# Загружаем переменные окружения
load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

# Импорты из ваших файлов
from handlers import *
from database import Database
from states import TextOrderState
from scheduler import NotificationScheduler  # 👈 Новый импорт

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

async def main():
    # Создаем папки
    os.makedirs("data", exist_ok=True)
    
    # Инициализация бота
    bot = Bot(token=TOKEN)
    storage = MemoryStorage()
    dp = Dispatcher(storage=storage)
    
    # Создаем планировщик
    scheduler = NotificationScheduler(bot)
    
    # Регистрация обработчиков
    dp.message.register(cmd_start, Command("start"))
    dp.message.register(start_order, F.text == "📝 Новый заказ")
    dp.message.register(process_instructor, TextOrderState.waiting_instructor)
    dp.message.register(process_quantity, TextOrderState.waiting_quantity)
    dp.callback_query.register(confirm_order, F.data.in_(["confirm_yes", "confirm_no", "cancel"]))
    dp.message.register(show_my_orders, F.text == "📋 Мои заказы")
    
    # Команды для подписки/отписки (можно добавить позже)
    dp.message.register(subscribe_notifications, F.text == "🔔 Подписаться на уведомления")
    dp.message.register(unsubscribe_notifications, F.text == "🔕 Отписаться")
    
    if ADMIN_ID:
        dp.message.register(export_to_excel, F.text == "📊 Выгрузить Excel")
        dp.message.register(show_excel_history, F.text == "📚 Архив Excel")
    
    print(f"🚀 Бот запущен на aiogram 3.x!")
    print(f"👑 Админ ID: {ADMIN_ID}")
    print(f"⏰ Планировщик уведомлений: пятница 08:00 МСК")
    
    # Запускаем планировщик в фоне
    asyncio.create_task(scheduler.scheduler_loop())
    
    # Запуск polling
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())