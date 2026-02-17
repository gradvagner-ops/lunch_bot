import asyncio
import logging
from datetime import datetime, time
from aiogram import Bot
import pytz

from config import ADMIN_ID
from utils import get_target_week_dates, get_week_range_display, get_deadline_status

logger = logging.getLogger(__name__)

# Московское время
MSK_TZ = pytz.timezone('Europe/Moscow')

class NotificationScheduler:
    def __init__(self, bot: Bot):
        self.bot = bot
        self.is_running = False
    
    async def send_reminder(self):
        """Отправка напоминания всем пользователям"""
        try:
            # Получаем список всех пользователей, кто хоть раз заказывал
            # Можно добавить отдельную таблицу для подписок на уведомления
            
            # Формируем сообщение
            target_dates, week_type, _ = get_target_week_dates()
            week_range = get_week_range_display(target_dates)
            deadline_status = get_deadline_status()
            
            reminder_text = (
                f"⏰ *НАПОМИНАНИЕ О ЗАКАЗЕ ОБЕДОВ*\n\n"
                f"📅 Сегодня пятница!\n\n"
                f"🍽️ *Нужно заказать обеды на следующую неделю:*\n"
                f"└ Период: `{week_range}`\n"
                f"└ {week_type}\n\n"
                f"⏳ *Дедлайн:* сегодня до 16:00\n\n"
                f"👇 Нажми «📝 Новый заказ» чтобы сделать заказ"
            )
            
            # TODO: Отправить всем пользователям
            # Пока отправим админу для теста
            await self.bot.send_message(
                ADMIN_ID,
                reminder_text,
                parse_mode="Markdown"
            )
            
            logger.info("✅ Напоминание отправлено")
            
        except Exception as e:
            logger.error(f"❌ Ошибка при отправке напоминания: {e}")
    
    async def scheduler_loop(self):
        """Бесконечный цикл проверки времени"""
        self.is_running = True
        logger.info("🔄 Планировщик уведомлений запущен")
        
        while self.is_running:
            try:
                # Текущее время в Москве
                now = datetime.now(MSK_TZ)
                
                # Проверяем: сегодня пятница и время 08:00
                if now.weekday() == 4 and now.hour == 8 and now.minute == 0:
                    await self.send_reminder()
                    # Ждем минуту, чтобы не отправить дважды
                    await asyncio.sleep(60)
                
                # Проверяем каждую минуту
                await asyncio.sleep(60)
                
            except Exception as e:
                logger.error(f"❌ Ошибка в планировщике: {e}")
                await asyncio.sleep(60)
    
    def stop(self):
        """Остановка планировщика"""
        self.is_running = False
        logger.info("🛑 Планировщик уведомлений остановлен")