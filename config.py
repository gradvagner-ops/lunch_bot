import os

# Токен берем из переменных окружения (секретов)
TOKEN = os.getenv("BOT_TOKEN", "ваш_токен_здесь")

# ID администратора - ОБЯЗАТЕЛЬНО ДОБАВЬТЕ ЭТУ СТРОКУ!
ADMIN_ID = int(os.getenv("ADMIN_ID", "5046675535"))

# Настройки компании
COMPANY_NAME = "Игора"
EXPORT_PATH = "exports"

# Дни недели
WEEKDAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

# Дедлайн
DEADLINE_DAY = 4  # Пятница
DEADLINE_HOUR = 16
DEADLINE_MINUTE = 0