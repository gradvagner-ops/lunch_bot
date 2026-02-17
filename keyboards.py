from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove

def get_remove_keyboard():
    """Убирает клавиатуру"""
    return ReplyKeyboardRemove()

def get_main_keyboard(is_admin=False):
    """Главное меню"""
    keyboard = [
        [KeyboardButton(text="📝 Новый заказ")],
        [KeyboardButton(text="📋 Мои заказы")],
        [KeyboardButton(text="🔔 Подписаться на уведомления")],
        [KeyboardButton(text="🔕 Отписаться")]
    ]
    if is_admin:
        keyboard.append([KeyboardButton(text="📊 Выгрузить Excel")])
        keyboard.append([KeyboardButton(text="📚 Архив Excel")])  # Новая кнопка
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)