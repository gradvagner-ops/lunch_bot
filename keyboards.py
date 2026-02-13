from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove

def get_remove_keyboard():
    """Убирает клавиатуру"""
    return ReplyKeyboardRemove()

def get_main_keyboard(is_admin=False):
    """Главное меню"""
    keyboard = [
        [KeyboardButton(text="📝 Новый заказ")],
        [KeyboardButton(text="📋 Мои заказы")]
    ]
    if is_admin:
        keyboard.append([KeyboardButton(text="📊 Выгрузить Excel")])
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)