from aiogram.fsm.state import State, StatesGroup

class TextOrderState(StatesGroup):
    """Текстовый заказ - только ввод чисел"""
    waiting_instructor = State()  # Ждём ФИО
    waiting_quantity = State()    # Ждём число (0,1,2)
    waiting_confirm = State()     # Ждём подтверждения