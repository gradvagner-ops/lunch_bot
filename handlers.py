from aiogram import F, types, Bot
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from datetime import datetime
import os
import logging
import asyncio
from concurrent.futures import ThreadPoolExecutor

from config import ADMIN_ID, WEEKDAYS
from database import Database
from keyboards import get_main_keyboard, get_remove_keyboard
from states import TextOrderState
from utils import (
    get_target_week_dates,
    get_deadline_status,
    get_week_range_display,
    format_date_for_db,
    create_excel_report
)
from cache import cache

executor = ThreadPoolExecutor(max_workers=1)
db = Database()
logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

# ==================== КОМАНДЫ ====================

async def cmd_start(message: types.Message):
    """🚀 Старт с подробной информацией"""
    user = message.from_user
    
    asyncio.create_task(register_user_async(user.id, user.username, user.full_name))
    
    target_dates, week_type, _ = get_target_week_dates()
    week_range = get_week_range_display(target_dates)
    deadline_status = get_deadline_status()
    
    await message.answer(
        f"👋 *Добрый день, {user.first_name}!*\n\n"
        f"🍽️ *Система заказа обедов для инструкторов*\n\n"
        f"📅 *Текущий период заказа:* `{week_range}`\n"
        f"└ {week_type}\n"
        f"{deadline_status}\n\n"
        f"📝 *Что нужно делать:*\n"
        f"• Нажать «📝 Новый заказ»\n"
        f"• Ввести ФИО инструктора\n"
        f"• На каждый день ввести **0**, **1** или **2**\n"
        f"• Проверить и подтвердить заказ\n\n"
        f"👇 *Нажмите кнопку «📝 Новый заказ» чтобы начать*",
        parse_mode="Markdown",
        reply_markup=get_main_keyboard(user.id == ADMIN_ID)
    )

async def register_user_async(user_id, username, full_name):
    """Фоновая регистрация"""
    try:
        db.register_employee(user_id, username, full_name)
    except:
        pass

# ==================== ПОДРОБНЫЙ ЗАКАЗ ====================

async def start_order(message: types.Message, state: FSMContext):
    """📝 Начало заказа с подробной информацией"""
    
    target_dates, week_type, _ = get_target_week_dates()
    date_keys = [format_date_for_db(d) for d in target_dates]
    week_range = get_week_range_display(target_dates)
    
    # Предрасчет всех форматов
    week_data = []
    for i, date_key in enumerate(date_keys):
        date_obj = cache.parse_date(date_key)
        week_data.append({
            'key': date_key,
            'day_name': WEEKDAYS[i],
            'display': date_obj.strftime("%d.%m.%Y"),
            'short': date_obj.strftime("%d.%m"),
            'full_date': date_obj.strftime("%d %B %Y"),
            'weekday_full': date_obj.strftime("%A")
        })
    
    await state.update_data(
        date_keys=date_keys,
        week_data=week_data,
        week_range=week_range,
        week_type=week_type,
        current_day=0,
        meals={}
    )
    
    await state.set_state(TextOrderState.waiting_instructor)
    
    await message.answer(
        f"📝 *Оформление нового заказа*\n\n"
        f"📅 *Период заказа:* `{week_range}`\n"
        f"└ {week_type}\n\n"
        f"👤 *Шаг 1 из 8:* Введите ФИО инструктора\n"
        f"└ Пример: *Иванов Иван Иванович*\n"
        f"└ Или: *Петрова Мария*\n\n"
        f"✏️ Напишите ФИО в ответном сообщении:",
        parse_mode="Markdown",
        reply_markup=get_remove_keyboard()
    )

async def process_instructor(message: types.Message, state: FSMContext):
    """Обработка ФИО с переходом к первому дню"""
    instructor = message.text.strip()
    
    if len(instructor) < 5:
        await message.answer(
            "❌ *Слишком короткое ФИО*\n\n"
            "Пожалуйста, введите полное ФИО:\n"
            "└ Пример: *Иванов Иван Иванович*",
            parse_mode="Markdown"
        )
        return
    
    await state.update_data(instructor=instructor)
    await state.set_state(TextOrderState.waiting_quantity)
    
    # Показываем первый день
    data = await state.get_data()
    day_info = data['week_data'][0]
    
    await message.answer(
        f"👤 *Инструктор:* {instructor}\n"
        f"📅 *Период:* {data['week_range']}\n\n"
        f"📝 *Шаг 2 из 8*\n"
        f"📅 *День 1: {day_info['day_name']}* ({day_info['display']})\n\n"
        f"🍽️ *Сколько обедов заказать на этот день?*\n\n"
        f"└ Введите **0** — не заказывать\n"
        f"└ Введите **1** — один обед\n"
        f"└ Введите **2** — два обеда\n\n"
        f"✏️ Напишите цифру (0, 1 или 2):",
        parse_mode="Markdown"
    )

async def process_quantity(message: types.Message, state: FSMContext):
    """⚡ Обработка числа с подробным подтверждением"""
    text = message.text.strip()
    
    if not text.isdigit() or text not in ('0', '1', '2'):
        await message.answer(
            "❌ *Неверный ввод*\n\n"
            "Пожалуйста, введите только **0**, **1** или **2**:\n"
            f"└ 0 — не заказывать\n"
            f"└ 1 — один обед\n"
            f"└ 2 — два обеда",
            parse_mode="Markdown"
        )
        return
    
    quantity = int(text)
    data = await state.get_data()
    current_day = data.get('current_day', 0)
    day_info = data['week_data'][current_day]
    
    # Сохраняем выбор
    meals = data.get('meals', {})
    meals[data['date_keys'][current_day]] = quantity
    
    # Показываем подтверждение выбора
    if quantity == 0:
        confirm = f"❌ *Не заказываем* обеды на {day_info['day_name']} ({day_info['short']})"
    elif quantity == 1:
        confirm = f"✅ *1 обед* на {day_info['day_name']} ({day_info['short']})"
    else:
        confirm = f"✅ *2 обеда* на {day_info['day_name']} ({day_info['short']})"
    
    await message.answer(confirm, parse_mode="Markdown")
    
    next_day = current_day + 1
    
    if next_day >= 7:
        # Все дни заполнены - показываем итоги
        await state.update_data(meals=meals)
        await show_summary(message, state)
        return
    
    # Показываем следующий день
    await state.update_data(
        meals=meals,
        current_day=next_day
    )
    
    next_day_info = data['week_data'][next_day]
    
    await message.answer(
        f"👤 *Инструктор:* {data['instructor']}\n"
        f"📅 *Период:* {data['week_range']}\n\n"
        f"📝 *Шаг {next_day + 2} из 8*\n"
        f"📅 *День {next_day + 1}: {next_day_info['day_name']}* ({next_day_info['display']})\n\n"
        f"🍽️ *Сколько обедов заказать на этот день?*\n\n"
        f"└ Введите **0** — не заказывать\n"
        f"└ Введите **1** — один обед\n"
        f"└ Введите **2** — два обеда\n\n"
        f"✏️ Напишите цифру (0, 1 или 2):",
        parse_mode="Markdown"
    )

async def show_summary(message: types.Message, state: FSMContext):
    """📋 Подробный показ итогов"""
    data = await state.get_data()
    meals = data.get('meals', {})
    week_data = data.get('week_data', [])
    instructor = data.get('instructor', '')
    week_range = data.get('week_range', '')
    
    # Подсчёт итогов
    total = 0
    days_count = 0
    lines = []
    
    for i, day_info in enumerate(week_data):
        qty = meals.get(day_info['key'], 0)
        if qty > 0:
            total += qty
            days_count += 1
            lines.append(f"✅ *{day_info['day_name']}* ({day_info['short']}): {qty} обед(ов)")
        else:
            lines.append(f"❌ *{day_info['day_name']}* ({day_info['short']}): 0")
    
    # Формируем подробный итог
    text = (
        f"📋 *Проверьте правильность заказа*\n\n"
        f"👤 *Инструктор:* {instructor}\n"
        f"📅 *Период:* {week_range}\n"
        f"📊 *Итого:* {days_count} дней, {total} обедов\n\n"
        f"*Детализация по дням:*\n"
    )
    
    text += "\n".join(lines)
    
    text += "\n\n⚠️ *Проверьте внимательно!*\n"
    text += "После подтверждения заказ будет сохранён."
    
    await state.update_data(total=total, days_count=days_count)
    await state.set_state(TextOrderState.waiting_confirm)
    
    from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Да, всё верно", callback_data="confirm_yes")],
            [InlineKeyboardButton(text="🔄 Заполнить заново", callback_data="confirm_no")],
            [InlineKeyboardButton(text="❌ Отменить заказ", callback_data="cancel")]
        ]
    )
    
    await message.answer(text, parse_mode="Markdown", reply_markup=keyboard)

async def confirm_order(callback: types.CallbackQuery, state: FSMContext):
    """✅ Подтверждение заказа"""
    
    if callback.data == "confirm_yes":
        data = await state.get_data()
        
        # Пакетное сохранение
        batch = []
        for date_key, qty in data.get('meals', {}).items():
            if qty > 0:
                batch.append((callback.from_user.id, data['instructor'], date_key, qty))
        
        if batch:
            asyncio.create_task(save_batch_async(batch))
        
        await state.clear()
        cache.clear_cache()
        
        await callback.message.edit_text(
            f"✅ *Заказ успешно подтверждён!*\n\n"
            f"👤 *Инструктор:* {data['instructor']}\n"
            f"📅 *Период:* {data['week_range']}\n"
            f"📊 *Всего:* {data['days_count']} дней, {data['total']} обедов\n\n"
            f"✨ Спасибо! Заказ передан на кухню.",
            parse_mode="Markdown"
        )
    
    elif callback.data == "confirm_no":
        await state.update_data(current_day=0, meals={})
        await state.set_state(TextOrderState.waiting_quantity)
        
        data = await state.get_data()
        day_info = data['week_data'][0]
        
        await callback.message.edit_text(
            f"🔄 *Начинаем заново*\n\n"
            f"👤 *Инструктор:* {data['instructor']}\n"
            f"📅 *Период:* {data['week_range']}\n\n"
            f"📅 *День 1: {day_info['day_name']}* ({day_info['display']})\n\n"
            f"🍽️ Сколько обедов? (0, 1, 2):",
            parse_mode="Markdown"
        )
    
    else:  # cancel
        await state.clear()
        await callback.message.edit_text(
            "❌ *Заказ отменён*\n\n"
            "Если передумаете, начните новый заказ.",
            parse_mode="Markdown"
        )
    
    await callback.message.answer(
        "👇 *Главное меню:*",
        parse_mode="Markdown",
        reply_markup=get_main_keyboard(callback.from_user.id == ADMIN_ID)
    )
    await callback.answer()

async def save_batch_async(batch):
    """Фоновое сохранение"""
    try:
        db.save_order_batch(batch)
    except:
        pass

# ==================== ПРОСМОТР ЗАКАЗОВ ====================

async def show_my_orders(message: types.Message):
    """📋 Подробный просмотр заказов"""
    user_id = message.from_user.id
    orders = db.get_user_orders_cached(user_id)
    
    if not orders:
        await message.answer(
            "📭 *У вас пока нет заказов*\n\n"
            "Нажмите «📝 Новый заказ» чтобы сделать первый заказ.",
            parse_mode="Markdown",
            reply_markup=get_main_keyboard(user_id == ADMIN_ID)
        )
        return
    
    # Группировка по инструкторам
    instructors = {}
    for name, date, qty in orders:
        if name not in instructors:
            instructors[name] = []
        instructors[name].append((date, qty))
    
    text = "📋 *Ваши текущие заказы*\n\n"
    total_all = 0
    
    for instructor, items in instructors.items():
        text += f"👤 *{instructor}*\n"
        instructor_total = 0
        
        for date, qty in sorted(items, reverse=True)[:7]:
            date_obj = cache.parse_date(date)
            date_str = date_obj.strftime("%a %d.%m")
            text += f"  • {date_str}: {qty} обед(ов)\n"
            instructor_total += qty
            total_all += qty
        
        text += f"  ✨ Итого по инструктору: {instructor_total} обедов\n\n"
    
    text += f"📊 *Всего заказов:* {total_all} обедов"
    
    await message.answer(
        text,
        parse_mode="Markdown",
        reply_markup=get_main_keyboard(user_id == ADMIN_ID)
    )

# ==================== АДМИНКА ====================

async def export_to_excel(message: types.Message, bot: Bot):
    """📊 Выгрузка Excel"""
    if message.from_user.id != ADMIN_ID:
        await message.answer("⛔ *Доступ запрещён*\n\nЭта команда только для администратора.")
        return
    
    status = await message.answer("🔄 *Формирую отчёт...*\nЭто займёт несколько секунд.")
    
    try:
        all_orders = db.get_all_orders_cached()
        if not all_orders:
            await status.edit_text("📭 *Нет заказов для выгрузки*")
            return
        
        target_dates, _, _ = get_target_week_dates()
        
        # В отдельном потоке
        loop = asyncio.get_event_loop()
        temp_path, saved_path = await loop.run_in_executor(
            executor, 
            create_excel_report, 
            all_orders, target_dates, True
        )
        
        await message.answer_document(
            types.FSInputFile(temp_path),
            caption=f"📊 *Отчёт по заказам готов*\n💾 Сохранён в папке exports/"
        )
        
        os.remove(temp_path)
        await status.delete()
        
    except Exception as e:
        await status.edit_text(f"❌ *Ошибка:* {str(e)[:50]}")