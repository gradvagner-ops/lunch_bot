from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
import shutil
import asyncio
from functools import wraps
from config import WEEKDAYS, COMPANY_NAME, EXPORT_PATH

# Константы дедлайна
DEADLINE_DAY = 4  # Пятница
DEADLINE_HOUR = 16
DEADLINE_MINUTE = 0

# ==================== ДЕКОРАТОР RETRY ====================
EXPORT_PATH = "/data/exports"  # Важно! /data/exports
def retry(max_retries=3, delay=1):
    """Повтор при ошибках сети"""
    def decorator(func):
        @wraps(func)
        async def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return await func(*args, **kwargs)
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise
                    await asyncio.sleep(delay * (attempt + 1))
            return None
        return wrapper
    return decorator

# ==================== ДАТЫ И ДЕДЛАЙНЫ ====================

def get_target_week_dates():
    """
    ОПРЕДЕЛЯЕТ ЦЕЛЕВУЮ НЕДЕЛЮ ДЛЯ ЗАКАЗА
    До пятницы 16:00 = следующая неделя
    После пятницы 16:00 = через неделю
    """
    now = datetime.now()
    current_weekday = now.weekday()
    current_hour = now.hour
    current_minute = now.minute
    
    # Проверка дедлайна
    is_after_deadline = False
    
    if current_weekday > DEADLINE_DAY:
        is_after_deadline = True
    elif current_weekday == DEADLINE_DAY:
        if current_hour > DEADLINE_HOUR or (current_hour == DEADLINE_HOUR and current_minute >= DEADLINE_MINUTE):
            is_after_deadline = True
    
    # Рассчитываем целевой понедельник
    days_to_monday = (7 - current_weekday) % 7
    next_monday = now + timedelta(days=days_to_monday)
    
    if is_after_deadline:
        target_monday = next_monday + timedelta(days=7)
        week_type = "через неделю"
    else:
        target_monday = next_monday
        week_type = "следующую неделю"
    
    # Генерируем 7 дней
    dates = []
    for i in range(7):
        dates.append(target_monday + timedelta(days=i))
    
    return dates, week_type, is_after_deadline

def get_deadline_status():
    """Возвращает статус дедлайна"""
    now = datetime.now()
    weekday = now.weekday()
    hour = now.hour
    minute = now.minute
    
    if weekday < DEADLINE_DAY:
        days_left = DEADLINE_DAY - weekday
        if days_left == 0:
            hours_left = DEADLINE_HOUR - hour - 1
            minutes_left = 60 - minute
            return f"⏳ До дедлайна: {hours_left} ч {minutes_left} мин"
        else:
            return f"⏳ Дедлайн: пятница 16:00 (осталось {days_left} дн.)"
    elif weekday == DEADLINE_DAY:
        if hour < DEADLINE_HOUR:
            hours_left = DEADLINE_HOUR - hour - 1
            minutes_left = 60 - minute
            return f"⏳ Сегодня до 16:00 (осталось {hours_left} ч {minutes_left} мин)"
        else:
            return "🔓 Приём заказов на неделю через одну"
    else:
        return "🔓 Приём заказов на неделю через одну"

def format_date_for_db(date_obj):
    """Быстрое форматирование для БД"""
    return date_obj.strftime("%Y%m%d")

def format_date_for_display(date_obj):
    """Быстрое форматирование для показа"""
    return date_obj.strftime("%d.%m.%Y")

def get_week_range_display(dates):
    """Красивый вывод диапазона недели"""
    start = dates[0].strftime("%d.%m")
    end = dates[6].strftime("%d.%m.%Y")
    return f"{start} - {end}"

def get_progress_bar(current, total=7, size=10):
    """Визуальный прогресс-бар"""
    filled = int((current / total) * size)
    empty = size - filled
    return "🟦" * filled + "⬜" * empty

# ==================== EXCEL ОТЧЁТЫ ====================

def create_excel_report(all_orders, dates, save_copy=True):
    """Создаёт Excel файл со всеми заказами"""
    
    os.makedirs(EXPORT_PATH, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"заказы_{COMPANY_NAME}_{timestamp}.xlsx"
    temp_path = os.path.join(EXPORT_PATH, f"temp_{filename}")
    saved_path = os.path.join(EXPORT_PATH, filename)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Заказы"
    
    # Стили
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    center_alignment = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Заголовок
    ws.merge_cells('A1:I1')
    ws['A1'] = f"Заказы обедов • {COMPANY_NAME}"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = center_alignment
    
    # Период
    ws.merge_cells('A2:I2')
    start_date = format_date_for_display(dates[0])
    end_date = format_date_for_display(dates[6])
    ws['A2'] = f"Период: {start_date} - {end_date}"
    ws['A2'].font = Font(size=11)
    ws['A2'].alignment = center_alignment
    
    # Заголовки таблицы
    headers = ["№", "Сотрудник", "Инструктор"] + \
              [f"{WEEKDAYS[i]}\n{d.strftime('%d.%m')}" for i, d in enumerate(dates)] + \
              ["Всего"]
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border
    
    # Группировка заказов
    from collections import defaultdict
    employees = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    
    for _, full_name, _, instructor_name, date, quantity in all_orders:
        employees[full_name][instructor_name][date] = quantity
    
    # Заполнение данных
    row = 4
    emp_idx = 1
    
    for employee, instructors in sorted(employees.items()):
        first_row = True
        for instructor, orders in sorted(instructors.items()):
            ws.cell(row=row, column=1, value=emp_idx if first_row else "")
            ws.cell(row=row, column=2, value=employee)
            ws.cell(row=row, column=3, value=instructor)
            
            total = 0
            col = 4
            for date in dates:
                qty = orders.get(date.strftime("%Y%m%d"), 0)
                ws.cell(row=row, column=col, value=qty if qty > 0 else "-")
                total += qty
                col += 1
            
            ws.cell(row=row, column=col, value=total)
            row += 1
            first_row = False
        emp_idx += 1
        row += 1
    
    # Автоширина
    for col in range(1, 12):
        max_len = 10
        for r in range(1, row):
            val = ws.cell(row=r, column=col).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 25)
    
    wb.save(temp_path)
    
    if save_copy:
        shutil.copy2(temp_path, saved_path)
        return temp_path, saved_path
    
    return temp_path, None