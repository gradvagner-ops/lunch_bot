from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
import shutil
from config import WEEKDAYS, COMPANY_NAME, EXPORT_PATH

# Константы дедлайна
DEADLINE_DAY = 4  # Пятница
DEADLINE_HOUR = 16
DEADLINE_MINUTE = 0

# ==================== ДАТЫ И ДЕДЛАЙНЫ ====================

def get_target_week_dates():
    """Определяет целевую неделю для заказа"""
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
    """Возвращает статус дедлайна для отображения пользователю"""
    now = datetime.now()
    weekday = now.weekday()
    hour = now.hour
    minute = now.minute
    
    # До пятницы
    if weekday < DEADLINE_DAY:
        days_left = DEADLINE_DAY - weekday
        if days_left == 1:
            return f"⏳ Дедлайн: завтра до 16:00"
        else:
            return f"⏳ Дедлайн: пятница 16:00 (осталось {days_left} дн.)"
    
    # Пятница
    elif weekday == DEADLINE_DAY:
        if hour < DEADLINE_HOUR:
            hours_left = DEADLINE_HOUR - hour - 1
            minutes_left = 60 - minute
            return f"⏳ Сегодня до 16:00 (осталось {hours_left} ч {minutes_left} мин)"
        else:
            return "🔓 Приём заказов на неделю через одну"
    
    # Суббота-воскресенье
    else:
        return "🔓 Приём заказов на неделю через одну"


def format_date_for_db(date_obj):
    """Форматирование для БД"""
    return date_obj.strftime("%Y%m%d")

def format_date_for_display(date_obj):
    """Форматирование для показа"""
    return date_obj.strftime("%d.%m.%Y")

def get_week_range_display(dates):
    """Диапазон недели"""
    start = dates[0].strftime("%d.%m")
    end = dates[6].strftime("%d.%m.%Y")
    return f"{start} - {end}"

# ==================== EXCEL ОТЧЁТЫ ====================

def create_excel_report(all_orders, dates, save_copy=True):
    """Создаёт Excel файл со всеми заказами.
       Каждая неделя сохраняется на отдельном листе."""
    
    os.makedirs(EXPORT_PATH, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"заказы_архив_{timestamp}.xlsx"
    temp_path = os.path.join(EXPORT_PATH, f"temp_{filename}")
    saved_path = os.path.join(EXPORT_PATH, filename)
    
    # Если файл уже существует - открываем его, иначе создаём новый
    if os.path.exists(saved_path):
        wb = openpyxl.load_workbook(saved_path)
    else:
        wb = openpyxl.Workbook()
        # Удаляем дефолтный лист
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
    
    # Создаём название листа на основе периода
    week_start = dates[0].strftime("%d.%m")
    week_end = dates[6].strftime("%d.%m")
    sheet_name = f"Неделя {week_start}-{week_end}"
    
    # Если лист с таким названием уже есть - добавляем суффикс
    original_name = sheet_name
    counter = 1
    while sheet_name in wb.sheetnames:
        sheet_name = f"{original_name} ({counter})"
        counter += 1
    
    # Создаём новый лист
    ws = wb.create_sheet(title=sheet_name)
    
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
    
    # Заголовок с информацией о периоде
    ws.merge_cells('A1:I1')
    ws['A1'] = f"Заказы обедов • {COMPANY_NAME}"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = center_alignment
    
    # Период
    ws.merge_cells('A2:I2')
    start_date = dates[0].strftime("%d.%m.%Y")
    end_date = dates[6].strftime("%d.%m.%Y")
    ws['A2'] = f"Период: {start_date} - {end_date}"
    ws['A2'].font = Font(size=11)
    ws['A2'].alignment = center_alignment
    
    # Дата создания отчёта
    ws.merge_cells('A3:I3')
    creation_time = datetime.now().strftime('%d.%m.%Y %H:%M')
    ws['A3'] = f"Отчёт создан: {creation_time}"
    ws['A3'].font = Font(size=11)
    ws['A3'].alignment = center_alignment
    
    # Заголовки таблицы
    headers = ["№", "Сотрудник", "Инструктор"] + \
              [f"{WEEKDAYS[i]}\n{d.strftime('%d.%m')}" for i, d in enumerate(dates)] + \
              ["Всего"]
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border
    
    # Группируем заказы по сотрудникам и инструкторам
    from collections import defaultdict
    employees = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    
    for order in all_orders:
        if len(order) == 5:
            user_id, full_name, instructor_name, date, quantity = order
        else:
            print(f"⚠️ Неправильный формат данных: {order}")
            continue
        
        employees[full_name][instructor_name][date] = quantity
    
    # Заполнение данных
    row = 5
    emp_idx = 1
    
    for employee, instructors in sorted(employees.items()):
        first_row = True
        for instructor, orders in sorted(instructors.items()):
            # Номер сотрудника (только для первой строки)
            if first_row:
                ws.cell(row=row, column=1, value=emp_idx)
                first_row = False
            else:
                ws.cell(row=row, column=1, value="")
            
            # ФИО сотрудника
            ws.cell(row=row, column=2, value=employee)
            
            # Инструктор
            ws.cell(row=row, column=3, value=instructor)
            
            # Заполняем дни недели
            total = 0
            for col, date in enumerate(dates, start=4):
                date_key = date.strftime("%Y%m%d")
                qty = orders.get(date_key, 0)
                ws.cell(row=row, column=col, value=qty if qty > 0 else "-")
                total += qty
            
            # Итого по строке
            ws.cell(row=row, column=11, value=total)
            
            row += 1
            if first_row:
                emp_idx += 1
        
        # Пустая строка между сотрудниками
        row += 1
    
    # Итоговая строка
    if row > 5:  # Если есть данные
        total_row = row
        ws.cell(row=total_row, column=2, value="ИТОГО:")
        ws.cell(row=total_row, column=2).font = Font(bold=True)
        
        # Подсчёт итогов по дням
        for col in range(4, 11):
            col_total = 0
            for r in range(5, total_row):
                val = ws.cell(row=r, column=col).value
                if isinstance(val, (int, float)):
                    col_total += val
            ws.cell(row=total_row, column=col, value=col_total)
            ws.cell(row=total_row, column=col).font = Font(bold=True)
        
        # Общий итог
        total_all = 0
        for r in range(5, total_row):
            val = ws.cell(row=r, column=11).value
            if isinstance(val, (int, float)):
                total_all += val
        ws.cell(row=total_row, column=11, value=total_all)
        ws.cell(row=total_row, column=11).font = Font(bold=True)
    
    # Автоширина колонок
    for col in range(1, 12):
        max_len = 10
        for r in range(1, row + 1):
            val = ws.cell(row=r, column=col).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(max_len + 2, 25)
    
    # Сохраняем файл
    wb.save(temp_path)
    
    if save_copy:
        # Копируем в постоянное место
        import shutil
        shutil.copy2(temp_path, saved_path)
        print(f"📁 Excel файл сохранён: {saved_path} с листом '{sheet_name}'")
        return temp_path, saved_path
    
    return temp_path, None