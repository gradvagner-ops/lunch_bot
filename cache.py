from functools import lru_cache
from datetime import datetime, timedelta
import pickle
import os
from config import WEEKDAYS

class Cache:
    """Централизованное кэширование для ускорения работы"""
    
    def __init__(self, cache_file=".cache"):
        self.cache_file = cache_file
        self._cache = {}
        self.load()
    
    @lru_cache(maxsize=32)
    def get_week_dates(self, week_offset=0):
        """
        Кэшированные даты недели
        week_offset: 0 - следующая неделя, 1 - через неделю
        """
        now = datetime.now()
        days_to_monday = (7 - now.weekday()) % 7
        target_monday = now + timedelta(days=days_to_monday + week_offset * 7)
        return tuple(target_monday + timedelta(days=i) for i in range(7))
    
    @lru_cache(maxsize=128)
    def parse_date(self, date_str):
        """Кэшированный парсинг даты"""
        return datetime.strptime(date_str, "%Y%m%d")
    
    @lru_cache(maxsize=128)
    def format_date_display(self, date_str):
        """Кэшированное форматирование даты"""
        date_obj = self.parse_date(date_str)
        return date_obj.strftime("%d.%m.%Y")
    
    @lru_cache(maxsize=128)
    def format_date_short(self, date_str):
        """Короткий формат даты"""
        date_obj = self.parse_date(date_str)
        return date_obj.strftime("%d.%m")
    
    @lru_cache(maxsize=128)
    def get_day_name(self, date_str):
        """Название дня недели"""
        date_obj = self.parse_date(date_str)
        return WEEKDAYS[date_obj.weekday()]
    
    @lru_cache(maxsize=128)
    def get_day_short(self, date_str):
        """Короткое название дня"""
        date_obj = self.parse_date(date_str)
        return date_obj.strftime("%a")
    
    def precalculate_week_dates(self, date_keys):
        """
        Предрасчет всех форматов дат для недели
        date_keys: список строк дат в формате YYYYMMDD
        """
        week_data = []
        for date_key in date_keys:
            week_data.append({
                'key': date_key,
                'obj': self.parse_date(date_key),
                'day_name': self.get_day_name(date_key),
                'display': self.format_date_display(date_key),
                'short': self.format_date_short(date_key),
                'day_short': self.get_day_short(date_key)
            })
        return week_data
    
    def clear_cache(self):
        """Очистка кэша"""
        self.get_week_dates.cache_clear()
        self.parse_date.cache_clear()
        self.format_date_display.cache_clear()
        self.format_date_short.cache_clear()
        self.get_day_name.cache_clear()
        self.get_day_short.cache_clear()
    
    def save(self):
        """Сохранить кэш на диск"""
        with open(self.cache_file, 'wb') as f:
            pickle.dump(self._cache, f)
    
    def load(self):
        """Загрузить кэш с диска"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'rb') as f:
                    self._cache = pickle.load(f)
            except:
                pass

cache = Cache()