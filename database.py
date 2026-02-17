import sqlite3
import os
from datetime import datetime

class Database:
    def __init__(self, db_file="orders.db"):
        self.db_file = db_file
        self.init_db()
    
    def init_db(self):
        """Создаём таблицы, если их нет"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            
            # Таблица сотрудников
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS employees (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    full_name TEXT,
                    first_registration DATE
                )
            ''')
            
            # Таблица заказов
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS orders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER,
                    instructor_name TEXT,
                    date TEXT,
                    quantity INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Индексы для скорости
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_orders_user ON orders(user_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_orders_date ON orders(date)')
            conn.commit()
    
    def register_employee(self, user_id, username, full_name):
        """Регистрируем сотрудника"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR IGNORE INTO employees (user_id, username, full_name, first_registration)
                VALUES (?, ?, ?, DATE('now'))
            ''', (user_id, username, full_name))
            conn.commit()
    
    def save_order(self, user_id, instructor_name, date, quantity):
        """Сохраняем заказ"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            
            # Сначала удаляем старый заказ на эту дату (если был)
            cursor.execute('''
                DELETE FROM orders 
                WHERE user_id = ? AND instructor_name = ? AND date = ?
            ''', (user_id, instructor_name, date))
            
            # Вставляем новый, если количество > 0
            if quantity > 0:
                cursor.execute('''
                    INSERT INTO orders (user_id, instructor_name, date, quantity)
                    VALUES (?, ?, ?, ?)
                ''', (user_id, instructor_name, date, quantity))
            
            conn.commit()
            return True
    
    def get_user_orders(self, user_id):
        """Получаем заказы сотрудника"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT instructor_name, date, quantity 
                FROM orders
                WHERE user_id = ? AND quantity > 0
                ORDER BY date DESC
            ''', (user_id,))
            return cursor.fetchall()
    
    def get_all_orders(self):
        """Получаем все заказы для Excel"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT 
                    o.user_id,
                    COALESCE(e.full_name, 'Неизвестно') as full_name,
                    o.instructor_name,
                    o.date,
                    o.quantity
                FROM orders o
                LEFT JOIN employees e ON e.user_id = o.user_id
                WHERE o.quantity > 0
                ORDER BY o.date DESC, o.instructor_name
            ''')
            result = cursor.fetchall()
            print(f"📤 get_all_orders вернул {len(result)} записей")
            if result:
                print(f"   Пример: {result[0]}")
            return result
    
    def delete_user_orders(self, user_id):
        """Удаляем все заказы сотрудника"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM orders WHERE user_id = ?', (user_id,))
            conn.commit()
            return cursor.rowcount > 0
    
    def get_employee_name(self, user_id):
        """Получаем имя сотрудника по ID"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT full_name FROM employees WHERE user_id = ?', (user_id,))
            result = cursor.fetchone()
            return result[0] if result else None
    
    def get_orders_count(self):
        """Сколько всего заказов в БД"""
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM orders WHERE quantity > 0')
            return cursor.fetchone()[0]