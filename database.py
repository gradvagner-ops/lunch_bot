import sqlite3
import os
from datetime import datetime

DB_PATH = "/data/orders.db"  # Для Amvera

class Database:
    def __init__(self, db_file=DB_PATH):
        # Создаем папку /data если её нет
        os.makedirs("/data", exist_ok=True)
        self.db_file = db_file
        self.init_db()
    
    def init_db(self):
        with sqlite3.connect(self.db_file) as conn:
            # Оптимизация для SQLite
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS employees (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    full_name TEXT,
                    first_registration DATE
                )
            ''')
            
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
            
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_orders_user ON orders(user_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_orders_date ON orders(date)')
            conn.commit()
    
    def save_order(self, user_id, instructor_name, date, quantity):
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO orders 
                (user_id, instructor_name, date, quantity)
                VALUES (?, ?, ?, ?)
            ''', (user_id, instructor_name, date, quantity))
            conn.commit()
    
    def get_user_orders(self, user_id):
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
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT o.user_id, e.full_name, o.instructor_name, o.date, o.quantity
                FROM orders o
                LEFT JOIN employees e ON e.user_id = o.user_id
                WHERE o.quantity > 0
                ORDER BY o.date DESC
            ''')
            return cursor.fetchall()
    
    def register_employee(self, user_id, username, full_name):
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR IGNORE INTO employees (user_id, username, full_name, first_registration)
                VALUES (?, ?, ?, DATE('now'))
            ''', (user_id, username, full_name))
            conn.commit()