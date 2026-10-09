"""
Database management module for Expense Tracker Telegram Bot.
SQLite is used for storing expenses and user custom categories.
"""

import sqlite3
from datetime import datetime, timedelta
import os
from typing import List, Dict, Any, Tuple, Optional
from config import DB_PATH

DEFAULT_CATEGORIES = [
    "⛽️ Бензин ва Транспорт",
    "🍔 Озиқ-овқат ва Бозор",
    "🏠 Коммунал ва Рўзғор",
    "🛍 Кийим ва Харидлар",
    "☕️ Кафе ва Ресторан",
    "💊 Дорихона ва Саломатлик",
    "📱 Телефон ва Интернет",
    "🎓 Таълим ва Китоблар",
    "📦 Бошқа харажатлар"
]


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Initializes tables in database."""
    with get_connection() as conn:
        cursor = conn.cursor()
        
        # Expenses table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                category TEXT NOT NULL,
                amount REAL NOT NULL,
                description TEXT,
                created_at TEXT NOT NULL
            )
        """)
        
        # Custom user categories table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS custom_categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        
        # Index on user_id and created_at for fast lookups
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_expenses_user_date 
            ON expenses (user_id, created_at)
        """)
        
        conn.commit()


def add_expense(user_id: int, category: str, amount: float, description: Optional[str] = None, dt: Optional[datetime] = None) -> int:
    """Adds a new expense for a user."""
    if dt is None:
        dt = datetime.now()
    created_at = dt.strftime("%Y-%m-%d %H:%M:%S")
    desc = (description or "").strip()
    
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO expenses (user_id, category, amount, description, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (user_id, category, amount, desc, created_at))
        conn.commit()
        return cursor.lastrowid


def delete_last_expense(user_id: int) -> Optional[Dict[str, Any]]:
    """Deletes the most recently added expense for a user."""
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM expenses
            WHERE user_id = ?
            ORDER BY id DESC
            LIMIT 1
        """, (user_id,))
        row = cursor.fetchone()
        if not row:
            return None
        
        last_item = dict(row)
        cursor.execute("DELETE FROM expenses WHERE id = ?", (last_item["id"],))
        conn.commit()
        return last_item


def get_expenses_by_period(user_id: int, start_date: str, end_date: str) -> List[Dict[str, Any]]:
    """
    Returns all expenses between start_date (inclusive, YYYY-MM-DD 00:00:00) 
    and end_date (inclusive, YYYY-MM-DD 23:59:59).
    """
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, category, amount, description, created_at
            FROM expenses
            WHERE user_id = ? AND created_at >= ? AND created_at <= ?
            ORDER BY created_at ASC
        """, (user_id, start_date, end_date))
        rows = cursor.fetchall()
        return [dict(r) for r in rows]


def get_summary_by_period(user_id: int, start_date: str, end_date: str) -> Dict[str, Any]:
    """
    Calculates summary totals and category breakdown.
    """
    with get_connection() as conn:
        cursor = conn.cursor()
        # Category breakdown
        cursor.execute("""
            SELECT category, SUM(amount) as total_amount, COUNT(id) as count
            FROM expenses
            WHERE user_id = ? AND created_at >= ? AND created_at <= ?
            GROUP BY category
            ORDER BY total_amount DESC
        """, (user_id, start_date, end_date))
        categories = [dict(r) for r in cursor.fetchall()]
        
        # Overall total
        cursor.execute("""
            SELECT COALESCE(SUM(amount), 0) as grand_total, COUNT(id) as total_count
            FROM expenses
            WHERE user_id = ? AND created_at >= ? AND created_at <= ?
        """, (user_id, start_date, end_date))
        tot_row = cursor.fetchone()
        
        return {
            "grand_total": tot_row["grand_total"] if tot_row else 0,
            "total_count": tot_row["total_count"] if tot_row else 0,
            "categories": categories
        }


def get_all_categories(user_id: int) -> List[str]:
    """Returns default categories plus user's custom categories."""
    cats = list(DEFAULT_CATEGORIES)
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT name FROM custom_categories
            WHERE user_id = ?
            ORDER BY id ASC
        """, (user_id,))
        customs = [r["name"] for r in cursor.fetchall()]
        for c in customs:
            if c not in cats:
                cats.append(c)
    return cats


def add_custom_category(user_id: int, name: str) -> bool:
    """Adds a new custom category for user."""
    clean_name = name.strip()
    if not clean_name:
        return False
    with get_connection() as conn:
        cursor = conn.cursor()
        # check duplicate
        cursor.execute("SELECT id FROM custom_categories WHERE user_id = ? AND name = ?", (user_id, clean_name))
        if cursor.fetchone() or clean_name in DEFAULT_CATEGORIES:
            return False
        cursor.execute("""
            INSERT INTO custom_categories (user_id, name, created_at)
            VALUES (?, ?, ?)
        """, (user_id, clean_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        return True
