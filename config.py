"""
Configuration module for Expense Tracker Telegram Bot.
Supports environment variables, .env file, and configurable data directory.
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Data directory (for database and generated reports)
DATA_DIR = os.getenv("DATA_DIR", BASE_DIR)

# Telegram Bot Token
BOT_TOKEN = os.getenv("BOT_TOKEN", "SIZNING_BOT_TOKENINGIZ")

# Read from .env if present
ENV_FILE = os.path.join(BASE_DIR, ".env")
if os.path.exists(ENV_FILE):
    with open(ENV_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("BOT_TOKEN="):
                token_val = line.split("=", 1)[1].strip().strip('"').strip("'")
                if token_val:
                    BOT_TOKEN = token_val
            elif line.startswith("DATA_DIR="):
                dir_val = line.split("=", 1)[1].strip().strip('"').strip("'")
                if dir_val:
                    DATA_DIR = dir_val

os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "expenses.db")
REPORTS_DIR = os.path.join(DATA_DIR, "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)
