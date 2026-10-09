#!/bin/bash
# ========================================================
# Expense Tracker Bot - Ubuntu/Debian Avtomatik O'rnatish
# ========================================================

set -e

echo "🚀 Tizim paketlari yangilanmoqda..."
sudo apt update && sudo apt upgrade -y

echo "🐍 Python va zarur vositalar o'rnatilmoqda..."
sudo apt install -y python3 python3-pip python3-venv git

PROJECT_DIR=$(pwd)

echo "📦 Virtual muhit (venv) yaratilmoqda..."
python3 -m venv venv
source venv/bin/activate

echo "📥 Kerakli kutubxonalar o'rnatilmoqda..."
pip install --upgrade pip
pip install -r requirements.txt

# .env tekshirish
if [ ! -f .env ]; then
    echo "⚠️  .env fayli topilmadi. .env.example dan nusxa olinmoqda..."
    cp .env.example .env
    echo "❗ DIQQAT: .env fayliga BOT_TOKEN ni yozishni unutmang!"
fi

echo "⚙️ Systemd servisi sozlanmoqda..."
SERVICE_FILE="/etc/systemd/system/expense-bot.service"

sudo bash -c "cat > $SERVICE_FILE" <<EOL
[Unit]
Description=Expense Tracker Telegram Bot
After=network.target

[Service]
Type=simple
User=$USER
WorkingDirectory=$PROJECT_DIR
ExecStart=$PROJECT_DIR/venv/bin/python bot.py
Restart=always
RestartSec=10
EnvironmentFile=$PROJECT_DIR/.env

[Install]
WantedBy=multi-user.target
EOL

echo "🔄 Servis ishga tushirilmoqda..."
sudo systemctl daemon-reload
sudo systemctl enable expense-bot
sudo systemctl restart expense-bot

echo ""
echo "========================================================"
echo "✅ Bot muvaffaqiyatli o'rnatildi va fonda (24/7) ishlamoqda!"
echo "Holatni tekshirish uchun: sudo systemctl status expense-bot"
echo "Loglarni ko'rish uchun:   sudo journalctl -u expense-bot -f"
echo "Botni to'xtatish uchun:   sudo systemctl stop expense-bot"
echo "Botni qayta ishga solish: sudo systemctl restart expense-bot"
echo "========================================================"
