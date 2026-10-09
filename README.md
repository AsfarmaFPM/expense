# 💰 Telegram Харажатлар ва Расходларни Ҳисоблаш Боти

Ушбу бот кунлик харажатларингизни категориялар бўйича ҳисоблаб боради, статистика ва ҳисоботларни чиқаради ҳамда исталган давр (кунлик, 1 ҳафталик, 30 кунлик, сана оралиғи) учун **Excel (.xlsx)** жадвал файли кўринишида юклаб олиш имконини беради.

---

## 🌐 Онлайн Серверга (Hosting / VPS) Жойлаштириш Қўлланмаси

Бот онлайн серверларда **24/7 узлуксиз** ишлаши учун тўлиқ тайёрлаб чиқилди. Сиз қуйидаги 3 хил усулдан бирини танлашингиз мумкин:

---

### 1-УСУЛ: Linux VPS Серверда (Ubuntu / Debian) — Энг яхши ва тавсия этиладиган усул

(Сервер провайдерлари: *Timeweb, Hetzner, Beget, DigitalOcean, AWS, Selectel* ва ҳ.к.)

1. **Серверга киринг (SSH орқали):**
   ```bash
   ssh root@сервер_ip_манзили
   ```

2. **Лойиҳа файлларини серверга юкланг:**
   - GitHub орқали (`git clone ...`) ёки `scp` / FileZilla орқали серверга ташланг.
   - Масалан, папкага киринг:
     ```bash
     cd expense_tracker_bot
     ```

3. **Токенни киритинг:**
   `.env` файлини яратинг ёки очинг:
   ```bash
   nano .env
   ```
   Ва ичига токенингизни ёзинг:
   ```env
   BOT_TOKEN=7123456789:AAFx...
   ```
   *(Сақлаш учун `Ctrl + O`, чиқиш учун `Ctrl + X`)*

4. **Автоматик ўрнатиш скриптини ишга туширинг:**
   Лойиҳа ичида тайёр [deploy.sh](file:///D:/Программы/awesome-free-llm-apis/expense_tracker_bot/deploy.sh) скрипти бор:
   ```bash
   chmod +x deploy.sh
   ./deploy.sh
   ```
   *Ушбу скрипт автоматик тарзда Python ва кутубхоналарни ўрнатади ҳамда ботни `systemd` сервисига қўшиб, 24/7 режимда ишга туширади!*

5. **Серверни бошқариш буйруқлари:**
   - Бот ишлаётганини кўриш: `sudo systemctl status expense-bot`
   - Жонли логларни кузатиш: `sudo journalctl -u expense-bot -f`
   - Ботни тўхтатиш: `sudo systemctl stop expense-bot`
   - Қайта ишга тушириш: `sudo systemctl restart expense-bot`

---

### 2-УСУЛ: Docker ва Docker Compose Орқали

Агар серверингизда Docker ўрнатилган бўлса:

1. `.env` файлига бот токенини киритинг:
   ```env
   BOT_TOKEN=7123456789:AAFx...
   ```

2. Ботни 1 та буйруқ билан контейнерда ишга туширинг:
   ```bash
   docker compose up -d --build
   ```

3. Логларни кўриш:
   ```bash
   docker compose logs -f
   ```

4. Тўхтатиш:
   ```bash
   docker compose down
   ```
*Маълумотлар базаси `./data` папкасида сақланади ва контейнер ўчса ҳам маълумотлар йўқолмайди.*

---

### 3-УСУЛ: Render.com ёки Railway.app (Бепул / Булутли платформалар)

Агар шахсий VPS серверингиз бўлмаса, бепул булутли платформалардан фойдаланишингиз мумкин:

1. Лойиҳани GitHub репозиторийингизга юкланг (`git push`).
2. [Render.com](https://render.com) ёки [Railway.app](https://railway.app) сайтида рўйхатдан ўтинг.
3. **New Web Service / Background Worker** тугмасини босиб, ўз GitHub репозиторийингизни танланг.
4. **Environment Variables** (Муҳит ўзгарувчилари) бўлимига:
   - `BOT_TOKEN` = `Сизнинг_Telegram_Токенингиз`
5. Лойиҳадаги [Procfile](file:///D:/Программы/awesome-free-llm-apis/expense_tracker_bot/Procfile) автоматик равишда ботни фонида ишга туширади!

---

## 💻 Компьютерда (Windows) Синовдан Ўтказиш

1. [config.py](file:///D:/Программы/awesome-free-llm-apis/expense_tracker_bot/config.py) ёки `.env` файлига токенни ёзинг.
2. [run.bat](file:///D:/Программы/awesome-free-llm-apis/expense_tracker_bot/run.bat) файлини 2 марта босиб ишга туширинг.

---

## 📂 Файллар тузилмаси

```
expense_tracker_bot/
├── bot.py                # Асосий Telegram бот логикаси
├── database.py           # SQLite база модули (харажатлар, категориялар)
├── excel_export.py       # Openpyxl орқали Excel генерацияси
├── config.py             # Созламалар ва DATA_DIR мослашувчанлиги
├── requirements.txt      # Зарур кутубхоналар
├── run.bat               # Windows учун 1-кликла ишга туширувчи файл
├── Dockerfile            # Docker контейнер файли
├── docker-compose.yml    # Docker Compose конфигурацияси
├── .dockerignore         # Ортиқча файлларни контейнерга қўшмаслик рўйхати
├── deploy.sh             # Linux (Ubuntu) учун авто-ўрнатиш ва ишга тушириш
├── expense-bot.service   # Systemd 24/7 авто-ишга тушириш сервиси
├── Procfile              # Render / Railway булут сервислари учун
└── .env.example          # Муҳит ўзгарувчилари намунаси
```
