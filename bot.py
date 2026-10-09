"""
Telegram Bot for tracking daily expenses with Excel report generation.
Built with python-telegram-bot (v22+) and openpyxl.
"""

import logging
import re
import os
import sys
from datetime import datetime, timedelta
from typing import Optional, Tuple

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from telegram import (
    Update,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    ReplyKeyboardRemove
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters
)

import config
import database
import excel_export

# Logging setup
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# --- Reply Keyboards ---
def get_main_keyboard():
    keyboard = [
        [KeyboardButton("➕ Харажат қўшиш"), KeyboardButton("📊 Ҳисоботлар (Отчет)")],
        [KeyboardButton("📥 Excel юклаб олиш"), KeyboardButton("🏷 Категориялар")],
        [KeyboardButton("↩️ Охиргисини ўчириш"), KeyboardButton("ℹ️ Ёрдам")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


# --- Inline Keyboards ---
def get_categories_keyboard(user_id: int):
    cats = database.get_all_categories(user_id)
    buttons = []
    # 2 columns
    row = []
    for cat in cats:
        row.append(InlineKeyboardButton(cat, callback_data=f"selcat:{cat}"))
        if len(row) == 2:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    
    # Extra options
    buttons.append([
        InlineKeyboardButton("➕ Янги категория қўшиш", callback_data="add_custom_cat")
    ])
    buttons.append([
        InlineKeyboardButton("❌ Бекор қилиш", callback_data="cancel_action")
    ])
    return InlineKeyboardMarkup(buttons)


def get_period_keyboard(action_prefix="rep"):
    keyboard = [
        [
            InlineKeyboardButton("📅 Бугунги (Кунлик)", callback_data=f"{action_prefix}:today"),
            InlineKeyboardButton("🗓 1 ҳафталик", callback_data=f"{action_prefix}:week")
        ],
        [
            InlineKeyboardButton("🗓 30 кунлик", callback_data=f"{action_prefix}:month"),
            InlineKeyboardButton("📊 Барча вақт", callback_data=f"{action_prefix}:all")
        ],
        [
            InlineKeyboardButton("📆 Сана оралиғини киритиш", callback_data=f"{action_prefix}:custom")
        ],
        [
            InlineKeyboardButton("❌ Ёпиш", callback_data="cancel_action")
        ]
    ]
    return InlineKeyboardMarkup(keyboard)


# --- Helper Parsing Functions ---
def parse_amount_and_note(text: str) -> Tuple[Optional[float], str]:
    """
    Parses amount and description from user text.
    Handles:
      - '300000'
      - '300 000'
      - '300000 бензин учун'
      - '300 минг бензин'
      - '300k taxi'
      - '50.000 тушлик'
    """
    text = text.strip()
    m = re.match(r"^([\d\s\.,]+)\s*(k|к|ming|минг)?\s*(so[\'\`\ʼ]?m|som|сум|сўм)?\s*(.*)$", text, re.IGNORECASE)
    if not m:
        return None, ""
    
    raw_num = m.group(1).replace(" ", "").replace(",", ".")
    multiplier_str = m.group(2)
    note = m.group(4).strip() if m.group(4) else ""
    
    # Handle dot separators (e.g. 300.000)
    if raw_num.count(".") > 1:
        raw_num = raw_num.replace(".", "")
    elif raw_num.count(".") == 1:
        parts = raw_num.split(".")
        if len(parts[1]) == 3 and parts[1].isdigit():
            raw_num = parts[0] + parts[1]

    try:
        val = float(raw_num)
    except ValueError:
        return None, ""

    if val <= 0:
        return None, ""

    if multiplier_str and multiplier_str.lower() in ["k", "к", "ming", "минг"]:
        val *= 1000

    return val, note


def parse_date_range(text: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Parses custom date ranges from user message.
    Returns (start_date_str, end_date_str, label)
    """
    text = text.strip()
    raw_dates = re.findall(r"\b(?:\d{4}[-\/\.]\d{1,2}[-\/\.]\d{1,2}|\d{1,2}[-\/\.]\d{1,2}(?:[-\/\.]\d{2,4})?)\b", text)
    curr_year = datetime.now().year
    dates = []
    
    for tok in raw_dates:
        parsed = None
        for fmt in ["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y.%m.%d"]:
            try:
                parsed = datetime.strptime(tok, fmt)
                break
            except ValueError:
                pass
        if not parsed:
            for fmt in ["%d.%m", "%d/%m"]:
                try:
                    dt = datetime.strptime(tok, fmt)
                    parsed = dt.replace(year=curr_year)
                    break
                except ValueError:
                    pass
        if parsed:
            dates.append(parsed)

    if len(dates) == 1:
        return dates[0].strftime("%Y-%m-%d 00:00:00"), dates[0].strftime("%Y-%m-%d 23:59:59"), dates[0].strftime("%d.%m.%Y")
    elif len(dates) >= 2:
        d1, d2 = sorted([dates[0], dates[1]])
        label = f"{d1.strftime('%d.%m.%Y')} — {d2.strftime('%d.%m.%Y')}"
        return d1.strftime("%Y-%m-%d 00:00:00"), d2.strftime("%Y-%m-%d 23:59:59"), label
    
    return None, None, None


def get_dates_for_period(period_code: str) -> Tuple[str, str, str]:
    now = datetime.now()
    if period_code == "today":
        start = now.strftime("%Y-%m-%d 00:00:00")
        end = now.strftime("%Y-%m-%d 23:59:59")
        label = f"Бугун ({now.strftime('%d.%m.%Y')})"
    elif period_code == "week":
        start_dt = now - timedelta(days=6)
        start = start_dt.strftime("%Y-%m-%d 00:00:00")
        end = now.strftime("%Y-%m-%d 23:59:59")
        label = f"Сўнгги 7 кун ({start_dt.strftime('%d.%m.%Y')} — {now.strftime('%d.%m.%Y')})"
    elif period_code == "month":
        start_dt = now - timedelta(days=29)
        start = start_dt.strftime("%Y-%m-%d 00:00:00")
        end = now.strftime("%Y-%m-%d 23:59:59")
        label = f"Сўнгги 30 кун ({start_dt.strftime('%d.%m.%Y')} — {now.strftime('%d.%m.%Y')})"
    else:  # all
        start = "2000-01-01 00:00:00"
        end = "2099-12-31 23:59:59"
        label = "Барча вақт бўйича"
    return start, end, label


def format_sum(amount: float) -> str:
    """Formats amount as readable number e.g. 300 000."""
    return f"{amount:,.0f}".replace(",", " ")


# --- Command Handlers ---
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    context.user_data.clear()
    msg = (
        f"Ассалому алайкум, <b>{user.first_name}</b>! 👋\n\n"
        f"Бу бот орқали сиз кунлик харажатларингизни осонгина ҳисоблаб боришингиз, "
        f"категориялар бўйича таҳлил қилишингиз ва исталган вақтда <b>Excel (.xlsx)</b> кўринишида "
        f"ҳисоботларни юклаб олишингиз мумкин.\n\n"
        f"👇 Қуйидаги тугмалардан бирини танланг:"
    )
    await update.message.reply_text(msg, parse_mode="HTML", reply_markup=get_main_keyboard())


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "📖 <b>Ботдан фойдаланиш бўйича қўлланма:</b>\n\n"
        "1️⃣ <b>Харажат қўшиш:</b>\n"
        "• «➕ Харажат қўшиш» тугмасини босинг.\n"
        "• Категорияни танланг (масалан: ⛽️ Бензин).\n"
        "• Суммани ва изоҳни киритинг: <code>300000 бензин учун</code> ёки шунчаки <code>300000</code>.\n\n"
        "2️⃣ <b>Ҳисобот кўриш ва Excel юклаб олиш:</b>\n"
        "• «📊 Ҳисоботлар» ёки «📥 Excel юклаб олиш» тугмасини босинг.\n"
        "• Керакли даврни танланг (Кунлик, 1 ҳафталик, 30 кунлик ёки сана оралиғи).\n"
        "• Бот сизга батафсил жамланма ва чиройли Excel жадвал файлини тақдим этади!\n\n"
        "3️⃣ <b>Сана оралиғи бўйича ҳисобот:</b>\n"
        "• «Сана оралиғини киритиш»ни танлаб, масалан: <code>01.10.2024 - 09.10.2024</code> кўринишида сана юборинг.\n\n"
        "4️⃣ <b>Охирги харажатни бекор қилиш:</b>\n"
        "• Агар хато киритган бўлсангиз, «↩️ Охиргисини ўчириш» тугмасини босинг."
    )
    await update.message.reply_text(msg, parse_mode="HTML", reply_markup=get_main_keyboard())


# --- Menu Handlers ---
async def handle_add_expense_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    context.user_data["state"] = "SELECT_CATEGORY"
    await update.message.reply_text(
        "🏷 <b>Харажат категориясини танланг:</b>",
        parse_mode="HTML",
        reply_markup=get_categories_keyboard(user_id)
    )


async def handle_reports_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "📊 <b>Қайси давр учун ҳисобот кўрмоқчисиз?</b>",
        parse_mode="HTML",
        reply_markup=get_period_keyboard(action_prefix="rep")
    )


async def handle_excel_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "📥 <b>Қайси давр учун Excel файлини юклаб олмоқчисиз?</b>",
        parse_mode="HTML",
        reply_markup=get_period_keyboard(action_prefix="xls")
    )


async def handle_categories_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    cats = database.get_all_categories(user_id)
    cats_text = "\n".join([f"• {c}" for c in cats])
    msg = (
        f"🏷 <b>Мавжуд категориялар рўйхати:</b>\n\n"
        f"{cats_text}\n\n"
        f"Янги ўз категориянгизни қўшиш учун қуйидаги тугмани босинг:"
    )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("➕ Янги категория қўшиш", callback_data="add_custom_cat")],
        [InlineKeyboardButton("❌ Ёпиш", callback_data="cancel_action")]
    ])
    await update.message.reply_text(msg, parse_mode="HTML", reply_markup=keyboard)


async def handle_delete_last_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    deleted = database.delete_last_expense(user_id)
    if not deleted:
        await update.message.reply_text(
            "⚠️ Ўчириш учун харажатлар топилмади.",
            reply_markup=get_main_keyboard()
        )
        return

    amt_str = format_sum(deleted['amount'])
    desc_str = f" ({deleted['description']})" if deleted['description'] else ""
    await update.message.reply_text(
        f"🗑 <b>Охирги харажат ўчирилди:</b>\n"
        f"🏷 Категория: <b>{deleted['category']}</b>\n"
        f"💰 Сумма: <b>{amt_str} сўм</b>{desc_str}\n"
        f"🕒 Сана: <i>{deleted['created_at']}</i>",
        parse_mode="HTML",
        reply_markup=get_main_keyboard()
    )


# --- Callback Query Handlers ---
async def handle_callback_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = update.effective_user.id

    if data == "cancel_action":
        context.user_data.clear()
        try:
            await query.edit_message_text("❌ Амал бекор қилинди.")
        except Exception:
            pass
        return

    # Category selection for expense
    if data.startswith("selcat:"):
        cat_name = data.split("selcat:", 1)[1]
        temp_amount = context.user_data.get("temp_amount")
        temp_note = context.user_data.get("temp_note")

        if temp_amount is not None:
            # User already provided the amount and note earlier!
            database.add_expense(user_id, cat_name, temp_amount, temp_note)
            context.user_data.clear()

            today_start = datetime.now().strftime("%Y-%m-%d 00:00:00")
            today_end = datetime.now().strftime("%Y-%m-%d 23:59:59")
            today_sum = database.get_summary_by_period(user_id, today_start, today_end)["grand_total"]

            note_str = f"\n📝 <b>Изоҳ:</b> {temp_note}" if temp_note else ""
            time_str = datetime.now().strftime("%d.%m.%Y %H:%M")

            confirm_msg = (
                f"✅ <b>Харажат муваффақиятли сақланди!</b>\n\n"
                f"🏷 <b>Категория:</b> {cat_name}\n"
                f"💰 <b>Сумма:</b> <b>{format_sum(temp_amount)} сўм</b>{note_str}\n"
                f"🕒 <b>Вақт:</b> {time_str}\n\n"
                f"📈 <i>Бугунги жами харажатингиз:</i> <b>{format_sum(today_sum)} сўм</b>"
            )
            await query.edit_message_text(confirm_msg, parse_mode="HTML")
            return

        context.user_data["state"] = "WAITING_AMOUNT"
        context.user_data["selected_category"] = cat_name

        text = (
            f"🏷 <b>Танланган категория:</b> {cat_name}\n\n"
            f"💰 <b>Қилинган расход суммасини киритинг:</b>\n\n"
            f"<i>Масалан:</i>\n"
            f"• <code>300000</code>\n"
            f"• <code>300000 бензин учун</code>\n"
            f"• <code>300 минг АИ-95</code>\n\n"
            f"<i>(Сумма ёнига нима учун сарфланганини ҳам ёзиб кетсангиз, ҳисобот ва Excel'да кўринади)</i>"
        )
        await query.edit_message_text(text, parse_mode="HTML")
        return

    # Add custom category prompt
    if data == "add_custom_cat":
        context.user_data["state"] = "WAITING_NEW_CAT_NAME"
        text = (
            "➕ <b>Янги категория қўшиш:</b>\n\n"
            "Илтимос, янги категория номини ёзиб юборинг.\n"
            "<i>Масалан: 🎁 Совғалар ёки 🚗 Автосервис</i>"
        )
        await query.edit_message_text(text, parse_mode="HTML")
        return

    # Report viewing callbacks
    if data.startswith("rep:"):
        period = data.split("rep:", 1)[1]
        if period == "custom":
            context.user_data["state"] = "WAITING_CUSTOM_DATE_FOR_REP"
            await query.edit_message_text(
                "📆 <b>Сана оралиғини киритинг:</b>\n\n"
                "<i>Масалан:</i>\n"
                "• <code>01.10.2024 - 09.10.2024</code>\n"
                "• <code>01.10.2024 дан 09.10.2024 гача</code>\n"
                "• <code>05.10.2024</code> (битта кун учун)",
                parse_mode="HTML"
            )
            return

        start_date, end_date, label = get_dates_for_period(period)
        await send_text_report(query, user_id, start_date, end_date, label, is_edit=True)
        return

    # Excel direct export callbacks
    if data.startswith("xls:"):
        period = data.split("xls:", 1)[1]
        if period == "custom":
            context.user_data["state"] = "WAITING_CUSTOM_DATE_FOR_XLS"
            await query.edit_message_text(
                "📆 <b>Excel учун сана оралиғини киритинг:</b>\n\n"
                "<i>Масалан:</i>\n"
                "• <code>01.10.2024 - 09.10.2024</code>\n"
                "• <code>2024-10-01 - 2024-10-09</code>",
                parse_mode="HTML"
            )
            return

        start_date, end_date, label = get_dates_for_period(period)
        await send_excel_file(context, update.effective_chat.id, user_id, start_date, end_date, label)
        return

    # Excel download triggered from a report view: dlxls:start|end|label
    if data.startswith("dlxls:"):
        payload = data.split("dlxls:", 1)[1]
        try:
            start_date, end_date, label = payload.split("|", 2)
            await send_excel_file(context, update.effective_chat.id, user_id, start_date, end_date, label)
        except Exception as e:
            logger.error(f"Error handling dlxls callback: {e}")
            await query.message.reply_text("❌ Файлни юклашда хатолик юз берди.")
        return


# --- Text Report & Excel Sending Functions ---
async def send_text_report(target, user_id: int, start_date: str, end_date: str, label: str, is_edit: bool = False):
    summary = database.get_summary_by_period(user_id, start_date, end_date)
    expenses = database.get_expenses_by_period(user_id, start_date, end_date)

    grand_total = summary["grand_total"]
    total_count = summary["total_count"]
    cats = summary["categories"]

    if total_count == 0:
        text = (
            f"📊 <b>Ҳисобот: {label}</b>\n\n"
            f"ℹ️ Ушбу давр учун ҳеч қандай харажат киритилмаган."
        )
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("🔙 Ҳисоботлар менюси", callback_data="cancel_action")]
        ])
    else:
        cats_lines = []
        for c in cats:
            c_name = c["category"]
            c_sum = c["total_amount"]
            percent = (c_sum / grand_total * 100) if grand_total > 0 else 0
            cats_lines.append(f"• <b>{c_name}:</b> {format_sum(c_sum)} сўм ({percent:.1f}%)")
        cats_formatted = "\n".join(cats_lines)

        # Recent 5 entries
        recent = expenses[-5:]
        recent_lines = []
        for r in reversed(recent):
            time_str = r["created_at"][11:16]
            desc_part = f" — <i>{r['description']}</i>" if r["description"] else ""
            recent_lines.append(f"⏱ {time_str} | {r['category']}: <b>{format_sum(r['amount'])}</b> сўм{desc_part}")
        recent_formatted = "\n".join(recent_lines)

        text = (
            f"📊 <b>ҲИСОБОТ: {label}</b>\n\n"
            f"💰 <b>Жами харажат:</b> <code>{format_sum(grand_total)}</code> сўм\n"
            f"🧾 <b>Харажатлар сони:</b> {total_count} та\n\n"
            f"📂 <b>Категориялар бўйича тақсимланиши:</b>\n"
            f"{cats_formatted}\n\n"
            f"🕒 <b>Сўнгги ёзувлар:</b>\n"
            f"{recent_formatted}"
        )

        dl_data = f"dlxls:{start_date}|{end_date}|{label}"
        # Keep dl_data within 64 bytes limit if possible
        if len(dl_data.encode("utf-8")) > 64:
            dl_data = f"dlxls:{start_date[:10]}|{end_date[:10]}|Ҳисобот"

        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📥 Ушбу ҳисоботни Excel'да юклаб олиш", callback_data=dl_data)],
            [InlineKeyboardButton("🔙 Бошқа давр ҳисоботи", callback_data="cancel_action")]
        ])

    if is_edit:
        try:
            await target.edit_message_text(text, parse_mode="HTML", reply_markup=keyboard)
            return
        except Exception:
            pass
    await target.message.reply_text(text, parse_mode="HTML", reply_markup=keyboard)


async def send_excel_file(context: ContextTypes.DEFAULT_TYPE, chat_id: int, user_id: int, start_date: str, end_date: str, label: str):
    expenses = database.get_expenses_by_period(user_id, start_date, end_date)
    summary = database.get_summary_by_period(user_id, start_date, end_date)
    grand_total = summary["grand_total"]

    excel_path = excel_export.generate_excel_report(expenses, label, start_date, end_date, user_id)

    caption = (
        f"📊 <b>Харажатлар Excel ҳисоботи</b>\n\n"
        f"🗓 <b>Давр:</b> {label}\n"
        f"💰 <b>Жами сарфланган:</b> <b>{format_sum(grand_total)} сўм</b>\n"
        f"🧾 <b>Ёзувлар сони:</b> {len(expenses)} та"
    )

    try:
        with open(excel_path, "rb") as f:
            await context.bot.send_document(
                chat_id=chat_id,
                document=f,
                filename=os.path.basename(excel_path),
                caption=caption,
                parse_mode="HTML"
            )
    except Exception as e:
        logger.error(f"Error sending excel file: {e}")
        await context.bot.send_message(chat_id=chat_id, text=f"❌ Файлни юборишда хатолик: {e}")
    finally:
        # Server disk cleanup: remove temporary file after sending
        try:
            if os.path.exists(excel_path):
                os.remove(excel_path)
        except Exception:
            pass


# --- Text Message Handler ---
async def handle_text_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    user_id = update.effective_user.id
    state = context.user_data.get("state")

    # Main menu button intercepts
    if text == "➕ Харажат қўшиш":
        await handle_add_expense_menu(update, context)
        return
    elif text == "📊 Ҳисоботлар (Отчет)":
        await handle_reports_menu(update, context)
        return
    elif text == "📥 Excel юклаб олиш":
        await handle_excel_menu(update, context)
        return
    elif text == "🏷 Категориялар":
        await handle_categories_menu(update, context)
        return
    elif text == "↩️ Охиргисини ўчириш":
        await handle_delete_last_expense(update, context)
        return
    elif text == "ℹ️ Ёрдам":
        await cmd_help(update, context)
        return

    # Process States
    if state == "WAITING_AMOUNT":
        category = context.user_data.get("selected_category", "Бошқа")
        amount, note = parse_amount_and_note(text)

        if amount is None or amount <= 0:
            await update.message.reply_text(
                "❌ <b>Илтимос, суммани тўғри киритинг!</b>\n\n"
                "<i>Мисоллар:</i>\n"
                "• <code>300000</code>\n"
                "• <code>300000 бензин учун</code>\n"
                "• <code>300 минг озиқ-овқат</code>\n"
                "• <code>50k taxi</code>\n\n"
                "Ёки бекор қилиш учун «➕ Харажат қўшиш»ни қайта танланг.",
                parse_mode="HTML"
            )
            return

        # Add to DB
        exp_id = database.add_expense(user_id, category, amount, note)
        context.user_data.clear()

        # Compute today's total
        today_start = datetime.now().strftime("%Y-%m-%d 00:00:00")
        today_end = datetime.now().strftime("%Y-%m-%d 23:59:59")
        today_sum = database.get_summary_by_period(user_id, today_start, today_end)["grand_total"]

        note_str = f"\n📝 <b>Изоҳ:</b> {note}" if note else ""
        time_str = datetime.now().strftime("%d.%m.%Y %H:%M")

        confirm_msg = (
            f"✅ <b>Харажат муваффақиятли сақланди!</b>\n\n"
            f"🏷 <b>Категория:</b> {category}\n"
            f"💰 <b>Сумма:</b> <b>{format_sum(amount)} сўм</b>{note_str}\n"
            f"🕒 <b>Вақт:</b> {time_str}\n\n"
            f"📈 <i>Бугунги жами харажатингиз:</i> <b>{format_sum(today_sum)} сўм</b>"
        )
        await update.message.reply_text(
            confirm_msg,
            parse_mode="HTML",
            reply_markup=get_main_keyboard()
        )
        return

    elif state == "WAITING_NEW_CAT_NAME":
        new_cat = text.strip()
        if len(new_cat) < 2 or len(new_cat) > 40:
            await update.message.reply_text("❌ Категория номи 2 дан 40 белгигача бўлиши керак. Қайта киритинг:")
            return

        if not any(new_cat.startswith(emoji) for emoji in ["📌", "🔹", "▫️", "•"]):
            # Add a clean bullet icon if user didn't provide emoji
            if not any(ord(char) > 10000 for char in new_cat[:2]):
                new_cat = f"📌 {new_cat}"

        success = database.add_custom_category(user_id, new_cat)
        context.user_data.clear()

        if success:
            await update.message.reply_text(
                f"✅ Янги категория қўшилди: <b>{new_cat}</b>",
                parse_mode="HTML",
                reply_markup=get_main_keyboard()
            )
        else:
            await update.message.reply_text(
                "⚠️ Бундай категория аллақачон мавжуд.",
                reply_markup=get_main_keyboard()
            )
        return

    elif state == "WAITING_CUSTOM_DATE_FOR_REP":
        start_date, end_date, label = parse_date_range(text)
        if not start_date:
            await update.message.reply_text(
                "❌ Сана нотўғри киритилди.\n\n"
                "<i>Илтимос, қуйидаги форматда киритинг:</i>\n"
                "• <code>01.10.2024 - 09.10.2024</code>\n"
                "• <code>2024-10-01 - 2024-10-09</code>",
                parse_mode="HTML"
            )
            return

        context.user_data.clear()
        await send_text_report(update, user_id, start_date, end_date, label, is_edit=False)
        return

    elif state == "WAITING_CUSTOM_DATE_FOR_XLS":
        start_date, end_date, label = parse_date_range(text)
        if not start_date:
            await update.message.reply_text(
                "❌ Сана нотўғри киритилди.\n\n"
                "<i>Илтимос, қуйидаги форматда киритинг:</i>\n"
                "• <code>01.10.2024 - 09.10.2024</code>",
                parse_mode="HTML"
            )
            return

        context.user_data.clear()
        await send_excel_file(context, update.effective_chat.id, user_id, start_date, end_date, label)
        return

    # Default fallback: check if user directly entered amount without selecting category
    # If user directly typed e.g. "300000 бензин учун", guide them friendly
    direct_amount, direct_note = parse_amount_and_note(text)
    if direct_amount and direct_amount > 0:
        context.user_data["temp_amount"] = direct_amount
        context.user_data["temp_note"] = direct_note
        await update.message.reply_text(
            f"💡 Сиз <b>{format_sum(direct_amount)} сўм</b> миқдорида харажат киритмоқчисиз.\n\n"
            f"Илтимос, бу харажат қайси категорияга тегишли эканини танланг:",
            parse_mode="HTML",
            reply_markup=get_categories_keyboard(user_id)
        )
        return

    await update.message.reply_text(
        "👋 Илтимос, қуйидаги меню тугмаларидан бирини танланг:",
        reply_markup=get_main_keyboard()
    )


def main():
    database.init_db()

    token = config.BOT_TOKEN
    if not token or token == "SIZNING_BOT_TOKENINGIZ":
        print("\n" + "="*60)
        print("ДИҚҚАТ! Telegram Bot Token кўрсатилмаган!")
        print("Илтимос, config.py ёки .env файлига BOT_TOKEN ни киритинг.")
        print("="*60 + "\n")
        return

    app = Application.builder().token(token).build()

    # Handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CallbackQueryHandler(handle_callback_query))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_message))

    print("🤖 Бот ишга тушди! Тўхтатиш учун Ctrl+C ни босинг.")
    app.run_polling()


if __name__ == "__main__":
    main()
