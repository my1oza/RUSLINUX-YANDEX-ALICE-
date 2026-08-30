import asyncio
from datetime import datetime, timedelta
import io
import logging
import os
import re
import sqlite3

from aiogram import Bot, Dispatcher, F, types
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup
import openpyxl

# ================= КОНФИГУРАЦИЯ =================
BOT_TOKEN = "8263516139:AAGE07RUcOaB73wr05KUCWTqAmKlDgKFN7M"

SUPER_ADMIN_ID = 7465556680
ADMIN_IDS = [7465556680]

ADMIN_STICKER_ID = "CAACAgEAAxkBAAPqaorW62jtj-Tz7zb-fyhYKOKDKAEAAjcDAALmrTFHw50hqgjc5Yc9BA"

# Прокси (если не нужен — укажи None)
PROXY_URL = "http://modeler_oddQ2Z:rKBbXp3YyxrE@89.46.235.76:11453"

SCHEDULE_FILE = "расп 2к 1с (2).xlsx"
STUDENTS_FILE = "spisokgr.txt"

MAX_SKIPS_PER_LESSON = 5
MONTHLY_SKIP_LIMIT_PER_USER = 10

# Смещение даты для тестов (в днях)
DATE_OFFSET_DAYS = 0
# ================================================

def get_current_date():
    now = datetime.now() + timedelta(days=DATE_OFFSET_DAYS)
    return now.strftime("%Y-%m-%d"), now.strftime("%Y-%m")

def get_main_reply_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="Отметиться")]],
        resize_keyboard=True
    )

def get_broadcast_inline_keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="Отметиться", callback_data="start_attendance_btn")]]
    )

def init_db():
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_name TEXT,
                student_name TEXT,
                tg_user_id INTEGER UNIQUE DEFAULT NULL,
                tg_username TEXT DEFAULT NULL,
                UNIQUE(group_name, student_name)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                username TEXT,
                group_name TEXT,
                student_name TEXT,
                lesson TEXT,
                lesson_type TEXT,
                status TEXT,
                reason TEXT,
                valid_until TEXT,
                date_str TEXT,
                month_str TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS schedule (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date_str TEXT,
                time_str TEXT,
                lesson_type TEXT,
                subject TEXT
            )
        """)

        # Авто-миграция колонок
        cur.execute("PRAGMA table_info(records)")
        existing_cols = [row[1] for row in cur.fetchall()]
        required_cols = {
            "lesson": "TEXT",
            "lesson_type": "TEXT",
            "reason": "TEXT",
            "valid_until": "TEXT"
        }
        for col, col_type in required_cols.items():
            if col not in existing_cols:
                cur.execute(f"ALTER TABLE records ADD COLUMN {col} {col_type}")

        cur.execute("PRAGMA table_info(schedule)")
        sched_cols = [row[1] for row in cur.fetchall()]
        if "lesson_type" not in sched_cols:
            cur.execute("ALTER TABLE schedule ADD COLUMN lesson_type TEXT")

        # Очистка дубликатов перед созданием индекса
        cur.execute("""
            DELETE FROM records 
            WHERE id NOT IN (
                SELECT MAX(id) 
                FROM records 
                GROUP BY user_id, date_str, lesson
            )
        """)

        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS idx_user_date_lesson 
            ON records (user_id, date_str, lesson)
        """)

        conn.commit()


def load_students_from_file(filepath=STUDENTS_FILE):
    if not os.path.exists(filepath):
        logging.warning(f"Файл {filepath} не найден.")
        return 0

    students_to_add = []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = re.split(r"[-:;—–]", line, maxsplit=1)
            if len(parts) == 2:
                grp = parts[0].strip()
                name = parts[1].strip()
                if grp and name:
                    students_to_add.append((grp, name))

    if students_to_add:
        with sqlite3.connect("attendance.db", timeout=20) as conn:
            cur = conn.cursor()
            cur.executemany("INSERT OR IGNORE INTO students (group_name, student_name) VALUES (?, ?)", students_to_add)
            conn.commit()
    return len(students_to_add)


def classify_cell_color(cell):
    fill = cell.fill
    if not fill or not fill.fgColor:
        return "Л"
    fg = fill.fgColor
    rgb = getattr(fg, 'rgb', None)
    theme = getattr(fg, 'theme', None)
    tint = getattr(fg, 'tint', None)

    if rgb and any(v in str(rgb).upper() for v in ['DF84E9', 'FFDF84E9']):
        return "С"

    if theme == 3:
        if tint is not None and tint < 0.82:
            return "С"
        else:
            return "Л"

    if theme == 9:
        return "Л"

    return "Л"


def load_schedule_from_excel(file_or_path):
    try:
        wb = openpyxl.load_workbook(file_or_path, data_only=False)
        ws = wb.active
    except Exception as e:
        logging.error(f"Ошибка загрузки расписания: {e}")
        return 0

    schedule_entries = []
    max_row = ws.max_row
    max_col = ws.max_column
    
    r = 1
    while r <= max_row:
        cell_a = ws.cell(row=r, column=1)
        val_a = str(cell_a.value).strip().lower() if cell_a.value is not None else ""
        
        if "неделя" in val_a:
            if r + 1 > max_row:
                break
            
            dates = []
            for col_idx in range(2, min(8, max_col + 1)):
                d_cell = ws.cell(row=r + 1, column=col_idx)
                if d_cell.value is not None:
                    if isinstance(d_cell.value, datetime):
                        d_str = d_cell.value.strftime("%Y-%m-%d")
                    else:
                        d_str = str(d_cell.value)[:10]
                    dates.append((col_idx, d_str))

            r += 2
            current_time = ""
            while r <= max_row:
                time_cell = ws.cell(row=r, column=1)
                t_val = str(time_cell.value).strip() if time_cell.value is not None else ""
                
                if "неделя" in t_val.lower():
                    break
                
                if t_val and t_val != "None":
                    current_time = t_val

                if current_time:
                    for col_idx, d_str in dates:
                        lesson_cell = ws.cell(row=r, column=col_idx)
                        subj = lesson_cell.value
                        if subj is not None and str(subj).strip() != "" and str(subj).strip().lower() != "none":
                            l_type = classify_cell_color(lesson_cell)
                            schedule_entries.append((d_str, current_time, l_type, str(subj).strip()))
                r += 1
        else:
            r += 1

    if schedule_entries:
        with sqlite3.connect("attendance.db", timeout=20) as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM schedule")
            cur.executemany("INSERT INTO schedule (date_str, time_str, lesson_type, subject) VALUES (?, ?, ?, ?)", schedule_entries)
            conn.commit()
    return len(schedule_entries)


def get_lessons_for_date(date_str):
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT time_str, lesson_type, subject FROM schedule WHERE date_str = ? ORDER BY id ASC", (date_str,))
        return cur.fetchall()

def get_user_marked_lessons(user_id, date_str):
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT lesson, status FROM records WHERE user_id = ? AND date_str = ?", (user_id, date_str))
        rows = cur.fetchall()
        return {r[0]: r[1] for r in rows}

def get_active_excused_absence(user_id, date_str):
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT valid_until, lesson FROM records 
            WHERE user_id = ? AND status = 'Отсутствует (Уважительная)' 
            AND date_str <= ? AND valid_until >= ?
            ORDER BY valid_until DESC LIMIT 1
        """, (user_id, date_str, date_str))
        return cur.fetchone()

def get_user_unexcused_skips(student_name, month_str):
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT COUNT(*) FROM records 
            WHERE student_name = ? AND month_str = ? 
            AND status = 'Отсутствует (Неуважительная)' 
            AND lesson_type = 'Л'
        """, (student_name, month_str))
        used = cur.fetchone()[0]
    return max(0, MONTHLY_SKIP_LIMIT_PER_USER - used)

def get_skips_count_for_lesson(date_str, lesson_name):
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT COUNT(*) FROM records 
            WHERE date_str = ? AND lesson = ? AND status = 'Отсутствует (Неуважительная)'
        """, (date_str, lesson_name))
        return cur.fetchone()[0]

def build_morning_message(date_str):
    lessons = get_lessons_for_date(date_str)
    if lessons:
        lectures = [l for l in lessons if l[1] == "Л"]
        seminars = [l for l in lessons if l[1] == "С"]
        sched_text = "<b>Расписание учебных занятий:</b>\n"
        if lectures:
            sched_text += "\n[Л] <i>Лекции:</i>\n"
            for _, _, l_subj in lectures:
                sched_text += f"• {l_subj}\n"
        if seminars:
            sched_text += "\n[С] <i>Семинары:</i>\n"
            for _, _, l_subj in seminars:
                sched_text += f"• {l_subj}\n"
    else:
        sched_text = "<i>На сегодня запланированных занятий по расписанию нет.</i>\n"

    return (
        f"📌 <b>Уведомление о начале учебного дня</b>\n"
        f"📅 Дата: <code>{date_str}</code>\n\n"
        f"{sched_text}\n"
        f"Для фиксации статуса посещаемости нажмите кнопку ниже."
    )

def get_student_lessons_keyboard(user_id, today_str):
    lessons = get_lessons_for_date(today_str)
    user_marked = get_user_marked_lessons(user_id, today_str)

    unmarked_all = [l for l in lessons if l[2] not in user_marked]

    buttons = []
    if lessons:
        for idx, (_, l_type, l_subj) in enumerate(lessons):
            badge = f"[{l_type}]"
            if l_subj in user_marked:
                st_icon = "✅" if "Присутствует" in user_marked[l_subj] else "❌"
                buttons.append([InlineKeyboardButton(text=f"🔒 {badge} {l_subj} [{st_icon}]", callback_data=f"les_locked_{idx}")])
            else:
                buttons.append([InlineKeyboardButton(text=f"{badge} {l_subj}", callback_data=f"les_single_{idx}")])
        
        if len(unmarked_all) > 1:
            buttons.append([InlineKeyboardButton(text=f"Весь день ({len(unmarked_all)})", callback_data="les_type_all")])
    else:
        if "Весь день" in user_marked:
            buttons.append([InlineKeyboardButton(text="🔒 Весь день [Отмечено]", callback_data="les_locked_all")])
        else:
            buttons.append([InlineKeyboardButton(text="Отметиться на весь день", callback_data="les_type_all")])

    buttons.append([InlineKeyboardButton(text="🏁 Завершить", callback_data="les_done")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

# --- FSM ---
class RegForm(StatesGroup):
    group = State()
    student = State()

class MarkForm(StatesGroup):
    lesson = State()
    status_choice = State()
    excused_date = State()

class AdminForm(StatesGroup):
    waiting_for_schedule_file = State()

if PROXY_URL:
    session = AiohttpSession(proxy=PROXY_URL)
    bot = Bot(token=BOT_TOKEN, session=session)
else:
    bot = Bot(token=BOT_TOKEN)

dp = Dispatcher(storage=MemoryStorage())

def is_super_admin(user_id: int) -> bool:
    return user_id == SUPER_ADMIN_ID

def is_any_admin(user_id: int) -> bool:
    return (user_id == SUPER_ADMIN_ID) or (user_id in ADMIN_IDS)


# ================= ОСНОВНОЕ АДМИН-МЕНЮ =================

def get_admin_keyboard(user_id: int) -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text="📋 Текстовый отчёт за сегодня", callback_data="adm_get_report")],
        [InlineKeyboardButton(text="🏥 Снять уваж. причину (досрочно)", callback_data="adm_cancel_excused_menu")],
        [InlineKeyboardButton(text="👤 Открепить студента от TG", callback_data="adm_unbind_menu")]
    ]
    if is_super_admin(user_id):
        buttons.append([InlineKeyboardButton(text="🛠 Меню разработчика", callback_data="adm_dev_menu")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def get_dev_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📈 Статистика за сегодня", callback_data="adm_today_stats")],
        [InlineKeyboardButton(text="📢 Протестировать рассылку сейчас", callback_data="adm_test_broadcast")],
        [InlineKeyboardButton(text="🔄 Обновить список из spisokgr.txt", callback_data="adm_reload_txt_students")],
        [InlineKeyboardButton(text="📅 Обновить файл расписания (.xlsx)", callback_data="adm_upload_sched")],
        [
            InlineKeyboardButton(text="⏪ -1 День", callback_data="adm_date_minus"),
            InlineKeyboardButton(text="⏩ +1 День", callback_data="adm_date_plus")
        ],
        [InlineKeyboardButton(text="🔄 Сбросить дату на реальную", callback_data="adm_date_reset")],
        [InlineKeyboardButton(text="🗑 Очистить отметки (сброс)", callback_data="adm_reset_records_menu")],
        [InlineKeyboardButton(text="⚠️ Сбросить ВСЕ привязки к TG", callback_data="adm_reset_bindings")],
        [InlineKeyboardButton(text="💰 Бабло", callback_data="adm_secret_money")],
        [InlineKeyboardButton(text="◀️ В главное админ-меню", callback_data="adm_back")]
    ])

@dp.message(Command("admin"))
async def cmd_admin(message: types.Message, state: FSMContext):
    if not is_any_admin(message.from_user.id):
        return
    await state.clear()
    today_str, _ = get_current_date()
    role_label = "Главный Администратор" if is_super_admin(message.from_user.id) else "Модератор"
    await message.answer(
        f"👑 <b>Панель администратора [{role_label}]</b>\n🕒 Дата бота: <code>{today_str}</code>",
        reply_markup=get_admin_keyboard(message.from_user.id),
        parse_mode="HTML"
    )

@dp.message(F.sticker.file_id == ADMIN_STICKER_ID)
async def handle_admin_sticker(message: types.Message, state: FSMContext):
    if not is_any_admin(message.from_user.id):
        return
    await state.clear()
    today_str, _ = get_current_date()
    role_label = "Главный Администратор" if is_super_admin(message.from_user.id) else "Модератор"
    await message.answer(
        f"👑 <b>Панель администратора [{role_label}]</b>\n🕒 Дата бота: <code>{today_str}</code>",
        reply_markup=get_admin_keyboard(message.from_user.id),
        parse_mode="HTML"
    )

@dp.callback_query(F.data == "adm_back")
async def adm_back(callback: types.CallbackQuery, state: FSMContext):
    if not is_any_admin(callback.from_user.id):
        return
    await state.clear()
    today_str, _ = get_current_date()
    role_label = "Главный Администратор" if is_super_admin(callback.from_user.id) else "Модератор"
    await callback.message.edit_text(
        f"👑 <b>Панель администратора [{role_label}]</b>\n🕒 Дата бота: <code>{today_str}</code>",
        reply_markup=get_admin_keyboard(callback.from_user.id),
        parse_mode="HTML"
    )

# --- МЕНЮ РАЗРАБОТЧИКА ---
@dp.callback_query(F.data == "adm_dev_menu")
async def adm_dev_menu(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    today_str, _ = get_current_date()
    await callback.message.edit_text(
        f"🛠 <b>Меню разработчика</b>\n\n"
        f"🕒 Виртуальная дата: <code>{today_str}</code> (Смещение: <b>{DATE_OFFSET_DAYS:+d}д</b>)\n"
        f"Выберите необходимый инструмент:",
        reply_markup=get_dev_keyboard(),
        parse_mode="HTML"
    )

# --- ДОСРОЧНОЕ СНЯТИЕ УВАЖИТЕЛЬНОЙ ПРИЧИНЫ ---
@dp.callback_query(F.data == "adm_cancel_excused_menu")
async def adm_cancel_excused_menu(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return
    today_str, _ = get_current_date()
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT DISTINCT user_id, student_name, group_name, valid_until 
            FROM records 
            WHERE status = 'Отсутствует (Уважительная)' AND valid_until >= ?
            ORDER BY valid_until ASC
        """, (today_str,))
        active_excused = cur.fetchall()

    if not active_excused:
        kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ В главное меню", callback_data="adm_back")]])
        await callback.message.edit_text("Сейчас нет активных уважительных причин на будущее.", reply_markup=kb)
        return

    buttons = []
    for u_id, name, grp, v_until in active_excused:
        btn_text = f"❌ {name} ({grp}) — до {v_until}"
        buttons.append([InlineKeyboardButton(text=btn_text, callback_data=f"cexc_{u_id}")])
    buttons.append([InlineKeyboardButton(text="◀️ В главное меню", callback_data="adm_back")])

    await callback.message.edit_text(
        "🏥 <b>Выберите студента для досрочного завершения уважительной причины:</b>",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="HTML"
    )

@dp.callback_query(F.data.startswith("cexc_"))
async def adm_cancel_excused_execute(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return
    u_id = int(callback.data.replace("cexc_", ""))
    today_str, _ = get_current_date()

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT student_name, group_name FROM students WHERE tg_user_id = ?", (u_id,))
        st_res = cur.fetchone()
        st_name = st_res[0] if st_res else "Студент"

        cur.execute("""
            UPDATE records 
            SET valid_until = ? 
            WHERE user_id = ? AND status = 'Отсутствует (Уважительная)' AND valid_until > ?
        """, (today_str, u_id, today_str))
        conn.commit()

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🏥 Снять еще", callback_data="adm_cancel_excused_menu")],
        [InlineKeyboardButton(text="◀️ В главное меню", callback_data="adm_back")]
    ])
    await callback.message.edit_text(
        f"✅ Уважительная причина для <b>{st_name}</b> досрочно закрыта! Студент теперь может отмечаться.",
        reply_markup=kb,
        parse_mode="HTML"
    )

# --- Кнопка перезагрузки списка из файла ---
@dp.callback_query(F.data == "adm_reload_txt_students")
async def adm_reload_txt_students(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    count = load_students_from_file(STUDENTS_FILE)
    if count > 0:
        await callback.message.answer(f"✅ Список успешно обновлен из <code>{STUDENTS_FILE}</code>! Найдено студентов: <b>{count}</b> (привязки сохранены).", parse_mode="HTML")
    else:
        await callback.message.answer(f"⚠️ Не удалось загрузить студентов из <code>{STUDENTS_FILE}</code>.", parse_mode="HTML")
    await callback.answer()

# --- Тест рассылки ---
@dp.callback_query(F.data == "adm_test_broadcast")
async def adm_test_broadcast(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return

    today_str, _ = get_current_date()
    broadcast_msg = build_morning_message(today_str)
    broadcast_kb = get_broadcast_inline_keyboard()

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT tg_user_id FROM students WHERE tg_user_id IS NOT NULL")
        users = cur.fetchall()

    if not users:
        await callback.message.answer("⚠️ В базе данных пока нет зарегистрированных студентов с Telegram ID.")
        await callback.answer()
        return

    success_count = 0
    fail_count = 0

    for (u_id,) in users:
        try:
            await bot.send_message(
                chat_id=u_id, 
                text=broadcast_msg, 
                reply_markup=broadcast_kb, 
                parse_mode="HTML"
            )
            success_count += 1
            await asyncio.sleep(0.05)
        except Exception as e:
            logging.error(f"Ошибка отправки пользователю {u_id}: {e}")
            fail_count += 1

    await callback.message.answer(
        f"📢 <b>Тестовая рассылка завершена!</b>\n\n"
        f"✅ Доставлено: <b>{success_count}</b>\n"
        f"❌ Ошибок: <b>{fail_count}</b>\n"
        f"📅 Дата: <code>{today_str}</code>",
        parse_mode="HTML"
    )
    await callback.answer("Рассылка выполнена!")

# --- Статистика ---
@dp.callback_query(F.data == "adm_today_stats")
async def adm_today_stats(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return
    today_str, _ = get_current_date()
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM records WHERE date_str = ?", (today_str,))
        total_marked = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM records WHERE date_str = ? AND status LIKE 'Отсутствует%'", (today_str,))
        total_skips = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM records WHERE date_str = ? AND status = 'Присутствует'", (today_str,))
        total_presents = cur.fetchone()[0]

    text = (
        f"📅 <b>Статистика за {today_str}:</b>\n\n"
        f"👥 Всего отметок: <b>{total_marked}</b>\n"
        f"✅ Присутствуют: <b>{total_presents}</b>\n"
        f"❌ Отсутствуют: <b>{total_skips}</b>"
    )
    buttons = [[InlineKeyboardButton(text="◀️ Назад в меню разработчика", callback_data="adm_dev_menu")]]
    await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

@dp.callback_query(F.data == "adm_secret_money")
async def adm_secret_money(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    text = (
        "💰 <b>СЕКРЕТНЫЙ ОФШОРНЫЙ СЧЁТ АКТИВИРОВАН</b> 💰\n\n"
        "💳 Баланс: <b>$1,000,000,000 USDT</b>\n"
        "💸 Статус: <i>Успешно отмыто через старостат</i>\n\n"
        "Все прогулы официально профинансированы! 😎🍾"
    )
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Назад в меню разработчика", callback_data="adm_dev_menu")]])
    await callback.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
    try:
        await callback.message.answer_sticker(ADMIN_STICKER_ID)
    except Exception:
        pass

# --- Открепление студентов ---
@dp.callback_query(F.data == "adm_unbind_menu")
async def adm_unbind_menu(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT group_name FROM students WHERE tg_user_id IS NOT NULL ORDER BY group_name ASC")
        groups = [r[0] for r in cur.fetchall()]
    if not groups:
        kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Назад в главное меню", callback_data="adm_back")]])
        await callback.message.edit_text("Зарегистрированных студентов нет.", reply_markup=kb)
        return
    buttons = [[InlineKeyboardButton(text=f"Группа {grp}", callback_data=f"unb_grp_{grp}")] for grp in groups]
    buttons.append([InlineKeyboardButton(text="◀️ Назад в главное меню", callback_data="adm_back")])
    await callback.message.edit_text("Выбери группу:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))

@dp.callback_query(F.data.startswith("unb_grp_"))
async def adm_unbind_group_select(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return
    group_name = callback.data.replace("unb_grp_", "")
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, student_name FROM students WHERE group_name = ? AND tg_user_id IS NOT NULL ORDER BY student_name ASC", (group_name,))
        students = cur.fetchall()
    buttons = []
    for s_id, s_name in students:
        buttons.append([InlineKeyboardButton(text=f"❌ {s_name}", callback_data=f"unb_id_{s_id}")])
    buttons.append([InlineKeyboardButton(text="◀️ Назад к группам", callback_data="adm_unbind_menu")])
    await callback.message.edit_text(f"Выбери студента группы <b>{group_name}</b> для открепления:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

@dp.callback_query(F.data.startswith("unb_id_"))
async def adm_unbind_execute(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return
    student_id = int(callback.data.replace("unb_id_", ""))
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT student_name, group_name FROM students WHERE id = ?", (student_id,))
        res = cur.fetchone()
        if res:
            cur.execute("UPDATE students SET tg_user_id = NULL, tg_username = NULL WHERE id = ?", (student_id,))
            conn.commit()
            msg = f"✅ Telegram-аккаунт откреплен от: <b>{res[0]}</b> ({res[1]})."
        else:
            msg = "Студент не найден."
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="👤 Открепить еще", callback_data="adm_unbind_menu")],
        [InlineKeyboardButton(text="◀️ Главное меню", callback_data="adm_back")]
    ])
    await callback.message.edit_text(msg, reply_markup=kb, parse_mode="HTML")

# --- Функции Главного Админа ---
@dp.callback_query(F.data == "adm_upload_sched")
async def adm_upload_sched_btn(callback: types.CallbackQuery, state: FSMContext):
    if not is_super_admin(callback.from_user.id):
        return
    await state.set_state(AdminForm.waiting_for_schedule_file)
    await callback.message.edit_text("Отправь новый файл расписания <code>.xlsx</code> документом сюда в чат.", parse_mode="HTML")

@dp.message(AdminForm.waiting_for_schedule_file, F.document)
async def process_schedule_file(message: types.Message, state: FSMContext):
    if not is_super_admin(message.from_user.id):
        return
    file_bytes = io.BytesIO()
    await bot.download(message.document, destination=file_bytes)
    file_bytes.seek(0)
    count = load_schedule_from_excel(file_bytes)
    await state.clear()
    await message.answer(f"✅ Расписание успешно обновлено! Загружено занятий: <b>{count}</b>", parse_mode="HTML")

@dp.callback_query(F.data.in_(["adm_date_plus", "adm_date_minus", "adm_date_reset"]))
async def adm_change_date(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    global DATE_OFFSET_DAYS
    if callback.data == "adm_date_plus":
        DATE_OFFSET_DAYS += 1
    elif callback.data == "adm_date_minus":
        DATE_OFFSET_DAYS -= 1
    elif callback.data == "adm_date_reset":
        DATE_OFFSET_DAYS = 0
    today_str, _ = get_current_date()
    await callback.message.edit_text(
        f"🛠 <b>Меню разработчика</b>\n\n"
        f"🕒 Виртуальная дата: <code>{today_str}</code> (Смещение: <b>{DATE_OFFSET_DAYS:+d}д</b>)\n"
        f"Выберите необходимый инструмент:",
        reply_markup=get_dev_keyboard(),
        parse_mode="HTML"
    )
    await callback.answer(f"Дата: {today_str}")

# --- Меню очистки отметок ---
@dp.callback_query(F.data == "adm_reset_records_menu")
async def adm_reset_records_menu(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    today_str, _ = get_current_date()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"🗑 Удалить только за сегодня ({today_str})", callback_data="adm_clear_today")],
        [InlineKeyboardButton(text="💥 Удалить ВСЮ историю посещений", callback_data="adm_clear_all")],
        [InlineKeyboardButton(text="◀️ Назад в меню разработчика", callback_data="adm_dev_menu")]
    ])
    await callback.message.edit_text("⚙️ <b>Выбери вариант очистки отметок:</b>\n<i>(Привязки студентов к Telegram затронуты НЕ будут)</i>", reply_markup=kb, parse_mode="HTML")

@dp.callback_query(F.data == "adm_clear_today")
async def adm_clear_today(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    today_str, _ = get_current_date()
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM records WHERE date_str = ?", (today_str,))
        conn.commit()
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Меню разработчика", callback_data="adm_dev_menu")]])
    await callback.message.edit_text(f"✅ Отметки за <code>{today_str}</code> удалены (привязки студентов сохранены).", reply_markup=kb, parse_mode="HTML")

@dp.callback_query(F.data == "adm_clear_all")
async def adm_clear_all(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM records")
        conn.commit()
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Меню разработчика", callback_data="adm_dev_menu")]])
    await callback.message.edit_text("💥 Вся история посещений очищена!", reply_markup=kb, parse_mode="HTML")

@dp.callback_query(F.data == "adm_reset_bindings")
async def adm_reset_bindings(callback: types.CallbackQuery):
    if not is_super_admin(callback.from_user.id):
        return
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("UPDATE students SET tg_user_id = NULL, tg_username = NULL")
        conn.commit()
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Меню разработчика", callback_data="adm_dev_menu")]])
    await callback.message.edit_text("🔄 Все привязки аккаунтов студентов сброшены.", reply_markup=kb)


# ================= ЧИТАБЕЛЬНЫЙ ОТЧЁТ (БЕЗ АЙДИ И ЮЗЕРНЕЙМОВ) =================

async def build_and_send_text_report(chat_id: int):
    today_str, _ = get_current_date()
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        
        # 1. Записи за текущую дату
        cur.execute("""
            SELECT group_name, student_name, lesson, lesson_type, status, reason, valid_until 
            FROM records 
            WHERE date_str = ?
        """, (today_str,))
        direct_rows = cur.fetchall()

        # 2. Активные длительные уважительные причины
        cur.execute("""
            SELECT group_name, student_name, lesson, lesson_type, status, reason, valid_until 
            FROM records 
            WHERE status = 'Отсутствует (Уважительная)' AND date_str < ? AND valid_until >= ?
        """, (today_str, today_str))
        active_excused_rows = cur.fetchall()

    all_rows = direct_rows + active_excused_rows

    if not all_rows:
        await bot.send_message(chat_id, f"📌 За день <code>{today_str}</code> отметок пока нет.", parse_mode="HTML")
        return

    groups = {}
    for grp, name, lesson, l_type, status, reason, valid_until in all_rows:
        if grp not in groups:
            groups[grp] = {}
        if name not in groups[grp]:
            groups[grp][name] = {
                "present": set(),
                "excused": set(),
                "unexcused": set()
            }
        
        badge = f"[{l_type.upper()}] " if l_type else ""
        entry_text = f"{badge}{lesson}" if lesson else "Весь день"

        if status == "Присутствует":
            groups[grp][name]["present"].add(entry_text)
        elif "Уважительная" in status:
            u_date = f" (до {valid_until})" if valid_until else ""
            groups[grp][name]["excused"].add(f"{entry_text}{u_date}")
        else:
            groups[grp][name]["unexcused"].add(entry_text)

    total_unique_students = sum(len(st_dict) for st_dict in groups.values())
    report_lines = [
        f"📊 <b>ОТЧЁТ ПО ПОСЕЩАЕМОСТИ</b>",
        f"📅 Дата: <code>{today_str}</code>",
        f"👥 Всего отметившихся: <b>{total_unique_students}</b>",
        "━━━━━━━━━━━━━━━━━━━━\n"
    ]

    for grp, students in groups.items():
        report_lines.append(f"🎓 <b>ГРУППА {grp}</b>")
        
        present_lines = []
        excused_lines = []
        unexcused_lines = []

        for name, data in students.items():
            if data["present"]:
                lessons_str = ", ".join(sorted(list(data["present"])))
                present_lines.append(f"  • <b>{name}</b>: {lessons_str}")
            if data["excused"]:
                lessons_str = ", ".join(sorted(list(data["excused"])))
                excused_lines.append(f"  • <b>{name}</b>: {lessons_str}")
            if data["unexcused"]:
                lessons_str = ", ".join(data["unexcused"])
                unexcused_lines.append(f"  • <b>{name}</b>: {lessons_str}")

        if present_lines:
            report_lines.append("✅ <b>Присутствуют:</b>")
            report_lines.extend(present_lines)
        if excused_lines:
            report_lines.append("\n🏥 <b>Уважительная причина:</b>")
            report_lines.extend(excused_lines)
        if unexcused_lines:
            report_lines.append("\n❌ <b>Прогулы (неуважительная):</b>")
            report_lines.extend(unexcused_lines)

        report_lines.append("────────────────────\n")

    await bot.send_message(chat_id, "\n".join(report_lines), parse_mode="HTML")

@dp.message(Command("report"))
async def cmd_report(message: types.Message):
    if not is_any_admin(message.from_user.id):
        return
    await build_and_send_text_report(message.chat.id)

@dp.callback_query(F.data == "adm_get_report")
async def cb_report(callback: types.CallbackQuery):
    if not is_any_admin(callback.from_user.id):
        return
    await build_and_send_text_report(callback.message.chat.id)
    await callback.answer()


# ================= КЛИЕНТСКАЯ ЧАСТЬ (/ZP И КНОПКА СТАРТА) =================

async def trigger_attendance_flow(event: types.Message | types.CallbackQuery, state: FSMContext):
    await state.clear()
    user = event.from_user
    user_id = user.id
    today_str, month_str = get_current_date()

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT group_name, student_name FROM students WHERE tg_user_id = ?", (user_id,))
        registered_user = cur.fetchone()

        if registered_user:
            group_name, student_name = registered_user
            await state.update_data(group=group_name, student=student_name)

            active_excused = get_active_excused_absence(user_id, today_str)
            if active_excused:
                v_until = active_excused[0]
                text = (
                    f"Привет, <b>{student_name}</b> ({group_name})!\n\n"
                    f"🏥 У вас зафиксировано отсутствие по <b>уважительной причине</b> до <code>{v_until}</code> включительно.\n"
                    f"Отмечаться повторно не требуется — вы автоматически отображаетесь в отчётах."
                )
                if isinstance(event, types.CallbackQuery):
                    await event.message.edit_text(text, parse_mode="HTML")
                else:
                    await event.answer(text, reply_markup=get_main_reply_keyboard(), parse_mode="HTML")
                return

            lessons = get_lessons_for_date(today_str)
            kb = get_student_lessons_keyboard(user_id, today_str)

            text = (
                f"Привет, <b>{student_name}</b> ({group_name})!\n"
                f"📅 Дата: <code>{today_str}</code> | Занятий сегодня: <b>{len(lessons)}</b>\n\n"
                f"Выбери пару для фиксации посещаемости (<b>внимание:</b> после выбора изменить решение нельзя):"
            )
            if isinstance(event, types.CallbackQuery):
                await event.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
            else:
                await event.answer(text, reply_markup=kb, parse_mode="HTML")
            await state.set_state(MarkForm.lesson)
            return

        cur.execute("SELECT DISTINCT group_name FROM students ORDER BY group_name ASC")
        groups = [r[0] for r in cur.fetchall()]

    if not groups:
        text = "База студентов еще не заполнена (проверь файл spisokgr.txt)."
        if isinstance(event, types.CallbackQuery):
            await event.message.edit_text(text)
        else:
            await event.answer(text, reply_markup=get_main_reply_keyboard())
        return

    buttons = [[InlineKeyboardButton(text=grp, callback_data=f"reg_grp_{grp}")] for grp in groups]
    text = "👋 <b>Регистрация в системе посещаемости</b>\n\nВыбери свою группу:"
    if isinstance(event, types.CallbackQuery):
        await event.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    else:
        await event.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(RegForm.group)


@dp.message(Command("start"))
async def cmd_start_redirect(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT group_name, student_name FROM students WHERE tg_user_id = ?", (user_id,))
        registered_user = cur.fetchone()

    if registered_user:
        await trigger_attendance_flow(message, state)
    else:
        kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Отметиться", callback_data="start_attendance_btn")]])
        await message.answer(
            "👋 Добро пожаловать в систему учёта посещаемости!\n\n"
            "Нажмите кнопку ниже, чтобы пройти регистрацию и отметить своё присутствие:",
            reply_markup=kb
        )

@dp.message(F.text == "Отметиться")
async def btn_start_attendance(message: types.Message, state: FSMContext):
    await trigger_attendance_flow(message, state)

@dp.callback_query(F.data == "start_attendance_btn")
async def cb_start_attendance(callback: types.CallbackQuery, state: FSMContext):
    await trigger_attendance_flow(callback, state)

@dp.message(Command("zp"))
async def cmd_zp(message: types.Message, state: FSMContext):
    await trigger_attendance_flow(message, state)


# --- Регистрация ---
@dp.callback_query(RegForm.group, F.data.startswith("reg_grp_"))
async def process_reg_group(callback: types.CallbackQuery, state: FSMContext):
    group_name = callback.data.replace("reg_grp_", "")
    await state.update_data(group=group_name)

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, student_name FROM students WHERE group_name = ? AND tg_user_id IS NULL ORDER BY student_name ASC", (group_name,))
        available_students = cur.fetchall()

    if not available_students:
        await callback.message.edit_text(f"В группе {group_name} все студенты уже зарегистрированы.")
        await state.clear()
        return

    buttons = [[InlineKeyboardButton(text=name, callback_data=f"reg_id_{s_id}")] for s_id, name in available_students]
    await callback.message.edit_text(f"Группа: <b>{group_name}</b>\n\nВыбери свое ФИО:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(RegForm.student)


@dp.callback_query(RegForm.student, F.data.startswith("reg_id_"))
async def process_reg_student(callback: types.CallbackQuery, state: FSMContext):
    student_id = int(callback.data.replace("reg_id_", ""))
    user_id = callback.from_user.id
    username = callback.from_user.username or "нет_юзернейма"
    today_str, _ = get_current_date()

    data = await state.get_data()
    group_name = data.get("group")

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT student_name, tg_user_id FROM students WHERE id = ?", (student_id,))
        res = cur.fetchone()
        if not res or res[1] is not None:
            await callback.message.edit_text("❌ Эту фамилию уже занял другой человек.")
            await state.clear()
            return
        cur.execute("UPDATE students SET tg_user_id = ?, tg_username = ? WHERE id = ?", (user_id, username, student_id))
        conn.commit()

    student_name = res[0]
    await state.update_data(student=student_name)

    kb = get_student_lessons_keyboard(user_id, today_str)
    await callback.message.edit_text(
        f"✅ <b>Регистрация завершена!</b>\n\n"
        f"Ты закреплен как: <b>{student_name}</b> ({group_name})\n"
        f"📅 Дата: <code>{today_str}</code>\n\n"
        f"Выбери пару или вариант отметки:",
        reply_markup=kb,
        parse_mode="HTML"
    )
    await state.set_state(MarkForm.lesson)


# --- Запрет на повторную отметку ---
@dp.callback_query(MarkForm.lesson, F.data.startswith("les_locked_"))
async def process_locked_lesson(callback: types.CallbackQuery):
    await callback.answer("🔒 Вы уже зафиксировали статус по этому занятию. Менять решение нельзя!", show_alert=True)


# --- Выбор пары или всего дня ---
@dp.callback_query(MarkForm.lesson, F.data.startswith("les_"))
async def process_lesson_choice(callback: types.CallbackQuery, state: FSMContext):
    action = callback.data.replace("les_", "")
    
    if action == "done":
        await callback.message.edit_text("✅ <b>Отметка посещаемости завершена.</b> Хорошего учебного дня!", parse_mode="HTML")
        await state.clear()
        return

    today_str, month_str = get_current_date()
    data = await state.get_data()
    student_name = data["student"]

    all_lessons = get_lessons_for_date(today_str)
    user_marked = get_user_marked_lessons(callback.from_user.id, today_str)

    target_lessons = []
    label_desc = ""

    if action.startswith("single_"):
        idx = int(action.replace("single_", ""))
        l_time, l_type, l_subj = all_lessons[idx]
        if l_subj in user_marked:
            await callback.answer("🔒 Это занятие уже отмечено ранее!", show_alert=True)
            return
        target_lessons = [(l_time, l_type, l_subj)]
        label_desc = f"[{l_type.upper()}] {l_subj}"
    elif action == "type_all":
        target_lessons = [l for l in all_lessons if l[2] not in user_marked]
        label_desc = f"Весь день ({len(target_lessons)})"

    if not target_lessons:
        await callback.answer("Все выбранные занятия уже были отмечены ранее!", show_alert=True)
        return

    lectures_count = sum(1 for l in target_lessons if l[1] == "Л")
    await state.update_data(target_lessons=target_lessons, label_desc=label_desc, lectures_count=lectures_count)

    skips_left = get_user_unexcused_skips(student_name, month_str)

    buttons = [
        [InlineKeyboardButton(text="✅ Буду присутствовать", callback_data="stat_present")],
        [InlineKeyboardButton(text="🏥 Уважительная причина (с выбором даты)", callback_data="stat_excused")]
    ]

    if skips_left >= lectures_count:
        if lectures_count > 0:
            skip_cost_text = f"-{lectures_count}"
        else:
            skip_cost_text = "0 (семинар)"
            
        buttons.append([InlineKeyboardButton(
            text=f"❌ Прогул ({skip_cost_text}, осталось: {skips_left}/{MONTHLY_SKIP_LIMIT_PER_USER})", 
            callback_data="stat_unexcused"
        )])

    buttons.append([InlineKeyboardButton(text="◀️ Назад к списку занятий", callback_data="stat_back_to_list")])

    await callback.message.edit_text(
        f"Выбрано: <b>{label_desc}</b>\n"
        f"Твой статус на <code>{today_str}</code> (<b>изменить в будущем будет нельзя</b>):",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="HTML"
    )
    await state.set_state(MarkForm.status_choice)


# --- Обработка статуса ---
@dp.callback_query(MarkForm.status_choice, F.data.startswith("stat_"))
async def process_status_choice(callback: types.CallbackQuery, state: FSMContext):
    stat_type = callback.data.replace("stat_", "")
    user_id = callback.from_user.id
    username = callback.from_user.username or "нет_юзернейма"
    data = await state.get_data()
    group_name = data["group"]
    student_name = data["student"]
    target_lessons = data.get("target_lessons", [])
    label_desc = data.get("label_desc", "Занятие")
    lectures_count = data.get("lectures_count", 0)
    today_str, month_str = get_current_date()

    if stat_type == "back_to_list":
        kb = get_student_lessons_keyboard(user_id, today_str)
        await callback.message.edit_text(
            f"👤 <b>{student_name}</b> ({group_name})\n"
            f"📅 Дата: <code>{today_str}</code>\n\n"
            f"Выбери пару или вариант отметки:",
            reply_markup=kb,
            parse_mode="HTML"
        )
        await state.set_state(MarkForm.lesson)
        return

    user_marked = get_user_marked_lessons(user_id, today_str)
    filtered_lessons = [l for l in target_lessons if l[2] not in user_marked]
    if not filtered_lessons and target_lessons:
        await callback.answer("🔒 Эти занятия уже зафиксированы ранее!", show_alert=True)
        kb = get_student_lessons_keyboard(user_id, today_str)
        await callback.message.edit_text(f"Все занятия уже отмечены.", reply_markup=kb, parse_mode="HTML")
        await state.set_state(MarkForm.lesson)
        return

    if stat_type == "present":
        status_text = "Присутствует"
        with sqlite3.connect("attendance.db", timeout=20) as conn:
            cur = conn.cursor()
            if filtered_lessons:
                for _, l_type, l_subj in filtered_lessons:
                    cur.execute("""
                        INSERT OR IGNORE INTO records (user_id, username, group_name, student_name, lesson, lesson_type, status, reason, valid_until, date_str, month_str)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (user_id, username, group_name, student_name, l_subj, l_type, status_text, "Присутствие", None, today_str, month_str))
            else:
                cur.execute("""
                    INSERT OR IGNORE INTO records (user_id, username, group_name, student_name, lesson, lesson_type, status, reason, valid_until, date_str, month_str)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (user_id, username, group_name, student_name, label_desc, "Общее", status_text, "Присутствие", None, today_str, month_str))
            conn.commit()

        kb = get_student_lessons_keyboard(user_id, today_str)
        await callback.message.edit_text(
            f"✅ Зафиксировано: <b>{label_desc}</b> — <i>{status_text}</i>\n\n"
            f"Выбери следующее занятие или нажми «Завершить»:",
            reply_markup=kb,
            parse_mode="HTML"
        )
        await state.set_state(MarkForm.lesson)

    elif stat_type == "unexcused":
        status_text = "Отсутствует (Неуважительная)"
        
        if filtered_lessons:
            for _, _, l_subj in filtered_lessons:
                current_skips = get_skips_count_for_lesson(today_str, l_subj)
                if current_skips >= MAX_SKIPS_PER_LESSON:
                    await callback.message.edit_text(
                        f"⛔ На пару <b>{l_subj}</b> уже исчерпан лимит прогулов ({MAX_SKIPS_PER_LESSON} человек).",
                        parse_mode="HTML"
                    )
                    await state.set_state(MarkForm.lesson)
                    return

        with sqlite3.connect("attendance.db", timeout=20) as conn:
            cur = conn.cursor()
            if filtered_lessons:
                for _, l_type, l_subj in filtered_lessons:
                    cur.execute("""
                        INSERT OR IGNORE INTO records (user_id, username, group_name, student_name, lesson, lesson_type, status, reason, valid_until, date_str, month_str)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (user_id, username, group_name, student_name, l_subj, l_type, status_text, "Неуважительная", None, today_str, month_str))
            else:
                cur.execute("""
                    INSERT OR IGNORE INTO records (user_id, username, group_name, student_name, lesson, lesson_type, status, reason, valid_until, date_str, month_str)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (user_id, username, group_name, student_name, label_desc, "Общее", status_text, "Неуважительная", None, today_str, month_str))
            conn.commit()

        skips_left = get_user_unexcused_skips(student_name, month_str)
        cost_msg = f"📉 Списано лекций: <b>{lectures_count}</b>" if lectures_count > 0 else "ℹ️ Семинары не тратят лимит прогулов"

        kb = get_student_lessons_keyboard(user_id, today_str)
        await callback.message.edit_text(
            f"❌ Зафиксировано: <b>{label_desc}</b> — <i>{status_text}</i>\n"
            f"{cost_msg}\n"
            f"Осталось прогулов лекций в месяце: <b>{skips_left}/{MONTHLY_SKIP_LIMIT_PER_USER}</b>\n\n"
            f"Выбери следующее занятие или нажми «Завершить»:",
            reply_markup=kb,
            parse_mode="HTML"
        )
        await state.set_state(MarkForm.lesson)

    elif stat_type == "excused":
        await state.set_state(MarkForm.excused_date)
        now_dt = datetime.strptime(today_str, "%Y-%m-%d")
        d1 = (now_dt + timedelta(days=1)).strftime("%Y-%m-%d")
        d3 = (now_dt + timedelta(days=3)).strftime("%Y-%m-%d")
        d7 = (now_dt + timedelta(days=7)).strftime("%Y-%m-%d")

        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=f"Только сегодня ({today_str})", callback_data=f"exd_{today_str}")],
            [InlineKeyboardButton(text=f"До завтра ({d1})", callback_data=f"exd_{d1}")],
            [InlineKeyboardButton(text=f"На 3 дня (до {d3})", callback_data=f"exd_{d3}")],
            [InlineKeyboardButton(text=f"На неделю (до {d7})", callback_data=f"exd_{d7}")]
        ])
        await callback.message.edit_text(
            "🏥 <b>Уважительная причина</b>\n\n"
            "Выбери дату окончания отсутствия кнопкой или отправь дату сообщением (в формате <code>ГГГГ-ММ-ДД</code>):\n"
            "<i>(До этой даты ты будешь автоматически отмечаться в отчётах)</i>",
            reply_markup=kb,
            parse_mode="HTML"
        )


# --- Выбор даты для уважительной причины ---
@dp.callback_query(MarkForm.excused_date, F.data.startswith("exd_"))
async def process_excused_date_btn(callback: types.CallbackQuery, state: FSMContext):
    valid_until = callback.data.replace("exd_", "")
    await save_excused_record(callback.message, state, callback.from_user, valid_until, is_callback=True)

@dp.message(MarkForm.excused_date)
async def process_excused_date_msg(message: types.Message, state: FSMContext):
    date_text = message.text.strip()
    try:
        valid_until = datetime.strptime(date_text, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        await message.answer("⚠️ Неверный формат даты. Введи в формате <code>ГГГГ-ММ-ДД</code> (например, <code>2026-09-10</code>):", parse_mode="HTML")
        return
    await save_excused_record(message, state, message.from_user, valid_until, is_callback=False)

async def save_excused_record(message, state, user, valid_until, is_callback=False):
    data = await state.get_data()
    group_name = data["group"]
    student_name = data["student"]
    target_lessons = data.get("target_lessons", [])
    label_desc = data.get("label_desc", "Занятие")
    today_str, month_str = get_current_date()
    status_text = "Отсутствует (Уважительная)"
    username = user.username or "нет_юзернейма"

    user_marked = get_user_marked_lessons(user.id, today_str)
    filtered_lessons = [l for l in target_lessons if l[2] not in user_marked]

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        if filtered_lessons:
            for _, l_type, l_subj in filtered_lessons:
                cur.execute("""
                    INSERT OR IGNORE INTO records (user_id, username, group_name, student_name, lesson, lesson_type, status, reason, valid_until, date_str, month_str)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (user.id, username, group_name, student_name, l_subj, l_type, status_text, "Уважительная", valid_until, today_str, month_str))
        else:
            cur.execute("""
                INSERT OR IGNORE INTO records (user_id, username, group_name, student_name, lesson, lesson_type, status, reason, valid_until, date_str, month_str)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (user.id, username, group_name, student_name, label_desc, "Общее", status_text, "Уважительная", valid_until, today_str, month_str))
        conn.commit()

    kb = get_student_lessons_keyboard(user.id, today_str)
    text = (
        f"🏥 Зафиксировано: <b>{label_desc}</b> — <i>{status_text}</i> (до {valid_until})\n"
        f"ℹ️ До <code>{valid_until}</code> ты автоматически будешь отображаться в отчётах.\n\n"
        f"Выбери следующее занятие или нажми «Завершить»:"
    )
    if is_callback:
        await message.edit_text(text, reply_markup=kb, parse_mode="HTML")
    else:
        await message.answer(text, reply_markup=kb, parse_mode="HTML")
    await state.set_state(MarkForm.lesson)


# ================= ЕЖЕДНЕВНАЯ УТРЕННЯЯ РАССЫЛКА В 07:45 =================

async def morning_broadcast_task():
    while True:
        now = datetime.now()
        target = now.replace(hour=7, minute=45, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        
        wait_seconds = (target - now).total_seconds()
        await asyncio.sleep(wait_seconds)

        today_str, _ = get_current_date()
        broadcast_msg = build_morning_message(today_str)
        broadcast_kb = get_broadcast_inline_keyboard()

        with sqlite3.connect("attendance.db", timeout=20) as conn:
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT tg_user_id FROM students WHERE tg_user_id IS NOT NULL")
            users = cur.fetchall()

        for (u_id,) in users:
            try:
                await bot.send_message(
                    chat_id=u_id, 
                    text=broadcast_msg, 
                    reply_markup=broadcast_kb, 
                    parse_mode="HTML"
                )
                await asyncio.sleep(0.05)
            except Exception as e:
                logging.error(f"Ошибка автоматической рассылки пользователю {u_id}: {e}")


# ================= ЗАПУСК БОТА =================

async def main():
    logging.basicConfig(level=logging.INFO)
    init_db()
    
    try:
        st_count = load_students_from_file(STUDENTS_FILE)
        print(f"[СТУДЕНТЫ] Загружено/проверено из '{STUDENTS_FILE}': {st_count}")
    except Exception as e:
        print(f"[СТУДЕНТЫ] Ошибка загрузки списка: {e}")

    try:
        count = load_schedule_from_excel(SCHEDULE_FILE)
        print(f"[РАСПИСАНИЕ] Загружено занятий из '{SCHEDULE_FILE}': {count}")
    except Exception as e:
        print(f"[РАСПИСАНИЕ] Ошибка загрузки расписания: {e}")

    asyncio.create_task(morning_broadcast_task())

    try:
        me = await bot.get_me()
        print(f"\n[УСПЕХ] Бот @{me.username} запущен и готов к работе!\n")
        await bot.delete_webhook(drop_pending_updates=True)
        await dp.start_polling(bot)
    finally:
        await bot.session.close()

if __name__ == "__main__":
    asyncio.run(main())