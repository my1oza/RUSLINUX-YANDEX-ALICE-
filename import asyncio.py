import asyncio
from datetime import datetime
import io
import logging
import re
import sqlite3

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
import openpyxl

# !!! ВСТАВЬ СЮДА ТОКЕН И СВОЙ ЧИСЛОВОЙ ID !!!
BOT_TOKEN = "8263516139:AAGE07RUcOaB73wr05KUCWTqAmKlDgKFN7M"
ADMIN_IDS = [7465556680]  # Вставь свой ID числом, без кавычек

DAILY_SKIP_LIMIT_ALL = 5
MONTHLY_SKIP_LIMIT_PER_USER = 10

# --- Инициализация БД ---
def init_db():
    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_name TEXT,
                student_name TEXT,
                UNIQUE(group_name, student_name)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                group_name TEXT,
                student_name TEXT,
                status TEXT,
                date_str TEXT,
                month_str TEXT
            )
        """)
        conn.commit()

class AttendanceForm(StatesGroup):
    group = State()
    student = State()
    status = State()

class AdminAddForm(StatesGroup):
    waiting_for_list = State()

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


# ================= АДМИНКА: ДОБАВЛЕНИЕ СТУДЕНТОВ =================

@dp.message(Command("add_students"))
async def cmd_add_students(message: types.Message, state: FSMContext):
    # Если пользователя нет в списке админов — выдаем ошибку с его ID
    if message.from_user.id not in ADMIN_IDS:
        await message.answer(
            f"⛔ Доступ запрещен!\n"
            f"Твой ID: `{message.from_user.id}`\n"
            f"Добавь это число в список `ADMIN_IDS` в коде бота.",
            parse_mode="Markdown"
        )
        return

    await state.set_state(AdminAddForm.waiting_for_list)
    await message.answer(
        "Отправь список студентов одним сообщением.\n\n"
        "**Пример формата:**\n"
        "```\n"
        "23а - Иванов Иван\n"
        "23а - Петров Петр\n"
        "23б - Сидоров Алексей\n"
        "```\n"
        "*(Старый список студентов будет перезаписан)*",
        parse_mode="Markdown"
    )

@dp.message(AdminAddForm.waiting_for_list)
async def process_students_text(message: types.Message, state: FSMContext):
    lines = message.text.strip().split("\n")
    students_to_add = []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        parts = re.split(r"[-:;—–]", line, maxsplit=1)
        if len(parts) == 2:
            grp = parts[0].strip()
            name = parts[1].strip()
            if grp and name:
                students_to_add.append((grp, name))

    if not students_to_add:
        await message.answer("Не удалось распознать формат строки. Отправь в виде:\n`23а - Иванов Иван`", parse_mode="Markdown")
        return

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM students")
        cur.executemany("""
            INSERT OR IGNORE INTO students (group_name, student_name) 
            VALUES (?, ?)
        """, students_to_add)
        conn.commit()

    await state.clear()
    await message.answer(f"✅ Готово! Добавлено студентов: **{len(students_to_add)}**", parse_mode="Markdown")


# ================= ОСНОВНОЕ МЕНЮ ОТМЕТКИ =================

@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    print(message.from_user.id)
    await state.clear()

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT group_name FROM students ORDER BY group_name ASC")
        groups = [r[0] for r in cur.fetchall()]

    if not groups:
        await message.answer(
            f"Список студентов пуст.\n"
            f"Твой ID: `{message.from_user.id}` (проверь, прописан ли он в `ADMIN_IDS`).\n"
            f"Загрузи студентов через команду /add_students.",
            parse_mode="Markdown"
        )
        return

    buttons = [[InlineKeyboardButton(text=grp, callback_data=f"grp_{grp}")] for grp in groups]
    kb = InlineKeyboardMarkup(inline_keyboard=buttons)

    await message.answer("Выбери свою группу:", reply_markup=kb)
    await state.set_state(AttendanceForm.group)


@dp.callback_query(AttendanceForm.group, F.data.startswith("grp_"))
async def process_group(callback: types.CallbackQuery, state: FSMContext):
    group_name = callback.data.replace("grp_", "")
    await state.update_data(group=group_name)

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT student_name FROM students WHERE group_name = ? ORDER BY student_name ASC", (group_name,))
        students = [r[0] for r in cur.fetchall()]

    if not students:
        await callback.message.edit_text(f"В группе {group_name} пока нет фамилий.")
        await state.clear()
        return

    buttons = [[InlineKeyboardButton(text=name, callback_data=f"std_{name}")] for name in students]
    kb = InlineKeyboardMarkup(inline_keyboard=buttons)

    await callback.message.edit_text(f"Группа: {group_name}\nВыбери фамилию:", reply_markup=kb)
    await state.set_state(AttendanceForm.student)


@dp.callback_query(AttendanceForm.student, F.data.startswith("std_"))
async def process_student(callback: types.CallbackQuery, state: FSMContext):
    student_name = callback.data.replace("std_", "")
    await state.update_data(student=student_name)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="Присутствую", callback_data="status_present"),
            InlineKeyboardButton(text="Прогуляю", callback_data="status_skip")
        ]
    ])
    await callback.message.edit_text(f"Студент: {student_name}\nТвой выбор на сегодня:", reply_markup=kb)
    await state.set_state(AttendanceForm.status)


@dp.callback_query(AttendanceForm.status, F.data.startswith("status_"))
async def process_status(callback: types.CallbackQuery, state: FSMContext):
    status_choice = callback.data.split("_")[1]
    status_text = "Присутствует" if status_choice == "present" else "Прогул"

    data = await state.get_data()
    group_name = data["group"]
    student_name = data["student"]

    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    month_str = now.strftime("%Y-%m")

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()

        # Проверка повторной отметки
        cur.execute("SELECT id FROM records WHERE student_name = ? AND date_str = ?", (student_name, today_str))
        if cur.fetchone():
            await callback.message.edit_text("Этот студент уже отмечен на сегодня.")
            await state.clear()
            return

        # Лимиты
        if status_text == "Прогул":
            cur.execute("SELECT COUNT(*) FROM records WHERE date_str = ? AND status = 'Прогул'", (today_str,))
            today_total_skips = cur.fetchone()[0]
            if today_total_skips >= DAILY_SKIP_LIMIT_ALL:
                await callback.message.edit_text(f"Лимит прогулов на сегодня ({DAILY_SKIP_LIMIT_ALL} человек на всех) исчерпан. Придется идти на пары.")
                await state.clear()
                return

            cur.execute("SELECT COUNT(*) FROM records WHERE student_name = ? AND month_str = ? AND status = 'Прогул'", (student_name, month_str))
            user_month_skips = cur.fetchone()[0]
            if user_month_skips >= MONTHLY_SKIP_LIMIT_PER_USER:
                await callback.message.edit_text(f"Твой лимит прогулов в этом месяце ({MONTHLY_SKIP_LIMIT_PER_USER}) исчерпан. Иди на лекцию.")
                await state.clear()
                return

        # Запись
        cur.execute("""
            INSERT INTO records (user_id, group_name, student_name, status, date_str, month_str)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (callback.from_user.id, group_name, student_name, status_text, today_str, month_str))
        conn.commit()

    await callback.message.edit_text(f"Отметка принята:\n{student_name} ({group_name}) — **{status_text}**", parse_mode="Markdown")
    await state.clear()


# ================= ВЫГРУЗКА EXCEL =================

@dp.message(Command("report"))
async def cmd_report(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.answer(f"⛔ Команда доступна только администраторам. Твой ID: `{message.from_user.id}`", parse_mode="Markdown")
        return

    with sqlite3.connect("attendance.db", timeout=20) as conn:
        cur = conn.cursor()
        cur.execute("SELECT group_name, student_name, status, date_str FROM records ORDER BY date_str DESC, group_name ASC")
        rows = cur.fetchall()

    if not rows:
        await message.answer("База отметок пуста.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Посещаемость"
    ws.append(["Группа", "ФИО", "Статус", "Дата"])

    for row in rows:
        ws.append(row)

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    document = types.BufferedInputFile(file_stream.getvalue(), filename=f"attendance_{datetime.now().strftime('%Y-%m-%d')}.xlsx")

    for admin_id in ADMIN_IDS:
        try:
            await bot.send_document(chat_id=admin_id, document=document, caption="Актуальный отчёт по посещаемости.")
        except Exception:
            pass


# ================= ЗАПУСК =================

async def main():
    logging.basicConfig(level=logging.INFO) # Выведет в консоль каждое сообщение
    init_db()
    await bot.delete_webhook(drop_pending_updates=True) # Сбросит все зависшие старые апдейты
    print(">>> БОТ УСПЕШНО ЗАПУЩЕН И СЛУШАЕТ СООБЩЕНИЯ <<<")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())