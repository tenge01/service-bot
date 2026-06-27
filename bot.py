import logging
import os
import asyncio
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
import sqlite3
import openpyxl
from docx import Document
from docx.shared import Pt
import re

# ─── Настройки ───────────────────────────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN", "ВСТАВЬ_ТОКЕН")
ADMIN_CHAT_ID = int(os.getenv("ADMIN_CHAT_ID", "1048969972"))

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# ─── Состояния диалога ───────────────────────────────────────
class Form(StatesGroup):
    phone       = State()
    city        = State()
    school      = State()
    serial      = State()
    description = State()
    confirm     = State()

# ─── База данных ─────────────────────────────────────────────
def init_db():
    con = sqlite3.connect("data/tickets.db")
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS tickets (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            ticket_num  TEXT,
            phone       TEXT,
            city        TEXT,
            school      TEXT,
            serial      TEXT,
            description TEXT,
            service     TEXT,
            service_contact TEXT,
            status      TEXT DEFAULT 'новая',
            created_at  TEXT,
            tg_user_id  INTEGER
        )
    """)
    con.commit()
    con.close()

def save_ticket(data: dict) -> int:
    con = sqlite3.connect("data/tickets.db")
    cur = con.cursor()
    cur.execute("""
        INSERT INTO tickets
        (ticket_num, phone, city, school, serial, description, service, service_contact, created_at, tg_user_id)
        VALUES (?,?,?,?,?,?,?,?,?,?)
    """, (
        data["ticket_num"], data["phone"], data["city"], data["school"],
        data["serial"], data["description"], data["service"],
        data["service_contact"], data["created_at"], data["tg_user_id"]
    ))
    con.commit()
    row_id = cur.lastrowid
    con.close()
    return row_id

def get_all_tickets():
    con = sqlite3.connect("data/tickets.db")
    cur = con.cursor()
    cur.execute("SELECT * FROM tickets ORDER BY id DESC")
    rows = cur.fetchall()
    con.close()
    return rows

def update_status(ticket_num: str, status: str):
    con = sqlite3.connect("data/tickets.db")
    cur = con.cursor()
    cur.execute("UPDATE tickets SET status=? WHERE ticket_num=?", (status, ticket_num))
    con.commit()
    con.close()

# ─── Маршрутизация по Excel ───────────────────────────────────
def find_service(city: str, serial: str) -> dict:
    """
    Ищет сервисный центр в файле data/services.xlsx
    Колонки: Город | Тип техники | Бренд | Сервис | Контакт
    Если файл не найден — возвращает заглушку.
    """
    path = "data/services.xlsx"
    if not os.path.exists(path):
        return {
            "service": "Сервисный центр (таблица не загружена)",
            "contact": "Не указан"
        }

    wb = openpyxl.load_workbook(path)
    ws = wb.active

    city_lower = city.strip().lower()

    # Сначала ищем по городу (минимум)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        row_city = str(row[0]).strip().lower()
        if row_city in city_lower or city_lower in row_city:
            return {
                "service": str(row[3]) if row[3] else "Не указан",
                "contact": str(row[4]) if row[4] else "Не указан"
            }

    return {
        "service": f"Сервис для города {city} не найден",
        "contact": "Уточните вручную"
    }

# ─── Генерация номера заявки ──────────────────────────────────
def gen_ticket_num() -> str:
    now = datetime.now()
    con = sqlite3.connect("data/tickets.db")
    cur = con.cursor()
    cur.execute("SELECT COUNT(*) FROM tickets")
    count = cur.fetchone()[0] + 1
    con.close()
    return f"ZV-{now.strftime('%Y%m%d')}-{count:04d}"

# ─── Генерация талона (DOCX) ──────────────────────────────────
def generate_tallon(data: dict) -> str:
    doc = Document()

    # Заголовок
    title = doc.add_heading("ТАЛОН НА РЕМОНТ ОБОРУДОВАНИЯ", 0)
    title.alignment = 1  # center

    doc.add_paragraph("")

    # Данные
    fields = [
        ("Номер заявки",        data["ticket_num"]),
        ("Дата",                data["created_at"][:10]),
        ("Телефон заявителя",   data["phone"]),
        ("Город",               data["city"]),
        ("Школа",               data["school"]),
        ("Серийный номер",      data["serial"]),
        ("Описание неисправности", data["description"]),
        ("Сервисный центр",     data["service"]),
        ("Контакт сервиса",     data["service_contact"]),
        ("Статус",              "Новая заявка"),
    ]

    table = doc.add_table(rows=len(fields), cols=2)
    table.style = "Table Grid"

    for i, (label, value) in enumerate(fields):
        row = table.rows[i]
        row.cells[0].text = label
        row.cells[1].text = str(value)
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(11)

    doc.add_paragraph("")
    doc.add_paragraph("Подпись представителя школы: ___________________")
    doc.add_paragraph("Подпись сервисного центра:  ___________________")

    path = f"data/tallon_{data['ticket_num']}.docx"
    doc.save(path)
    return path

# ─── Хэндлеры бота ───────────────────────────────────────────

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "👋 Здравствуйте!\n\n"
        "Это сервис подачи заявок на ремонт оборудования.\n\n"
        "Давайте начнём. Укажите ваш номер телефона:",
        reply_markup=ReplyKeyboardRemove()
    )
    await state.set_state(Form.phone)


@dp.message(Form.phone)
async def got_phone(message: types.Message, state: FSMContext):
    phone = message.text.strip()
    # Простая проверка что похоже на номер
    if len(phone) < 7:
        await message.answer("Пожалуйста, введите корректный номер телефона:")
        return
    await state.update_data(phone=phone)
    await message.answer("Из какого вы города?")
    await state.set_state(Form.city)


@dp.message(Form.city)
async def got_city(message: types.Message, state: FSMContext):
    await state.update_data(city=message.text.strip())
    await message.answer("Укажите название вашей школы:")
    await state.set_state(Form.school)


@dp.message(Form.school)
async def got_school(message: types.Message, state: FSMContext):
    await state.update_data(school=message.text.strip())
    await message.answer(
        "Отправьте серийный номер оборудования.\n\n"
        "Его можно найти на наклейке сзади или снизу устройства."
    )
    await state.set_state(Form.serial)


@dp.message(Form.serial)
async def got_serial(message: types.Message, state: FSMContext):
    await state.update_data(serial=message.text.strip())
    await message.answer("Опишите неисправность. Что случилось с оборудованием?")
    await state.set_state(Form.description)


@dp.message(Form.description)
async def got_description(message: types.Message, state: FSMContext):
    await state.update_data(description=message.text.strip())
    data = await state.get_data()

    # Ищем сервис
    service_info = find_service(data["city"], data["serial"])

    await state.update_data(
        service=service_info["service"],
        service_contact=service_info["contact"]
    )

    # Показываем сводку для подтверждения
    kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Всё верно, отправить")],
            [KeyboardButton(text="❌ Начать заново")]
        ],
        resize_keyboard=True
    )

    await message.answer(
        f"📋 *Проверьте данные заявки:*\n\n"
        f"📞 Телефон: {data['phone']}\n"
        f"🏙 Город: {data['city']}\n"
        f"🏫 Школа: {data['school']}\n"
        f"🔢 Серийный номер: {data['serial']}\n"
        f"🔧 Неисправность: {data['description']}\n\n"
        f"🏢 Сервисный центр: {service_info['service']}\n"
        f"📱 Контакт сервиса: {service_info['contact']}\n\n"
        f"Всё верно?",
        parse_mode="Markdown",
        reply_markup=kb
    )
    await state.set_state(Form.confirm)


@dp.message(Form.confirm, F.text == "✅ Всё верно, отправить")
async def confirmed(message: types.Message, state: FSMContext):
    data = await state.get_data()

    ticket_num  = gen_ticket_num()
    created_at  = datetime.now().strftime("%Y-%m-%d %H:%M")

    full_data = {**data, "ticket_num": ticket_num, "created_at": created_at,
                 "tg_user_id": message.from_user.id}

    # Сохраняем в БД
    save_ticket(full_data)

    # Генерируем талон
    tallon_path = generate_tallon(full_data)

    # Отвечаем школе
    await message.answer(
        f"✅ *Заявка принята!*\n\n"
        f"Номер вашей заявки: *{ticket_num}*\n\n"
        f"Сервисный центр свяжется с вами в ближайшее время.\n"
        f"Сохраните номер заявки для отслеживания.",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )

    # Уведомляем админа
    admin_text = (
        f"🆕 *Новая заявка {ticket_num}*\n\n"
        f"📞 {full_data['phone']}\n"
        f"🏙 {full_data['city']}\n"
        f"🏫 {full_data['school']}\n"
        f"🔢 {full_data['serial']}\n"
        f"🔧 {full_data['description']}\n\n"
        f"🏢 Сервис: {full_data['service']}\n"
        f"📱 Контакт: {full_data['service_contact']}"
    )
    try:
        await bot.send_message(ADMIN_CHAT_ID, admin_text, parse_mode="Markdown")
        await bot.send_document(
            ADMIN_CHAT_ID,
            types.FSInputFile(tallon_path),
            caption=f"Талон — {ticket_num}"
        )
    except Exception as e:
        log.error(f"Не удалось отправить уведомление админу: {e}")

    await state.clear()


@dp.message(Form.confirm, F.text == "❌ Начать заново")
async def restart(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "Хорошо, начнём сначала.\n\nУкажите ваш номер телефона:",
        reply_markup=ReplyKeyboardRemove()
    )
    await state.set_state(Form.phone)


# ─── Команды для админа ───────────────────────────────────────

@dp.message(F.text == "/tickets")
async def admin_tickets(message: types.Message):
    if message.from_user.id != ADMIN_CHAT_ID:
        return
    rows = get_all_tickets()
    if not rows:
        await message.answer("Заявок пока нет.")
        return
    text = "📋 *Последние заявки:*\n\n"
    for r in rows[:10]:
        text += f"*{r[1]}* | {r[4]} | {r[5]} | {r[9]}\n"
    await message.answer(text, parse_mode="Markdown")


@dp.message(F.text.startswith("/status "))
async def admin_set_status(message: types.Message):
    if message.from_user.id != ADMIN_CHAT_ID:
        return
    parts = message.text.split(" ", 2)
    if len(parts) < 3:
        await message.answer("Формат: /status ZV-20240101-0001 готово")
        return
    ticket_num, status = parts[1], parts[2]
    update_status(ticket_num, status)
    await message.answer(f"✅ Статус заявки {ticket_num} обновлён: {status}")


# ─── Запуск ───────────────────────────────────────────────────

async def main():
    init_db()
    log.info("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
