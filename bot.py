import os
import asyncio
import logging
from datetime import datetime
from openpyxl import Workbook, load_workbook
from vkbottle.bot import Bot, Message
from vkbottle import Keyboard, KeyboardButtonColor, Text
from dotenv import load_dotenv

# Загружаем переменные из .env
load_dotenv()

# Отключаем логирование vkbottle (решает проблему с logger.opt)
logging.getLogger("vkbottle").setLevel(logging.CRITICAL)
logging.getLogger("loguru").setLevel(logging.CRITICAL)

# ТОКЕН ИЗ ПЕРЕМЕННОЙ ОКРУЖЕНИЯ
TOKEN = os.getenv("VK_BOT_TOKEN")

if not TOKEN:
    raise ValueError("❌ Токен не найден! Установите VK_BOT_TOKEN")

bot = Bot(token=TOKEN)

DATA_FILE = "users_data.xlsx"
user_data_temp = {}

def init_excel():
    if not os.path.exists(DATA_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Users"
        headers = ["ID", "Имя", "Телефон", "Возраст", "Город", "Дата регистрации"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        wb.save(DATA_FILE)

def save_user_data(user_id, name, phone, age, city):
    wb = load_workbook(DATA_FILE)
    ws = wb.active
    row = ws.max_row + 1
    ws.cell(row=row, column=1, value=user_id)
    ws.cell(row=row, column=2, value=name)
    ws.cell(row=row, column=3, value=phone)
    ws.cell(row=row, column=4, value=age)
    ws.cell(row=row, column=5, value=city)
    ws.cell(row=row, column=6, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    wb.save(DATA_FILE)

def get_main_keyboard():
    keyboard = Keyboard(inline=False)
    keyboard.add(Text("📝 Регистрация"), color=KeyboardButtonColor.PRIMARY)
    keyboard.row()
    keyboard.add(Text("📊 Посмотреть данные"), color=KeyboardButtonColor.SECONDARY)
    keyboard.row()
    keyboard.add(Text("❓ Помощь"), color=KeyboardButtonColor.SECONDARY)
    return keyboard

def get_cancel_keyboard():
    keyboard = Keyboard(inline=False)
    keyboard.add(Text("❌ Отмена"), color=KeyboardButtonColor.NEGATIVE)
    return keyboard

init_excel()

@bot.on.private_message(text=["начать", "старт", "start", "/start"])
async def start_command(message: Message):
    await message.answer(
        "👋 Привет! Я бот для сбора данных.\n\nВыберите действие:",
        keyboard=get_main_keyboard()
    )

@bot.on.private_message(text="📝 Регистрация")
async def register_start(message: Message):
    if message.peer_id in user_data_temp:
        del user_data_temp[message.peer_id]
    await message.answer(
        "Пожалуйста, введите ваше имя:",
        keyboard=get_cancel_keyboard()
    )
    user_data_temp[message.peer_id] = {"state": "waiting_for_name"}

@bot.on.private_message(text="📊 Посмотреть данные")
async def view_data(message: Message):
    if not os.path.exists(DATA_FILE):
        await message.answer("📭 Данных пока нет.")
        return
    wb = load_workbook(DATA_FILE)
    ws = wb.active
    if ws.max_row == 1:
        await message.answer("📭 Данных пока нет.")
        return
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == message.from_id:
            await message.answer(
                f"📋 Ваши данные:\n\n"
                f"👤 Имя: {ws.cell(row=row, column=2).value}\n"
                f"📱 Телефон: {ws.cell(row=row, column=3).value}\n"
                f"🎂 Возраст: {ws.cell(row=row, column=4).value}\n"
                f"🏙️ Город: {ws.cell(row=row, column=5).value}",
                keyboard=get_main_keyboard()
            )
            return
    await message.answer("❌ Вы ещё не зарегистрированы!", keyboard=get_main_keyboard())

@bot.on.private_message(text="❓ Помощь")
async def help_command(message: Message):
    await message.answer(
        "🤖 Инструкция:\n\n"
        "1. '📝 Регистрация' - заполнить анкету\n"
        "2. '📊 Посмотреть данные' - показать анкету\n"
        "3. '❌ Отмена' - отменить действие",
        keyboard=get_main_keyboard()
    )

@bot.on.private_message(text="❌ Отмена")
async def cancel_action(message: Message):
    if message.peer_id in user_data_temp:
        del user_data_temp[message.peer_id]
    await message.answer("✅ Действие отменено", keyboard=get_main_keyboard())

@bot.on.private_message()
async def handle_all_messages(message: Message):
    if message.peer_id not in user_data_temp:
        await message.answer(
            "❓ Неизвестная команда.\nИспользуйте кнопки.",
            keyboard=get_main_keyboard()
        )
        return
    
    current_state = user_data_temp[message.peer_id].get("state")
    
    if current_state == "waiting_for_name":
        name = message.text.strip()
        if len(name) < 2:
            await message.answer("❌ Имя слишком короткое. Попробуйте ещё раз:")
            return
        user_data_temp[message.peer_id]["name"] = name
        user_data_temp[message.peer_id]["state"] = "waiting_for_phone"
        await message.answer("📱 Введите ваш номер телефона:", keyboard=get_cancel_keyboard())
    
    elif current_state == "waiting_for_phone":
        phone = message.text.strip()
        if len(phone) < 10:
            await message.answer("❌ Некорректный номер. Попробуйте ещё раз:")
            return
        user_data_temp[message.peer_id]["phone"] = phone
        user_data_temp[message.peer_id]["state"] = "waiting_for_age"
        await message.answer("🎂 Введите ваш возраст:", keyboard=get_cancel_keyboard())
    
    elif current_state == "waiting_for_age":
        try:
            age = int(message.text.strip())
            if age < 1 or age > 150:
                raise ValueError
        except ValueError:
            await message.answer("❌ Введите число от 1 до 150:")
            return
        user_data_temp[message.peer_id]["age"] = age
        user_data_temp[message.peer_id]["state"] = "waiting_for_city"
        await message.answer("🏙️ Введите ваш город:", keyboard=get_cancel_keyboard())
    
    elif current_state == "waiting_for_city":
        city = message.text.strip()
        if len(city) < 2:
            await message.answer("❌ Название города слишком короткое. Попробуйте ещё раз:")
            return
        data = user_data_temp[message.peer_id]
        save_user_data(
            message.from_id,
            data.get("name", "Не указано"),
            data.get("phone", "Не указано"),
            data.get("age", "Не указано"),
            city
        )
        del user_data_temp[message.peer_id]
        await message.answer(
            "✅ Регистрация завершена! 🎉",
            keyboard=get_main_keyboard()
        )

# ===== ВАЖНО: ИСПРАВЛЕННЫЙ ЗАПУСК =====
if __name__ == "__main__":
    print("🤖 Бот запущен!")
    print("✅ Нажмите Ctrl+C для остановки")
    asyncio.run(bot.run_polling())
