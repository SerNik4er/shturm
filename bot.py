import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from vkbottle.bot import Bot, Message
from vkbottle import Keyboard, KeyboardButtonColor, Text

# ТОКЕН БОТА - ЗАМЕНИТЕ НА СВОЙ!
TOKEN = os.getenv("VK_BOT_TOKEN")

# Создаем бота
bot = Bot(token=TOKEN)

# Файл для хранения данных
DATA_FILE = "users_data.xlsx"

# Временное хранилище данных пользователей (в памяти)
user_data_temp = {}

# Создаем Excel файл если его нет
def init_excel():
    if not os.path.exists(DATA_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Users"
        headers = ["ID", "Имя", "Телефон", "Возраст", "Город", "Дата регистрации"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        wb.save(DATA_FILE)

# Сохраняем данные пользователя
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

# Клавиатура для главного меню
def get_main_keyboard():
    keyboard = Keyboard(inline=False)
    keyboard.add(Text("📝 Регистрация"), color=KeyboardButtonColor.PRIMARY)
    keyboard.row()
    keyboard.add(Text("📊 Посмотреть данные"), color=KeyboardButtonColor.SECONDARY)
    keyboard.row()
    keyboard.add(Text("❓ Помощь"), color=KeyboardButtonColor.SECONDARY)
    return keyboard

# Клавиатура для отмены
def get_cancel_keyboard():
    keyboard = Keyboard(inline=False)
    keyboard.add(Text("❌ Отмена"), color=KeyboardButtonColor.NEGATIVE)
    return keyboard

# Инициализация Excel файла
init_excel()

# Команда старт
@bot.on.private_message(text=["начать", "старт", "start", "/start"])
async def start_command(message: Message):
    await message.answer(
        "👋 Привет! Я бот для сбора данных.\n\nВыберите действие:",
        keyboard=get_main_keyboard()
    )

# Обработчик для главного меню
@bot.on.private_message(text="📝 Регистрация")
async def register_start(message: Message):
    # Очищаем предыдущие данные пользователя
    if message.peer_id in user_data_temp:
        del user_data_temp[message.peer_id]
    
    await message.answer(
        "Пожалуйста, введите ваше имя:",
        keyboard=get_cancel_keyboard()
    )
    # Сохраняем состояние в словарь
    user_data_temp[message.peer_id] = {"state": "waiting_for_name"}

# Обработчик для просмотра данных
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
    
    user_data = None
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == message.from_id:
            user_data = {
                "name": ws.cell(row=row, column=2).value,
                "phone": ws.cell(row=row, column=3).value,
                "age": ws.cell(row=row, column=4).value,
                "city": ws.cell(row=row, column=5).value,
                "date": ws.cell(row=row, column=6).value
            }
            break
    
    if user_data:
        response = f"📋 Ваши данные:\n\n"
        response += f"👤 Имя: {user_data['name']}\n"
        response += f"📱 Телефон: {user_data['phone']}\n"
        response += f"🎂 Возраст: {user_data['age']}\n"
        response += f"🏙️ Город: {user_data['city']}\n"
        response += f"📅 Дата: {user_data['date']}"
        await message.answer(response, keyboard=get_main_keyboard())
    else:
        await message.answer(
            "❌ Вы ещё не зарегистрированы!\nНажмите '📝 Регистрация'",
            keyboard=get_main_keyboard()
        )

# Обработчик помощи
@bot.on.private_message(text="❓ Помощь")
async def help_command(message: Message):
    await message.answer(
        "🤖 Инструкция по использованию:\n\n"
        "1. Нажмите '📝 Регистрация' для заполнения анкеты\n"
        "2. Введите запрашиваемые данные\n"
        "3. Нажмите '📊 Посмотреть данные' для просмотра вашей анкеты\n"
        "4. Все данные сохраняются в Excel файл\n\n"
        "Команды:\n"
        "📝 Регистрация - начать регистрацию\n"
        "📊 Посмотреть данные - показать ваши данные\n"
        "❓ Помощь - показать эту справку",
        keyboard=get_main_keyboard()
    )

# Обработчик отмены
@bot.on.private_message(text="❌ Отмена")
async def cancel_action(message: Message):
    if message.peer_id in user_data_temp:
        del user_data_temp[message.peer_id]
    await message.answer(
        "✅ Действие отменено",
        keyboard=get_main_keyboard()
    )

# ГЛАВНЫЙ ОБРАБОТЧИК - проверяем состояние вручную
@bot.on.private_message()
async def handle_all_messages(message: Message):
    # Проверяем, есть ли состояние для этого пользователя
    if message.peer_id not in user_data_temp:
        # Если нет состояния и это не специальная команда
        await message.answer(
            "❓ Неизвестная команда.\nИспользуйте кнопки для навигации.",
            keyboard=get_main_keyboard()
        )
        return
    
    current_state = user_data_temp[message.peer_id].get("state")
    
    # Обработка состояния "waiting_for_name"
    if current_state == "waiting_for_name":
        name = message.text.strip()
        if len(name) < 2:
            await message.answer("❌ Имя слишком короткое. Попробуйте ещё раз:")
            return
        
        user_data_temp[message.peer_id]["name"] = name
        user_data_temp[message.peer_id]["state"] = "waiting_for_phone"
        
        await message.answer(
            "📱 Введите ваш номер телефона:",
            keyboard=get_cancel_keyboard()
        )
    
    # Обработка состояния "waiting_for_phone"
    elif current_state == "waiting_for_phone":
        phone = message.text.strip()
        if len(phone) < 10:
            await message.answer("❌ Некорректный номер. Попробуйте ещё раз:")
            return
        
        user_data_temp[message.peer_id]["phone"] = phone
        user_data_temp[message.peer_id]["state"] = "waiting_for_age"
        
        await message.answer(
            "🎂 Введите ваш возраст:",
            keyboard=get_cancel_keyboard()
        )
    
    # Обработка состояния "waiting_for_age"
    elif current_state == "waiting_for_age":
        age_text = message.text.strip()
        try:
            age = int(age_text)
            if age < 1 or age > 150:
                raise ValueError
        except ValueError:
            await message.answer("❌ Пожалуйста, введите корректный возраст (число от 1 до 150):")
            return
        
        user_data_temp[message.peer_id]["age"] = age
        user_data_temp[message.peer_id]["state"] = "waiting_for_city"
        
        await message.answer(
            "🏙️ Введите ваш город:",
            keyboard=get_cancel_keyboard()
        )
    
    # Обработка состояния "waiting_for_city"
    elif current_state == "waiting_for_city":
        city = message.text.strip()
        if len(city) < 2:
            await message.answer("❌ Название города слишком короткое. Попробуйте ещё раз:")
            return
        
        # Получаем все данные
        data = user_data_temp[message.peer_id]
        name = data.get("name", "Не указано")
        phone = data.get("phone", "Не указано")
        age = data.get("age", "Не указано")
        
        # Сохраняем в Excel
        save_user_data(message.from_id, name, phone, age, city)
        
        # Очищаем временные данные
        del user_data_temp[message.peer_id]
        
        await message.answer(
            f"✅ Регистрация завершена!\n\n"
            f"Ваши данные:\n"
            f"👤 Имя: {name}\n"
            f"📱 Телефон: {phone}\n"
            f"🎂 Возраст: {age}\n"
            f"🏙️ Город: {city}\n\n"
            "Спасибо за регистрацию! 🎉",
            keyboard=get_main_keyboard()
        )
    
    else:
        # Если состояние не распознано
        await message.answer(
            "❓ Произошла ошибка. Начните заново с кнопки '📝 Регистрация'",
            keyboard=get_main_keyboard()
        )

# Запуск бота
if __name__ == "__main__":
    print("🤖 Бот запущен!")
    print("✅ Нажмите Ctrl+C для остановки")
    bot.run()
