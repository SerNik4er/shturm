import os
import asyncio
import logging
from datetime import datetime
from openpyxl import Workbook, load_workbook
from vkbottle.bot import Bot, Message
from vkbottle import Keyboard, KeyboardButtonColor, Text
from dotenv import load_dotenv
load_dotenv()

# Отключаем логирование
logging.getLogger("vkbottle").setLevel(logging.CRITICAL)
logging.getLogger("loguru").setLevel(logging.CRITICAL)

# ТОКЕН ИЗ ПЕРЕМЕННОЙ ОКРУЖЕНИЯ
TOKEN = os.getenv("VK_BOT_TOKEN")

if not TOKEN:
    raise ValueError("❌ Токен не найден! Установите VK_BOT_TOKEN")

# ID админа (кто может использовать !список и !очистить)
ADMIN_IDS = [164876852]  # ← ЗАМЕНИТЕ НА СВОЙ VK ID!

# ===== СПИСОК ГРУПП ДЛЯ ПРОВЕРКИ ПОДПИСКИ =====
# Укажите ID групп, на которые нужно подписаться
REQUIRED_GROUPS = [
    237204348,  # Группа 1 (замените на свой ID)
    108117049,  # Группа "Штурмовой бой для детей | Калуга"
    # Добавьте сколько нужно групп
]

# Текст для неподписанных пользователей
NOT_SUBSCRIBED_TEXT = (
    "❌ Вы не подписаны на обязательные группы!\n\n"
    "Чтобы продолжить, подпишитесь на все группы:\n"
)
# ===== КОНЕЦ НАСТРОЕК =====

# Отключаем лишние логи
logging.getLogger("vkbottle").setLevel(logging.CRITICAL)
logging.getLogger("loguru").setLevel(logging.CRITICAL)

bot = Bot(token=TOKEN)

DATA_FILE = "users_data.xlsx"
user_data_temp = {}


# ===== ФУНКЦИИ РАБОТЫ С EXCEL =====

def init_excel():
    """Создает Excel-файл с заголовками, если его нет"""
    if not os.path.exists(DATA_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Users"
        headers = ["ID", "Имя", "Телефон", "Возраст", "Город", "Дата регистрации"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        wb.save(DATA_FILE)


def save_user_data(user_id, name, phone, age, city):
    """Сохраняет данные пользователя в Excel"""
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


# ===== КЛАВИАТУРЫ =====

def get_main_keyboard():
    """Главное меню"""
    keyboard = Keyboard(inline=False)
    keyboard.add(Text("📝 Регистрация"), color=KeyboardButtonColor.PRIMARY)
    keyboard.row()
    keyboard.add(Text("📊 Посмотреть данные"), color=KeyboardButtonColor.SECONDARY)
    keyboard.row()
    keyboard.add(Text("❓ Помощь"), color=KeyboardButtonColor.SECONDARY)
    return keyboard


def get_cancel_keyboard():
    """Клавиатура с кнопкой отмены"""
    keyboard = Keyboard(inline=False)
    keyboard.add(Text("❌ Отмена"), color=KeyboardButtonColor.NEGATIVE)
    return keyboard


def get_subscription_keyboard():
    """Клавиатура для проверки подписки"""
    keyboard = Keyboard(inline=False)
    keyboard.add(Text("✅ Проверить подписку"), color=KeyboardButtonColor.POSITIVE)
    keyboard.row()
    keyboard.add(Text("❌ Отмена"), color=KeyboardButtonColor.NEGATIVE)
    return keyboard


# ===== ПРОВЕРКА ПОДПИСКИ НА НЕСКОЛЬКО ГРУПП =====

async def check_subscription(user_id: int) -> tuple:
    """
    Проверяет, подписан ли пользователь на все группы из списка.
    Возвращает: (все_подписаны, список_неподписанных_групп)
    """
    if not REQUIRED_GROUPS:
        return True, []

    not_subscribed = []

    for group_id in REQUIRED_GROUPS:
        try:
            response = await bot.api.request(
                "groups.isMember",
                {
                    "group_id": group_id,
                    "user_id": user_id,
                    "v": "5.131"
                }
            )
            if response.get("response") != 1:
                not_subscribed.append(group_id)
        except Exception as e:
            logging.error(f"Ошибка проверки подписки на группу {group_id}: {e}")
            not_subscribed.append(group_id)

    return len(not_subscribed) == 0, not_subscribed


def get_groups_links(groups_ids):
    """Формирует список ссылок на группы"""
    return "\n".join([f"👉 https://vk.com/club{g}" for g in groups_ids])


# ===== ОБЫЧНЫЕ КОМАНДЫ =====

@bot.on.private_message(text=["начать", "старт", "start", "/start"])
async def start_command(message: Message):
    """Стартовая команда с проверкой подписки"""
    is_subscribed, not_subscribed = await check_subscription(message.from_id)

    if not is_subscribed:
        await message.answer(
            NOT_SUBSCRIBED_TEXT + get_groups_links(not_subscribed) +
            "\n\nПодпишитесь и нажмите '✅ Проверить подписку'",
            keyboard=get_subscription_keyboard()
        )
        return

    await message.answer(
        "👋 Привет! Я бот для сбора данных.\n\nВыберите действие:",
        keyboard=get_main_keyboard()
    )


@bot.on.private_message(text="✅ Проверить подписку")
async def check_subscription_button(message: Message):
    """Кнопка проверки подписки"""
    is_subscribed, not_subscribed = await check_subscription(message.from_id)

    if is_subscribed:
        await message.answer(
            "✅ Отлично! Вы подписаны на все группы.\n\nТеперь вам доступны все функции:",
            keyboard=get_main_keyboard()
        )
    else:
        await message.answer(
            NOT_SUBSCRIBED_TEXT + get_groups_links(not_subscribed) +
            "\n\nПодпишитесь и нажмите кнопку снова",
            keyboard=get_subscription_keyboard()
        )


@bot.on.private_message(text="📝 Регистрация")
async def register_start(message: Message):
    """Начинает процесс регистрации"""
    # Проверяем подписку
    is_subscribed, not_subscribed = await check_subscription(message.from_id)
    if not is_subscribed:
        await message.answer(
            NOT_SUBSCRIBED_TEXT + get_groups_links(not_subscribed) +
            "\n\nПодпишитесь и нажмите '✅ Проверить подписку'",
            keyboard=get_subscription_keyboard()
        )
        return

    if message.peer_id in user_data_temp:
        del user_data_temp[message.peer_id]

    await message.answer(
        "Пожалуйста, введите ваше имя:",
        keyboard=get_cancel_keyboard()
    )
    user_data_temp[message.peer_id] = {"state": "waiting_for_name"}


@bot.on.private_message(text="📊 Посмотреть данные")
async def view_data(message: Message):
    """Показывает данные пользователя"""
    # Проверяем подписку
    is_subscribed, not_subscribed = await check_subscription(message.from_id)
    if not is_subscribed:
        await message.answer(
            NOT_SUBSCRIBED_TEXT + get_groups_links(not_subscribed) +
            "\n\nПодпишитесь и нажмите '✅ Проверить подписку'",
            keyboard=get_subscription_keyboard()
        )
        return

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

    await message.answer(
        "❌ Вы ещё не зарегистрированы!\nНажмите '📝 Регистрация'",
        keyboard=get_main_keyboard()
    )


@bot.on.private_message(text="❓ Помощь")
async def help_command(message: Message):
    """Показывает справку"""
    await message.answer(
        "🤖 Инструкция:\n\n"
        "📝 Регистрация - заполнить анкету\n"
        "📊 Посмотреть данные - показать вашу анкету\n"
        "❌ Отмена - отменить текущее действие\n\n"
        "⚠️ Для использования бота нужно быть подписанным на все группы!\n\n"
        "👑 Админ-команды:\n"
        "!список - показать всех пользователей\n"
        "!очистить - удалить все данные",
        keyboard=get_main_keyboard()
    )


@bot.on.private_message(text="❌ Отмена")
async def cancel_action(message: Message):
    """Отменяет текущее действие"""
    if message.peer_id in user_data_temp:
        del user_data_temp[message.peer_id]
    await message.answer("✅ Действие отменено", keyboard=get_main_keyboard())


# ===== АДМИН-КОМАНДЫ =====

@bot.on.private_message(text="!список")
async def admin_list_users(message: Message):
    """Отправляет список зарегистрированных пользователей в виде текстового файла"""
    # Проверка прав админа
    if message.from_id not in ADMIN_IDS:
        await message.answer("⛔ У вас нет прав для этой команды.")
        return
    
    if not os.path.exists(DATA_FILE):
        await message.answer("📭 Данных пока нет.")
        return
    
    wb = load_workbook(DATA_FILE)
    ws = wb.active
    
    if ws.max_row == 1:
        await message.answer("📭 Данных пока нет.")
        return
    
    # Создаём текстовый файл со списком
    import tempfile
    temp_file = tempfile.NamedTemporaryFile(suffix='.txt', mode='w', encoding='utf-8', delete=False)
    
    # Записываем заголовок
    temp_file.write("=" * 50 + "\n")
    temp_file.write("СПИСОК ЗАРЕГИСТРИРОВАННЫХ ПОЛЬЗОВАТЕЛЕЙ\n")
    temp_file.write("=" * 50 + "\n\n")
    
    total_users = 0
    for row in range(2, ws.max_row + 1):
        user_id = ws.cell(row=row, column=1).value
        name = ws.cell(row=row, column=2).value
        phone = ws.cell(row=row, column=3).value
        age = ws.cell(row=row, column=4).value
        city = ws.cell(row=row, column=5).value
        date = ws.cell(row=row, column=6).value
        
        temp_file.write(f"#{row-1}\n")
        temp_file.write(f"👤 Имя: {name}\n")
        temp_file.write(f"🆔 VK ID: {user_id}\n")
        temp_file.write(f"📱 Телефон: {phone}\n")
        temp_file.write(f"🎂 Возраст: {age} лет\n")
        temp_file.write(f"🏙️ Город: {city}\n")
        temp_file.write(f"📅 Дата регистрации: {date}\n")
        temp_file.write("─" * 30 + "\n\n")
        total_users += 1
    
    temp_file.write(f"\n{'=' * 30}\n")
    temp_file.write(f"Всего пользователей: {total_users}\n")
    temp_file.close()
    
    # Отправляем файл через VK API напрямую
    try:
        # Открываем файл для чтения
        with open(temp_file.name, 'rb') as f:
            file_data = f.read()
        
        # Получаем сервер для загрузки документа
        upload_server = await bot.api.request(
            "docs.getUploadServer",
            {"type": "doc", "peer_id": message.peer_id}
        )
        
        # Загружаем файл
        import aiohttp
        async with aiohttp.ClientSession() as session:
            form_data = aiohttp.FormData()
            form_data.add_field(
                'file',
                file_data,
                filename='users_list.txt',
                content_type='text/plain'
            )
            async with session.post(
                upload_server['response']['upload_url'],
                data=form_data
            ) as response:
                upload_result = await response.json()
        
        # Сохраняем документ
        save_result = await bot.api.request(
            "docs.save",
            {
                "file": upload_result['file'],
                "title": "users_list.txt"
            }
        )
        
        # Получаем attachment
        doc = save_result['response'][0]
        attachment = f"doc{doc['owner_id']}_{doc['id']}"
        
        await message.answer(
            f"📊 Всего зарегистрировано: {total_users}\n\n📎 Файл с полным списком прикреплён ниже.",
            attachment=attachment
        )
        
    except Exception as e:
        # Если файл не отправился - показываем первые 10 пользователей текстом
        await message.answer(f"❌ Не удалось отправить файл.\n\nПоказываю первых 10 пользователей:")
        
        users = []
        for row in range(2, min(ws.max_row + 1, 12)):
            name = ws.cell(row=row, column=2).value
            phone = ws.cell(row=row, column=3).value
            age = ws.cell(row=row, column=4).value
            city = ws.cell(row=row, column=5).value
            date = ws.cell(row=row, column=6).value
            users.append(
                f"👤 {name}\n📱 {phone}\n🎂 {age} лет\n🏙️ {city}\n📅 {date}\n{'─'*15}"
            )
        
        if users:
            await message.answer("\n\n".join(users))
        else:
            await message.answer("📭 Данных пока нет.")
    
    finally:
        # Удаляем временный файл
        try:
            os.unlink(temp_file.name)
        except:
            pass


@bot.on.private_message(text="!очистить")
async def admin_clear_users(message: Message):
    """Очищает все данные (с подтверждением)"""
    if message.from_id not in ADMIN_IDS:
        await message.answer("⛔ У вас нет прав для этой команды.")
        return

    if not os.path.exists(DATA_FILE):
        await message.answer("📭 Данных пока нет.")
        return

    await message.answer(
        "⚠️ ВНИМАНИЕ! Вы уверены, что хотите удалить ВСЕ данные?\n\n"
        "Для подтверждения напишите: **да, очистить**\n"
        "Для отмены напишите: **отмена**"
    )
    user_data_temp[message.peer_id] = {"state": "waiting_for_clear"}


# ===== ГЛАВНЫЙ ОБРАБОТЧИК СОСТОЯНИЙ =====

@bot.on.private_message()
async def handle_all_messages(message: Message):
    """Обрабатывает все остальные сообщения (анкетирование, подтверждения)"""

    # Если нет состояния - показываем главное меню
    if message.peer_id not in user_data_temp:
        await message.answer(
            "❓ Неизвестная команда.\nИспользуйте кнопки.",
            keyboard=get_main_keyboard()
        )
        return

    current_state = user_data_temp[message.peer_id].get("state")

    # ===== ПОДТВЕРЖДЕНИЕ ОЧИСТКИ =====
    if current_state == "waiting_for_clear":
        text = message.text.strip().lower()

        if text == "да, очистить":
            if os.path.exists(DATA_FILE):
                os.remove(DATA_FILE)
            init_excel()
            await message.answer("✅ Все данные успешно очищены!")
            del user_data_temp[message.peer_id]
            return

        elif text == "отмена":
            await message.answer("❌ Очистка отменена.")
            del user_data_temp[message.peer_id]
            return

        else:
            await message.answer(
                "❌ Неверная команда.\n"
                "Для подтверждения напишите: **да, очистить**\n"
                "Для отмены напишите: **отмена**"
            )
            return

    # ===== АНКЕТИРОВАНИЕ =====
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
            "✅ Регистрация завершена! 🎉\n\n"
            "Теперь вы можете посмотреть свои данные через '📊 Посмотреть данные'",
            keyboard=get_main_keyboard()
        )


# ===== ЗАПУСК =====

if __name__ == "__main__":
    # Создаём Excel-файл при старте
    init_excel()

    print("🤖 Бот запущен!")
    print(f"✅ Проверяются группы: {REQUIRED_GROUPS}")
    print("✅ Нажмите Ctrl+C для остановки")
    asyncio.run(bot.run_polling())
