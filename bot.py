import asyncio
import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from vkbottle import Bot
from vkbottle.bot import Message
from vkbottle.tools import DocMessagesUploader

# --- НАСТРОЙКИ ---
BOT_TOKEN = ""

# Список групп для обязательной проверки подписки
# Формат: {"id группы без минуса": "Название или ссылка"}
REQUIRED_GROUPS = {
    123456789: "Основная группа",
    987654321: "Группа партнёра",
    111222333: "Спонсор мероприятия"
}

ADMIN_IDS = [123456789]  # ID администраторов
# ------------------

bot = Bot(token=BOT_TOKEN)

# База зарегистрированных
registered_users_db = {}
# Состояния пользователей
user_states = {}
user_last_active = {}
STATE_TIMEOUT_MINUTES = 10

# Счетчик регистраций
registration_counter = 0


def get_groups_text() -> str:
    """Формирует текст со ссылками на все обязательные группы."""
    lines = []
    for idx, (group_id, name) in enumerate(REQUIRED_GROUPS.items(), 1):
        lines.append(f"{idx}. [{name}](https://vk.com/club{group_id})")
    return "\n".join(lines)


async def check_all_subscriptions(user_id: int) -> dict:
    """
    Проверяет подписку на все обязательные группы.
    Возвращает словарь: {group_id: bool} (True = подписан).
    """
    results = {}
    for group_id in REQUIRED_GROUPS.keys():
        try:
            response = await bot.api.groups.is_member(
                group_id=group_id,
                user_id=user_id,
                extended=False
            )
            results[group_id] = (response == 1)
        except Exception as e:
            print(f"Ошибка проверки группы {group_id}: {e}")
            results[group_id] = False
    return results


def get_unsubscribed_groups(check_results: dict) -> list:
    """Возвращает список групп, на которые пользователь не подписан."""
    unsubscribed = []
    for group_id, is_member in check_results.items():
        if not is_member:
            group_name = REQUIRED_GROUPS[group_id]
            unsubscribed.append(f"👉 [{group_name}](https://vk.com/club{group_id})")
    return unsubscribed


async def is_admin(user_id: int) -> bool:
    """Проверка прав администратора."""
    return user_id in ADMIN_IDS


async def generate_excel() -> io.BytesIO:
    """Создаёт Excel-файл со списком всех зарегистрированных."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Участники"

    # Заголовки
    ws.append(["№", "ID пользователя", "Имя", "Контакты", "Дата регистрации"])

    # Данные
    for idx, (user_id, data) in enumerate(registered_users_db.items(), 1):
        ws.append([
            idx,
            user_id,
            data["name"],
            data["contact"],
            data["timestamp"]
        ])

    # Автоширина колонок
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    # Сохраняем в буфер
    file_buffer = io.BytesIO()
    wb.save(file_buffer)
    file_buffer.seek(0)
    file_buffer.name = f"participants_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return file_buffer


def update_user_activity(user_id: int):
    """Обновляет время последней активности пользователя."""
    user_last_active[user_id] = datetime.now()


async def cleanup_expired_states():
    """Удаляет просроченные состояния."""
    now = datetime.now()
    expired = [
        uid for uid, last_time in user_last_active.items()
        if now - last_time > timedelta(minutes=STATE_TIMEOUT_MINUTES)
    ]
    for uid in expired:
        user_states.pop(uid, None)
        user_last_active.pop(uid, None)


# --- ОБРАБОТЧИКИ ---

@bot.on.private_message(text=["начать", "старт", "Начать", "Старт"])
async def start_handler(message: Message):
    """Начало регистрации — показываем список групп и запрашиваем имя."""
    await cleanup_expired_states()
    user_id = message.from_id
    update_user_activity(user_id)

    groups_text = get_groups_text()

    user_states[user_id] = {"state": "waiting_for_name"}

    await message.answer(
        "👋 Добро пожаловать на регистрацию!\n\n"
        "📌 **Для участия необходимо быть подписанным на следующие группы:**\n\n"
        f"{groups_text}\n\n"
        "Пожалуйста, подпишитесь на все группы, если ещё не сделали этого.\n"
        "А сейчас напишите ваше **Имя и Фамилию**."
    )


@bot.on.private_message(state="waiting_for_name")
async def get_name(message: Message):
    """Получаем имя."""
    user_id = message.from_id
    update_user_activity(user_id)

    full_name = message.text.strip()

    if len(full_name) < 3:
        await message.answer("⚠️ Пожалуйста, введите корректные Имя и Фамилию (минимум 3 символа).")
        return

    user_states[user_id] = {
        "name": full_name,
        "state": "waiting_for_contact"
    }

    await message.answer(
        f"Спасибо, {full_name}!\n"
        "Теперь напишите ваш **Email** или **номер телефона** для связи."
    )


@bot.on.private_message(state="waiting_for_contact")
async def get_contact_and_check(message: Message):
    """Получаем контакт и проверяем все подписки."""
    user_id = message.from_id
    update_user_activity(user_id)

    contact = message.text.strip()

    if len(contact) < 5:
        await message.answer("⚠️ Пожалуйста, введите корректный email или телефон (минимум 5 символов).")
        return

    name = user_states[user_id]["name"]

    # Сохраняем контакт в состоянии
    user_states[user_id]["contact"] = contact

    await message.answer("⏳ Проверяем ваши подписки на все необходимые группы...")

    # Проверяем подписку на ВСЕ группы
    check_results = await check_all_subscriptions(user_id)
    unsubscribed = get_unsubscribed_groups(check_results)

    if not unsubscribed:
        # Все подписки есть — регистрируем
        global registration_counter
        registration_counter += 1

        registered_users_db[user_id] = {
            "name": name,
            "contact": contact,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")
        }

        del user_states[user_id]

        await message.answer(
            f"✅ Отлично! Вы успешно зарегистрированы на мероприятие!\n\n"
            f"📋 Ваши данные:\n"
            f"👤 Имя: {name}\n"
            f"📞 Контакты: {contact}\n\n"
            f"Ждём вас!\n"
            f"Если нужно отменить регистрацию, напишите 'Отмена'."
        )

        # Уведомление админам
        for admin_id in ADMIN_IDS:
            try:
                await bot.api.messages.send(
                    user_id=admin_id,
                    random_id=0,
                    message=f"📢 Новая регистрация!\n"
                            f"👤 {name}\n"
                            f"📞 {contact}\n"
                            f"🆔 ID: {user_id}\n"
                            f"📊 Всего: {registration_counter}"
                )
            except Exception as e:
                print(f"Ошибка уведомления админа {admin_id}: {e}")
    else:
        # Есть невыполненные подписки
        unsub_text = "\n".join(unsubscribed)
        await message.answer(
            "❌ **Вы подписаны не на все необходимые группы.**\n\n"
            f"Осталось подписаться на:\n{unsub_text}\n\n"
            "Как подпишетесь — напишите **Проверить**."
        )


@bot.on.private_message(text=["проверить", "Проверить"])
async def recheck_handler(message: Message):
    """Повторная проверка всех подписок."""
    user_id = message.from_id
    update_user_activity(user_id)

    if user_id not in user_states or "contact" not in user_states[user_id]:
        await message.answer(
            "Вы ещё не завершили ввод данных. Давайте начнём заново — напишите **Начать**."
        )
        return

    await message.answer("⏳ Проверяем подписки повторно...")

    check_results = await check_all_subscriptions(user_id)
    unsubscribed = get_unsubscribed_groups(check_results)

    if not unsubscribed:
        name = user_states[user_id]["name"]
        contact = user_states[user_id]["contact"]

        global registration_counter
        registration_counter += 1

        registered_users_db[user_id] = {
            "name": name,
            "contact": contact,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")
        }

        del user_states[user_id]

        await message.answer(
            f"✅ Теперь всё в порядке! Вы зарегистрированы.\n\n"
            f"👤 Имя: {name}\n"
            f"📞 Контакты: {contact}\n\n"
            f"Хорошего мероприятия!"
        )
    else:
        unsub_text = "\n".join(unsubscribed)
        await message.answer(
            f"❌ Вы всё ещё не подписаны на:\n{unsub_text}\n\n"
            f"Подпишитесь и напишите **Проверить** снова."
        )


@bot.on.private_message(text=["отмена", "Отмена"])
async def cancel_handler(message: Message):
    """Отмена регистрации."""
    user_id = message.from_id
    update_user_activity(user_id)

    if user_id in user_states:
        del user_states[user_id]
        if user_id in registered_users_db:
            del registered_users_db[user_id]
        await message.answer("🚫 Регистрация отменена. Напишите **Начать** для повторной регистрации.")
    else:
        await message.answer("Вы не начинали регистрацию.")


# --- КОМАНДЫ АДМИНИСТРАТОРА ---

@bot.on.private_message(text=["список", "Список", "/list"])
async def admin_list_handler(message: Message):
    """Админ-команда: получить Excel-файл со всеми участниками."""
    user_id = message.from_id

    if not await is_admin(user_id):
        await message.answer("⛔ У вас нет доступа к этой команде.")
        return

    if not registered_users_db:
        await message.answer("📭 Список зарегистрированных пока пуст.")
        return

    total = len(registered_users_db)
    await message.answer(f"📊 Формирую файл. Всего участников: {total}. Подождите немного...")

    try:
        # Генерируем Excel
        excel_file = await generate_excel()

        # Загружаем как документ
        uploader = DocMessagesUploader(bot.api)
        doc = await uploader.upload(
            file_source=excel_file,
            title=f"Участники ({total})",
            peer_id=user_id
        )

        # Отправляем файл
        await message.answer(
            f"📋 **Список участников**\n"
            f"Всего зарегистрировано: **{total}**\n"
            f"Дата выгрузки: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
            f"Файл во вложении 👇",
            attachment=doc
        )

    except Exception as e:
        print(f"Ошибка при отправке файла: {e}")

        # Если файл не отправился — выводим текстом (обрезая)
        message_text = f"📋 Список участников (всего: {total}):\n\n"
        for idx, (uid, data) in enumerate(list(registered_users_db.items())[:50], 1):
            message_text += f"{idx}. {data['name']} | {data['contact']}\n"

        if total > 50:
            message_text += f"\n... и ещё {total - 50} участников."

        message_text += "\n\n⚠️ Не удалось отправить файл. Проверьте логи."
        await message.answer(message_text)


@bot.on.private_message(text=["статистика", "Статистика", "/stat"])
async def admin_stats_handler(message: Message):
    """Краткая статистика для админа."""
    user_id = message.from_id

    if not await is_admin(user_id):
        await message.answer("⛔ У вас нет доступа к этой команде.")
        return

    total = len(registered_users_db)
    today = datetime.now().strftime("%Y-%m-%d")
    today_regs = sum(1 for data in registered_users_db.values() if data["timestamp"].startswith(today))

    # Группы и количество неподписанных (приблизительно)
    await message.answer(
        f"📊 **Статистика регистрации**\n\n"
        f"👥 Всего участников: **{total}**\n"
        f"📅 За сегодня: **{today_regs}**\n"
        f"📌 Групп для подписки: **{len(REQUIRED_GROUPS)}**"
    )


# --- ЗАПУСК ---
if __name__ == "__main__":
    print("✅ Бот запущен!")
    print(f"👑 Администраторы: {ADMIN_IDS}")
    print(f"📌 Групп для проверки: {len(REQUIRED_GROUPS)}")

    # Проверка наличия openpyxl
    try:
        import openpyxl

        print("📦 openpyxl: OK")
    except ImportError:
        print("⚠️ Установите openpyxl: pip install openpyxl")
        exit(1)

    bot.run_forever()