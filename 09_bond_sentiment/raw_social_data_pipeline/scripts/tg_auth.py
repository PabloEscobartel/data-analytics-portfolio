"""
Отдельный скрипт авторизации в Telegram.
Запусти его один раз — создастся файл bond_parser.session.
После этого основной парсер будет работать без повторной авторизации.

Запуск: python tg_auth.py
"""

import asyncio
import os
from telethon import TelegramClient
from telethon.errors import SessionPasswordNeededError

API_ID = int(os.environ["TELEGRAM_API_ID"])
API_HASH = os.environ["TELEGRAM_API_HASH"]
SESSION  = "bond_parser"

async def main():
    client = TelegramClient(
        SESSION, API_ID, API_HASH,
        device_model="MacBook Air",
        system_version="macOS 14.0",
        app_version="1.0",
        lang_code="ru",
        system_lang_code="ru-RU"
    )
    await client.connect()

    if await client.is_user_authorized():
        me = await client.get_me()
        print(f"Уже авторизован как: {me.first_name} (@{me.username})")
        await client.disconnect()
        return

    phone = input("Введи номер телефона (например +79991234567): ").strip()

    await client.send_code_request(phone)
    print("Код отправлен — проверь приложение Telegram на телефоне (не SMS!)")

    code = input("Введи код: ").strip()

    try:
        await client.sign_in(phone, code)
    except SessionPasswordNeededError:
        password = input("Введи пароль двухфакторной аутентификации: ").strip()
        await client.sign_in(password=password)

    me = await client.get_me()
    print(f"\nУспешно! Авторизован как: {me.first_name} (@{me.username})")
    print(f"Файл сессии создан: {SESSION}.session")
    print(f"\nТеперь запускай:")
    print(f"  python telegram_parser.py --channels probonds --limit 50")

    await client.disconnect()


if __name__ == "__main__":
    asyncio.run(main())
