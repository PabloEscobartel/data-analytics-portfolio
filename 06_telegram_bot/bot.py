import os
import asyncio
from aiogram import Bot, Dispatcher
from handlers import commands, dynamics, converter


async def main():
    bot = Bot(token=os.getenv('BOT_TOKEN'), parse_mode="HTML")
    dp = Dispatcher()

    dp.include_routers(commands.router)
    dp.include_routers(dynamics.router)
    dp.include_routers(converter.router)

    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print('Exit')
