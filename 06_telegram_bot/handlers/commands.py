from aiogram import Router
from aiogram.types import Message
from aiogram.filters import Command

from functions import cloud as cd
from keyboards import keyboards as kb

router = Router()


@router.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer('Добро пожаловать!\n'
                         '1. <b>"Динамика курса"</b> - позволяет узнать как изменялась цена валюты за интересующий отрезок времени\n'
                         '2. <b>"Конвертер"</b> - позволяет перевести одну валюту в другую по актуальному официальному курсу',
                         reply_markup=kb.main)


@router.message(Command('help'))
async def cmd_help(message: Message):
    char_code = cd.get_currency_values()
    await message.reply(f'<u><b>Список соответствия валют и кодов:</b></u>\n{char_code}')
