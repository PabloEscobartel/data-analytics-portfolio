from aiogram import F, Router
from aiogram.types import Message
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from functions import cloud as cd


router = Router()


class Converter(StatesGroup):
    input_summa = State()
    input_first_curr = State()
    input_second_curr = State()


@router.message(F.text == 'Конвертер')
async def curs_dyn(message: Message, state: FSMContext):
    await message.answer('Введите сумму:')
    await state.set_state(Converter.input_summa)


@router.message(Converter.input_summa, F.text)
async def curs_dyn(message: Message, state: FSMContext):
    try:
        if float(message.text.replace(",", ".")) > 0:
            await state.update_data(summa=message.text)
            await message.answer('Введите трёхзначный код или название валюты, которую хотите конвертировать \n'
                                 '(например, <b>Евро</b> или <b>EUR</b>):')
            await state.set_state(Converter.input_first_curr)
        else:
            await message.answer('Введённые данные некорректны, проверьте, что вы ввели число больше нуля')
            await state.set_state(Converter.input_summa)
    except ValueError:
        await message.answer('Введённые данные некорректны, проверьте, что вы ввели число больше нуля')
        await state.set_state(Converter.input_summa)


@router.message(Converter.input_first_curr, F.text)
async def curs_dyn(message: Message, state: FSMContext):
    char = cd.determine_char(message.text)
    if char == 'Нет совпадений':
        await message.answer('Введённые данные некорректны, проверьте, что вы ввели код валюты или название правильно \n'
                             '(например, <b>Евро - EUR</b>):\n'
                             '(/help - возможные названия и коды)')
        await state.set_state(Converter.input_first_curr)
    else:
        await state.update_data(first_curr=char)
        await message.answer('Введите трёхзначный код или название валюты, в которую хотите конвертировать \n'
                             '(например, <b>Евро</b> или <b>EUR</b>):')
        await state.set_state(Converter.input_second_curr)


@router.message(Converter.input_second_curr, F.text)
async def curs_dyn(message: Message, state: FSMContext):
    char = cd.determine_char(message.text)
    if char == 'Нет совпадений':
        await message.answer('Введённые данные некорректны, проверьте, что вы ввели код валюты или название правильно \n'
                             '(например, <b>Евро - EUR</b>):\n'
                             '(/help - возможные названия и коды)')
        await state.set_state(Converter.input_second_curr)
    else:
        second_curr = char
        user_data = await state.get_data()
        date = cd.det_data()
        relation = cd.konverter_valut(user_data['first_curr'], second_curr, date)
        result = round(float(user_data['summa'].replace(",", ".")) * relation, 2)
        await message.answer(f'<b>{user_data["first_curr"]}({user_data["summa"]}) → {second_curr}: {result}</b>')
        await state.clear()
