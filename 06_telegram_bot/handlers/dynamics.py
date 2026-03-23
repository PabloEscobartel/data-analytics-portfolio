from aiogram import F, Router
from aiogram.types import Message, CallbackQuery
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from functions import cloud as cd
from keyboards import keyboards as kb
import re
from datetime import datetime


router = Router()


class DynamicsOption(StatesGroup):
    choosing_curr = State()
    choosing_another_curr = State()
    choosing_period = State()
    choosing_first_date = State()
    choosing_second_date = State()


@router.message(F.text == 'Динамика курса')
async def curs_dyn(message: Message, state: FSMContext):
    await message.answer('Выберите валюту:', reply_markup=kb.currencies)
    await state.set_state(DynamicsOption.choosing_curr)


currencies = ['USD', 'EUR']


@router.callback_query(DynamicsOption.choosing_curr, F.data.in_(currencies))
async def send_usd_dyn(callback: CallbackQuery, state: FSMContext):
    currency = callback.data
    await state.update_data(index=currency)
    await callback.message.answer(f'Вы выбрали: <b>{currency}</b>.\nВыберите период:', reply_markup=kb.period)
    await callback.answer()
    await state.set_state(DynamicsOption.choosing_period)


@router.callback_query(DynamicsOption.choosing_curr, F.data == "another")
async def send_usd_dyn(callback: CallbackQuery, state: FSMContext):
    await callback.message.answer(f'Введите полное название валюты или её трёхзначный код:')
    await state.set_state(DynamicsOption.choosing_another_curr)
    await callback.answer()


@router.message(DynamicsOption.choosing_another_curr, F.text)
async def curs_dyn(message: Message, state: FSMContext):
    user_input = message.text
    index = cd.determine_char(user_input)
    await state.update_data(index=index)
    if index == 'Нет совпадений':
        await message.answer(f'Такая валюта не найдена. Попробуйте ввести снова!\n'
                             'Список возможных валют и их кодов: /help')
        await state.set_state(DynamicsOption.choosing_another_curr)
    else:
        await message.answer(f'Вы выбрали: <b>{user_input}</b>.\nВыберите период:', reply_markup=kb.period)
        await state.set_state(DynamicsOption.choosing_period)


periods = ['сегодня', '5 дней', 'месяц', 'год']


@router.callback_query(DynamicsOption.choosing_period, F.data.in_(periods))
async def send_usd_dyn(callback: CallbackQuery, state: FSMContext):
    user_data = await state.get_data()
    name_curr = cd.determine_name(user_data["index"])
    days = callback.data
    n = cd.have_n(days)
    gh = cd.graph(user_data['index'], n)
    if isinstance(gh, float):
        await callback.message.answer(f'Официальный курс валюты <b>"{name_curr}"</b> на сегодня равен <b>{gh}</b> рубля(м)')
    else:
        await callback.message.answer_photo(photo=gh,
                                            caption=f'Изменение курса валюты<b>"{name_curr}"</b> за последние(й) <b>{days}</b>')
    await callback.answer()


@router.callback_query(DynamicsOption.choosing_period, F.data == 'other period')
async def send_usd_dyn(callback: CallbackQuery, state: FSMContext):
    await callback.message.answer(f'Укажите начальную дату в формате дд/мм/гггг:')
    await state.set_state(DynamicsOption.choosing_first_date)
    await callback.answer()


@router.message(DynamicsOption.choosing_first_date, F.text)
async def first_date_det(message: Message, state: FSMContext):
    user_input = message.text
    pattern = r'\d{2}/\d{2}/\d{4}'
    if re.match(pattern, user_input):
        await state.update_data(first_date=message.text)
        await message.reply('Введите конечную дату:')
        await state.set_state(DynamicsOption.choosing_second_date)
    else:
        await message.answer('Введите корректные данные')
        await state.set_state(DynamicsOption.choosing_first_date)


@router.message(DynamicsOption.choosing_second_date, F.text)
async def first_date(message: Message, state: FSMContext):
    user_input = message.text
    pattern = r'\d{2}/\d{2}/\d{4}'
    user_data = await state.get_data()
    name_curr = cd.determine_name(user_data["index"])
    date1 = datetime.strptime(user_data["first_date"], "%d/%m/%Y").date()
    date2 = datetime.strptime(user_input, "%d/%m/%Y").date()
    if re.match(pattern, user_input):
        if date2 > date1:
            data = cd.generate_url(user_data["first_date"], user_input, user_data['index'])
            gh = cd.create_and_upload_graph_to_cloudinary(data)
            await message.answer_photo(photo=gh, caption=f'Изменение курса валюты <b>"{name_curr}"</b> за период с <b>{user_data["first_date"]}</b> по <b>{user_input}</b>')
            await state.clear()
        else:
            await message.answer('Введены некорректные данные. Попробуйте ввести данные снова. \n'
                                 'Введите начальную дату:')
            await state.set_state(DynamicsOption.choosing_first_date)
    else:
        await message.answer('Введите корректные данные')
        await state.set_state(DynamicsOption.choosing_second_date)
