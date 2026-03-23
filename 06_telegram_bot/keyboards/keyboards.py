from aiogram.types import (ReplyKeyboardMarkup, KeyboardButton,
                           InlineKeyboardMarkup, InlineKeyboardButton)

main_kb = [
    [KeyboardButton(text='Конвертер'),
     KeyboardButton(text='Динамика курса')],
]

main = ReplyKeyboardMarkup(keyboard=main_kb,
                           resize_keyboard=True,
                           input_field_placeholder='Выберите пункт ниже')

currencies = InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text='USD', callback_data='USD'),
     InlineKeyboardButton(text='EUR', callback_data='EUR')],
    [InlineKeyboardButton(text='Другая валюта', callback_data='another')]
])

period = InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text='На сегодня', callback_data='сегодня'),
     InlineKeyboardButton(text='За 5 дней', callback_data='5 дней')],
    [InlineKeyboardButton(text='За месяц', callback_data='месяц'),
     InlineKeyboardButton(text='За год', callback_data='год')],
    [InlineKeyboardButton(text='Другой период', callback_data='other period')]
])
