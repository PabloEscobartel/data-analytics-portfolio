# 💱 Telegram-бот: Курсы валют ЦБ РФ

## Описание
Асинхронный Telegram-бот для работы с валютными курсами Центрального банка РФ. Два режима: просмотр динамики курса с графиком и конвертер валют по актуальному курсу.

**Деплой:** Amvera Cloud

## Функционал

### 📈 Динамика курса
- Выбор валюты (USD, EUR или любая из ~45 валют ЦБ)
- Периоды: сегодня, 5 дней, месяц, год или произвольный диапазон
- Генерация графика (matplotlib) с автоматическим масштабированием осей
- Загрузка графика в Cloudinary и отправка пользователю как фото

### 🔄 Конвертер валют
- Конвертация между любыми валютами из списка ЦБ (включая RUB)
- Актуальный курс на текущую дату через XML API ЦБ РФ
- Валидация ввода на каждом шаге (FSM)

## Архитектура
```
bot.py                          ← точка входа, инициализация Bot + Dispatcher
├── handlers/
│   ├── commands.py             ← /start, /help
│   ├── dynamics.py             ← FSM: выбор валюты → период → график
│   └── converter.py            ← FSM: сумма → валюта 1 → валюта 2 → результат
├── functions/
│   └── cloud.py                ← API ЦБ, генерация графиков, Cloudinary, конвертация
└── keyboards/
    └── keyboards.py            ← ReplyKeyboard + InlineKeyboard
```

## Ключевые технические решения
- **aiogram 3.x** — асинхронный фреймворк (async/await), роутеры, FSM (Finite State Machine) для пошагового диалога
- **API ЦБ РФ** — парсинг XML-ответов (`xml.etree.ElementTree`), динамические курсы за произвольный период
- **matplotlib** — генерация графиков в буфер (`io.BytesIO`), адаптивное форматирование осей в зависимости от периода
- **Cloudinary** — загрузка изображений через API для отправки в Telegram
- **pandas** — обработка XML-данных ЦБ (`pd.read_xml`)

## Стек
- **Python:** asyncio, aiogram 3.x, pandas, matplotlib
- **API:** ЦБ РФ (XML), Cloudinary
- **Деплой:** Amvera Cloud

## Запуск
```bash
# Установить зависимости
pip install aiogram pandas matplotlib cloudinary pytz requests

# Задать переменные окружения
export BOT_TOKEN="your_bot_token"
export CLOUDINARY_CLOUD_NAME="your_cloud_name"
export CLOUDINARY_API_KEY="your_api_key"
export CLOUDINARY_API_SECRET="your_api_secret"

# Запустить
python bot.py
```
