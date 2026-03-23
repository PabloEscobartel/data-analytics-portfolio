import os
import cloudinary
import cloudinary.uploader
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import io
from datetime import datetime, timedelta
import pytz
import requests
import xml.etree.ElementTree as ET


def det_data():
    timezone = pytz.timezone('Asia/Yekaterinburg')
    current_date = datetime.now(timezone)
    current_date = current_date.strftime("%d/%m/%Y")
    return current_date


def graph(index, n):

    timezone = pytz.timezone('Asia/Yekaterinburg')
    current_date = datetime.now(timezone)

    if n > 4:

        five_days_ago = current_date - timedelta(days=n)
        current_date, five_days_ago = current_date.strftime("%d/%m/%Y"), five_days_ago.strftime("%d/%m/%Y")
        data = generate_url(five_days_ago, current_date, index)

        return create_and_upload_graph_to_cloudinary(data)

    else:

        five_days_ago = current_date - timedelta(days=4)
        current_date, five_days_ago = current_date.strftime("%d/%m/%Y"), five_days_ago.strftime("%d/%m/%Y")
        data = generate_url(five_days_ago, current_date, index)
        last_row = data.iloc[-1]
        name_value = last_row['Value']

        return round(float(name_value.replace(",", ".")), 2)


def generate_url(first_date, second_date, index):
    ID = determine_id(index)
    url_curr = f'https://www.cbr.ru/scripts/XML_dynamic.asp?date_req1={first_date}&date_req2={second_date}&VAL_NM_RQ={ID}'
    data = pd.read_xml(url_curr)
    return data


def determine_id(index):
    url_code = 'http://www.cbr.ru/scripts/XML_val.asp?d=0'
    data = pd.read_xml(url_code, encoding='cp1251')
    index = determine_name(index)
    row = data[data['Name'] == index]
    ID = row['ID'].values[0]
    return ID


def create_and_upload_graph_to_cloudinary(data):

    cloudinary.config(
        cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
        api_key=os.getenv("CLOUDINARY_API_KEY"),
        api_secret=os.getenv("CLOUDINARY_API_SECRET")
    )

    data['Date'] = pd.to_datetime(data['Date'], format='%d.%m.%Y')
    data['Value'] = data['Value'].str.replace(',', '.').astype(float)

    num_dates = len(data['Date'])
    if 5 < num_dates <= 31:
        plt.figure(figsize=(12, 6))
        plt.plot(data['Date'], data['Value'], linestyle='-', color='b', label='Значение валюты')
        plt.title('График изменения валюты во времени', fontsize=16)
        plt.xlabel('Дата', fontsize=14)
        plt.ylabel('Значение', fontsize=14)
        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=5))
        plt.xticks(fontsize=12)
    elif num_dates <= 5:
        plt.figure(figsize=(12, 6))
        plt.plot(data['Date'], data['Value'], marker='o', linestyle='-', color='b', label='Значение валюты')
        plt.title('График изменения валюты во времени', fontsize=16)
        plt.xlabel('Дата', fontsize=14)
        plt.ylabel('Значение', fontsize=14)
        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=1))
        plt.xticks(fontsize=12)
    else:
        plt.figure(figsize=(12, 6))
        plt.plot(data['Date'], data['Value'], linestyle='-', color='b', label='Значение валюты')
        plt.title('График изменения валюты во времени', fontsize=16)
        plt.xlabel('Дата', fontsize=14)
        plt.ylabel('Значение', fontsize=14)
        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=30))
        plt.xticks(fontsize=12, rotation=45)

    plt.grid(True, linestyle='--', alpha=0.7)
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d.%m.%Y'))
    plt.legend(fontsize=12)
    plt.gca().set_facecolor('#f3f3f3')
    plt.gca().set_axisbelow(True)
    plt.yticks(fontsize=12)
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    plt.ylim(min(data['Value']) - 1, max(data['Value']) + 1)
    plt.tight_layout()

    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png')
    img_buffer.seek(0)

    try:
        response = cloudinary.uploader.upload(img_buffer)
        return response['secure_url']
    except Exception as e:
        print(f"Произошла ошибка при загрузке изображения: {str(e)}")
        return None


def determine_char(input_currency):
    for currency, code in currency_dict.items():
        if input_currency == currency or input_currency == code:
            return code
    return "Нет совпадений"


def determine_name(input_currency):
    for currency, code in currency_dict.items():
        if input_currency == currency or input_currency == code:
            return currency
    return "Нет совпадений"


def konverter_valut(index1, index2, current_date):
    if index1 == "RUB":
        curr1_rate = 1
        curr1_nominal = 1
    else:
        curr1_rate = float(ET.fromstring(requests.get(f"http://www.cbr.ru/scripts/XML_daily.asp?date_req={current_date}").text).find(f"./Valute[CharCode='{index1}']/Value").text.replace(",", "."))
        curr1_nominal = float(ET.fromstring(requests.get(f"http://www.cbr.ru/scripts/XML_daily.asp?date_req={current_date}").text).find(f"./Valute[CharCode='{index1}']/Nominal").text.replace(",", "."))
    real_curr1_rate = curr1_rate/curr1_nominal

    if index2 == "RUB":
        curr2_rate = 1
        curr2_nominal = 1
    else:
        curr2_rate = float(ET.fromstring(requests.get(f"http://www.cbr.ru/scripts/XML_daily.asp?date_req={current_date}").text).find(f"./Valute[CharCode='{index2}']/Value").text.replace(",", "."))
        curr2_nominal = float(ET.fromstring(requests.get(f"http://www.cbr.ru/scripts/XML_daily.asp?date_req={current_date}").text).find(f"./Valute[CharCode='{index2}']/Nominal").text.replace(",", "."))
    real_curr2_rate = curr2_rate/curr2_nominal
    relat = real_curr1_rate/real_curr2_rate
    return(relat)


def have_n(days):
    if days == '5 дней':
        n = 5
    elif days == 'месяц':
        n = 30
    elif days == 'год':
        n = 365
    else:
        n = 4
    return n


currency_dict = {
    "Австралийский доллар": "AUD",
    "Азербайджанский манат": "AZN",
    "Фунт стерлингов Соединенного королевства": "GBP",
    "Армянский драм": "AMD",
    "Белорусский рубль": "BYN",
    "Болгарский лев": "BGN",
    "Бразильский реал": "BRL",
    "Венгерский форинт": "HUF",
    "Вьетнамский донг": "VND",
    "Гонконгский доллар": "HKD",
    "Грузинский лари": "GEL",
    "Датская крона": "DKK",
    "Дирхам ОАЭ": "AED",
    "Доллар США": "USD",
    "Евро": "EUR",
    "Египетский фунт": "EGP",
    "Индийская рупия": "INR",
    "Индонезийская рупия": "IDR",
    "Казахстанский тенге": "KZT",
    "Канадский доллар": "CAD",
    "Катарский риал": "QAR",
    "Киргизский сом": "KGS",
    "Китайский юань": "CNY",
    "Молдавский лей": "MDL",
    "Новозеландский доллар": "NZD",
    "Норвежская крона": "NOK",
    "Польский злотый": "PLN",
    "Российский рубль": "RUB",
    "Румынский лей": "RON",
    "СДР (специальные права заимствования)": "XDR",
    "Сингапурский доллар": "SGD",
    "Таджикский сомони": "TJS",
    "Таиландский бат": "THB",
    "Турецкая лира": "TRY",
    "Новый туркменский манат": "TMT",
    "Узбекский сум": "UZS",
    "Украинская гривна": "UAH",
    "Чешская крона": "CZK",
    "Шведская крона": "SEK",
    "Швейцарский франк": "CHF",
    "Сербский динар": "RSD",
    "Южноафриканский рэнд": "ZAR",
    "Вон Республики Корея": "KRW",
    "Японская иена": "JPY"
}


def get_currency_values():
    values = []
    for currency, code in currency_dict.items():
        values.append(f"{currency}: {code}")
    return '\n'.join(values)
