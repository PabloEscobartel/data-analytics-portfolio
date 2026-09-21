"""
Дособор/перепарсинг рейтингов эмитентов Эксперт РА.
Читает retake_issuers_expertra.xlsx:
  - Если есть URL → переходит и парсит
  - Если URL нет → ставит пустую запись в кэш

Парсит ТОЛЬКО таблицу, перед которой span содержит "Рейтинги кредитоспособности".
Название компании берёт из h1.b-title.-white.

Перезаписывает записи в issuer_cache_expertra.json.
"""

import json
import os
import re
import time
import random

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# НАСТРОЙКИ
# ============================================================

RETAKE_FILE = "retake_issuers_expertra.xlsx"
CACHE_FILE = "issuer_cache_expertra.json"
DELAY = 1.5

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "Referer": "https://raexpert.ru/",
}


# ============================================================
# КЭШ
# ============================================================

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ============================================================
# ПАРСИНГ СТРАНИЦЫ ЭМИТЕНТА
# ============================================================

def parse_issuer_page(url: str, session: requests.Session):
    """
    Парсит страницу эмитента на Эксперт РА.
    Ищет таблицу, которая идёт после span с текстом "Рейтинги кредитоспособности ...".
    Возвращает (history, site_name).
    """
    time.sleep(DELAY)
    r = session.get(url, headers=HEADERS, verify=False, timeout=15)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    # Название компании с сайта: h1.b-title.-white
    site_name = None
    h1 = soup.select_one("h1.b-title.-white")
    if h1:
        site_name = h1.get_text(strip=True)

    history = []

    # Перебираем все блоки div.b-actions__rates.-black.-company
    # Ищем тот, у которого span.b-actions__subtitle содержит "Рейтинги кредитоспособности"
    table = None
    for block in soup.select("div.b-actions__rates.-black.-company"):
        subtitle = block.select_one("span.b-actions__subtitle")
        if not subtitle:
            continue

        subtitle_text = subtitle.get_text(strip=True).lower()
        if not subtitle_text.startswith("рейтинги кредитоспособности"):
            continue

        # Нашли нужный блок — берём таблицу
        table = block.select_one("table.object-rating-table")
        if not table:
            table = block.select_one("table")
        if table:
            break

    if not table:
        return history, site_name

    for row in table.select("tbody tr, tr"):
        cells = row.select("td")
        if len(cells) < 2:
            continue

        # Рейтинг — первый td
        rating = cells[0].get_text(strip=True)
        rating = re.sub(r'\s+', ' ', rating).strip()

        # Дата — последний td
        date_raw = cells[-1].get_text(strip=True)
        date_parsed = None
        match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_raw)
        if match:
            date_parsed = f"{match.group(3)}-{match.group(2)}-{match.group(1)}"

        # Прогноз — второй td (если >= 3 колонок)
        forecast = ""
        if len(cells) >= 3:
            forecast = cells[1].get_text(strip=True)

        # Пропускаем заголовки
        skip_words = ["национальная", "шкала", "дата", "прогноз"]
        if not rating or any(s in rating.lower() for s in skip_words):
            continue

        history.append({
            "rating": rating,
            "forecast": forecast,
            "date": date_parsed,
            "date_raw": date_raw,
        })

    return history, site_name


# ============================================================
# РЕЙТИНГ НА ДАТУ
# ============================================================

def get_rating_on_date(history: list, target_date: str):
    """
    Возвращает (rating, forecast, date, flag).
    flag = "до размещения" / "после размещения" / None.
    """
    if not history or not target_date:
        return None, None, None, None

    dated = [h for h in history if h.get("date") and h.get("rating")]
    active = [h for h in dated if "отозван" not in h["rating"].lower()]
    if not active:
        active = dated
    if not active:
        return None, None, None, None

    active.sort(key=lambda x: x["date"])

    best = None
    for r in active:
        if r["date"] <= target_date:
            best = r

    if best:
        return best["rating"], best.get("forecast", ""), best["date"], "до размещения"

    earliest = active[0]
    return earliest["rating"], earliest.get("forecast", ""), earliest["date"], "после размещения"


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("Эксперт РА — дособор рейтингов эмитентов")
    print("=" * 60)

    df = pd.read_excel(RETAKE_FILE)
    print(f"Всего эмитентов в retake: {len(df)}")
    print(f"  С URL: {df['url'].notna().sum()}")
    print(f"  Без URL: {df['url'].isna().sum()}")

    cache = load_cache()
    session = requests.Session()

    with_url = df[df["url"].notna()]
    without_url = df[df["url"].isna()]

    # --- Эмитенты БЕЗ URL → пустая запись ---
    for _, row in without_url.iterrows():
        name = row["name"]
        cache[name] = {
            "history": [],
            "url": None,
            "site_name": None,
            "error": "Не найден в #companies",
        }
    print(f"\nПроставлено 'не найден' для {len(without_url)} эмитентов")

    # --- Эмитенты С URL → парсим ---
    success = 0
    errors = 0

    for _, row in tqdm(with_url.iterrows(), total=len(with_url), desc="Парсинг"):
        name = row["name"]
        url = row["url"]

        try:
            history, site_name = parse_issuer_page(url, session)
            cache[name] = {
                "history": history,
                "url": url,
                "site_name": site_name,
                "error": None,
            }

            if history:
                success += 1
                tqdm.write(f"  {name} → {site_name} | рейтингов: {len(history)}")
            else:
                tqdm.write(f"  {name} → {site_name} | рейтингов: 0 (нет блока кредитоспособности)")

        except Exception as e:
            errors += 1
            cache[name] = {
                "history": [],
                "url": url,
                "site_name": None,
                "error": str(e),
            }
            tqdm.write(f"  {name} → ОШИБКА: {e}")

        # Сохраняем каждые 5
        if (success + errors) % 5 == 0:
            save_cache(cache)

        # Случайная пауза
        if random.random() < 0.1:
            time.sleep(random.uniform(3, 8))

    save_cache(cache)

    print(f"\n{'='*60}")
    print(f"Спарсено с рейтингами: {success}")
    print(f"Ошибки: {errors}")
    print(f"Не найдено (без URL): {len(without_url)}")
    print(f"Кэш обновлён: {CACHE_FILE}")

    # --- Применяем к bonds_retry.xlsx ---
    BONDS_FILE = "bonds_retry.xlsx"
    print(f"\nПрименяю рейтинги к {BONDS_FILE}...")

    bonds = pd.read_excel(BONDS_FILE)

    col_rating = []
    col_date = []
    col_forecast = []
    col_flag = []
    col_name = []

    for _, row in bonds.iterrows():
        emitter = row.get("Эмитент")
        placement = row.get("Начало размещения")
        target = None
        if pd.notna(placement):
            target = pd.Timestamp(placement).strftime("%Y-%m-%d")

        rating, forecast, date, flag = None, None, None, None
        site_name = None

        if emitter and isinstance(emitter, str) and emitter in cache:
            entry = cache[emitter]
            hist = entry.get("history", [])
            site_name = entry.get("site_name")
            rating, forecast, date, flag = get_rating_on_date(hist, target)

        col_rating.append(rating)
        col_date.append(date)
        col_forecast.append(forecast)
        col_flag.append(flag)
        col_name.append(site_name)

    bonds["expertra_issuer_rating"] = col_rating
    bonds["expertra_issuer_rating_date"] = col_date
    bonds["expertra_issuer_forecast"] = col_forecast
    bonds["expertra_issuer_flag"] = col_flag
    bonds["expertra_issuer_name"] = col_name

    found = sum(1 for r in col_rating if r is not None)
    after = sum(1 for f in col_flag if f == "после размещения")
    print(f"Найдено рейтингов: {found}/{len(bonds)} (из них {after} после размещения)")

    bonds.to_excel(BONDS_FILE, index=False)
    print(f"Сохранено: {BONDS_FILE}")


if __name__ == "__main__":
    main()
