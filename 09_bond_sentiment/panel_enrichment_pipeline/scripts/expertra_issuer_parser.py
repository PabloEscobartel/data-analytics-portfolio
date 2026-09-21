"""
Парсер рейтингов ЭМИТЕНТОВ от Эксперт РА (requests).
Ищет по ИНН (столбец TIN) → блок #companies → первая ссылка.
Парсит рейтинги кредитоспособности (span начинается с "Рейтинги кредитоспособности").
Название компании берёт из h1.b-title.-white.

Столбцы:
  - expertra_issuer_rating, _date, _forecast, _flag, _name

Вход/выход: bonds_retry.xlsx
Кэш: issuer_cache_expertra.json (ключ = имя эмитента)

Требования: pip install requests beautifulsoup4 pandas openpyxl tqdm lxml
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

INPUT_FILE = "bonds_ext_ratings.xlsx"
OUTPUT_FILE = "bonds_ext_ratings.xlsx"
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
# ПОИСК ПО ИНН
# ============================================================

def search_by_inn(inn: str, session: requests.Session):
    """
    Ищет компанию по ИНН на raexpert.ru/search/.
    Берёт первую ссылку на /database/ из результатов (ИНН даёт единственный результат).
    Возвращает (url, search_name) или (None, None).
    """
    url = "https://raexpert.ru/search/"

    time.sleep(DELAY)
    r = session.get(url, headers=HEADERS, verify=False, timeout=15)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    csrf = soup.select_one("input[name='CSRFToken']")
    if not csrf:
        return None, None

    time.sleep(DELAY)
    r = session.post(
        url,
        headers={**HEADERS, "Referer": url},
        data={"CSRFToken": csrf["value"], "search": inn},
        verify=False, timeout=15,
    )
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    # Берём первую ссылку на /database/ из любого блока результатов
    for link in soup.select("a.b-table__text, a.b-table_text"):
        href = link.get("href", "")
        if "/database/" in href:
            search_name = link.get_text(strip=True)
            full_url = ("https://raexpert.ru" + href) if href.startswith("/") else href
            return full_url, search_name

    return None, None


# ============================================================
# ПАРСИНГ СТРАНИЦЫ ЭМИТЕНТА
# ============================================================

def parse_issuer_page(url: str, session: requests.Session):
    """
    Парсит рейтинги кредитоспособности со страницы эмитента.
    Ищет блок, где span.b-actions__subtitle начинается с "Рейтинги кредитоспособности".
    Название берёт из h1.b-title.-white.
    Возвращает (history, site_name).
    """
    time.sleep(DELAY)
    r = session.get(url, headers=HEADERS, verify=False, timeout=15)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    site_name = None
    h1 = soup.select_one("h1.b-title.-white")
    if h1:
        site_name = h1.get_text(strip=True)

    history = []

    table = None
    for block in soup.select("div.b-actions__rates.-black.-company"):
        subtitle = block.select_one("span.b-actions__subtitle")
        if not subtitle:
            continue
        subtitle_text = subtitle.get_text(strip=True).lower()
        if not subtitle_text.startswith("рейтинги кредитоспособности"):
            continue

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

        rating = cells[0].get_text(strip=True)
        rating = re.sub(r'\s+', ' ', rating).strip()

        date_raw = cells[-1].get_text(strip=True)
        date_parsed = None
        match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_raw)
        if match:
            date_parsed = f"{match.group(3)}-{match.group(2)}-{match.group(1)}"

        forecast = ""
        if len(cells) >= 3:
            forecast = cells[1].get_text(strip=True)

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

def get_rating_on_date(history, target_date):
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
    print("Эксперт РА — рейтинги ЭМИТЕНТОВ (поиск по ИНН)")
    print("=" * 60)

    df = pd.read_excel(INPUT_FILE)
    print(f"Загружено: {len(df)} строк")

    # Уникальные пары (эмитент, ИНН)
    pairs = df[["Эмитент", "TIN"]].dropna().drop_duplicates()
    pairs = pairs[pairs["Эмитент"].apply(lambda x: isinstance(x, str) and len(x.strip()) > 1)]
    pairs["TIN"] = pairs["TIN"].astype(str).str.strip().str.split('.').str[0].str.zfill(10)
    print(f"Уникальных эмитентов: {len(pairs)}")

    cache = load_cache()

    # Парсим заново ВСЕ, перезаписывая кэш
    session = requests.Session()
    success = 0
    errors = 0
    not_found = 0

    for _, row in tqdm(pairs.iterrows(), total=len(pairs), desc="Эксперт РА (ИНН)"):
        emitter = row["Эмитент"]
        inn = row["TIN"]

        entry = {"history": [], "url": None, "site_name": None, "error": None, "inn": inn}

        try:
            url, search_name = search_by_inn(inn, session)
            if url:
                history, page_name = parse_issuer_page(url, session)
                site_name = page_name or search_name
                entry["history"] = history
                entry["url"] = url
                entry["site_name"] = site_name

                if history:
                    success += 1
                    tqdm.write(f"  {emitter} (ИНН {inn}) → {site_name} | рейтингов: {len(history)}")
                else:
                    not_found += 1
                    tqdm.write(f"  {emitter} (ИНН {inn}) → {site_name} | нет рейтингов кредитоспособности")
            else:
                not_found += 1
                entry["error"] = "Не найден по ИНН"
                tqdm.write(f"  {emitter} (ИНН {inn}) → не найден")
        except Exception as e:
            errors += 1
            entry["error"] = str(e)
            tqdm.write(f"  {emitter} (ИНН {inn}) → ОШИБКА: {e}")

        cache[emitter] = entry

        if (success + not_found + errors) % 10 == 0:
            save_cache(cache)

        if random.random() < 0.08:
            time.sleep(random.uniform(4, 10))

    save_cache(cache)

    print(f"\n{'='*60}")
    print(f"Спарсено: {success}, не найдено: {not_found}, ошибки: {errors}")

    # --- Применяем к таблице ---
    print("\nПрименяю рейтинги к таблице...")

    col_rating = []
    col_date = []
    col_forecast = []
    col_flag = []
    col_name = []

    for _, row in df.iterrows():
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

    df["expertra_issuer_rating"] = col_rating
    df["expertra_issuer_rating_date"] = col_date
    df["expertra_issuer_forecast"] = col_forecast
    df["expertra_issuer_flag"] = col_flag
    df["expertra_issuer_name"] = col_name

    found = sum(1 for r in col_rating if r is not None)
    after = sum(1 for f in col_flag if f == "после размещения")
    print(f"Найдено рейтингов: {found}/{len(df)} (из них {after} после размещения)")

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Сохранено: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
