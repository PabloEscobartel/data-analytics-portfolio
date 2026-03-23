"""
Парсинг рейтингов облигаций от Эксперт РА.
Для каждого ISIN из bonds_retry.xlsx ищет рейтинг, актуальный на дату размещения.
Добавляет два столбца: expertra_rating, expertra_rating_date.

Требования: pip install requests beautifulsoup4 pandas openpyxl tqdm lxml
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import json
import os
import urllib3
from tqdm import tqdm

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# НАСТРОЙКИ
# ============================================================

INPUT_FILE = "bonds_retry.xlsx"
OUTPUT_FILE = "bonds_retry.xlsx"          # перезаписываем тот же файл
CACHE_FILE = "rating_cache_expertra.json"
DELAY = 1.0

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "Referer": "https://raexpert.ru/",
}


# ============================================================
# КЭШ
# ============================================================

def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_cache(cache: dict):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ============================================================
# ЭКСПЕРТ РА — поиск и парсинг
# ============================================================

def expertra_search(isin: str, session: requests.Session) -> str:
    """Ищет ISIN на raexpert.ru, возвращает URL страницы рейтинга или None."""
    search_url = "https://raexpert.ru/search/"

    time.sleep(DELAY)
    resp = session.get(search_url, headers=HEADERS, verify=False, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")
    csrf_input = soup.select_one("input[name='CSRFToken']")
    if not csrf_input:
        return None
    csrf_token = csrf_input["value"]

    time.sleep(DELAY)
    resp = session.post(
        search_url,
        headers={**HEADERS, "Referer": search_url},
        data={"CSRFToken": csrf_token, "search": isin},
        verify=False, timeout=15,
    )
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")

    # Стратегия 1: div.b-table__item содержит ISIN
    for item_div in soup.select("div.b-table__item"):
        if isin in item_div.get_text():
            link = item_div.select_one("a.b-table__text")
            if link:
                href = link.get("href", "")
                if "/database/" in href:
                    return ("https://raexpert.ru" + href) if href.startswith("/") else href

    # Стратегия 2: текст ссылки содержит ISIN
    for a in soup.select("a.b-table__text"):
        if isin in a.get_text() and "/database/" in a.get("href", ""):
            href = a["href"]
            return ("https://raexpert.ru" + href) if href.startswith("/") else href

    return None


def expertra_get_history(page_url: str, session: requests.Session) -> list:
    """Получает историю рейтингов со страницы Эксперт РА."""
    time.sleep(DELAY)
    resp = session.get(page_url, headers=HEADERS, verify=False, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")

    history = []
    for div in soup.select("div.b-actions"):
        table = div.select_one("table")
        if not table:
            continue
        for row in table.select("tr"):
            cells = row.select("td")
            if len(cells) >= 2:
                rating = cells[0].get_text(strip=True)
                date_text = cells[1].get_text(strip=True)
                date_match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_text)
                date_parsed = None
                if date_match:
                    date_parsed = f"{date_match.group(3)}-{date_match.group(2)}-{date_match.group(1)}"

                skip = ["национальная", "шкала", "рейтинг", "дата"]
                if rating and not any(s in rating.lower() for s in skip):
                    history.append({
                        "rating": rating,
                        "date": date_parsed,
                        "date_raw": date_text,
                    })

    return history


# ============================================================
# РЕЙТИНГ НА ДАТУ РАЗМЕЩЕНИЯ
# ============================================================

def get_rating_on_date(history: list, target_date: str):
    """
    Возвращает (rating, date) — рейтинг, актуальный на target_date.
    Берём самый свежий рейтинг, присвоенный не позже target_date.
    Если все рейтинги позже target_date — берём самый ранний (ближайший к дате).
    Пропускает отозванные рейтинги.
    """
    if not history or not target_date:
        return None, None

    dated = [h for h in history if h.get("date") and h.get("rating")]
    active = [h for h in dated if "отозван" not in h["rating"].lower()]
    if not active:
        active = dated
    if not active:
        return None, None

    active.sort(key=lambda x: x["date"])

    # Ищем последний рейтинг, присвоенный <= target_date
    best = None
    for r in active:
        if r["date"] <= target_date:
            best = r

    if best:
        return best["rating"], best["date"]

    # Все рейтинги позже — берём самый ранний
    return active[0]["rating"], active[0]["date"]


# ============================================================
# СБОР ДАННЫХ
# ============================================================

def scrape_all(isins: list, cache: dict) -> dict:
    session = requests.Session()
    new = 0
    for isin in tqdm(isins, desc="Эксперт РА"):
        if isin in cache:
            continue

        cache[isin] = {"history": [], "url": None, "error": None}

        try:
            url = expertra_search(isin, session)
            if url:
                history = expertra_get_history(url, session)
                cache[isin]["history"] = history
                cache[isin]["url"] = url
        except Exception as e:
            cache[isin]["error"] = str(e)

        new += 1
        if new % 20 == 0:
            save_cache(cache)
            time.sleep(1.5)

    save_cache(cache)
    return cache


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("Парсинг рейтингов Эксперт РА")
    print("=" * 60)

    df = pd.read_excel(INPUT_FILE)
    print(f"Загружено: {len(df)} строк")

    isins = df["ISIN"].dropna().unique().tolist()
    isins = [i for i in isins if isinstance(i, str) and len(i.strip()) > 10]

    cache = load_cache()
    to_scrape = [i for i in isins if i not in cache]

    print(f"Всего ISIN: {len(isins)}")
    print(f"Уже в кэше: {len(isins) - len(to_scrape)}")
    print(f"Нужно спарсить: {len(to_scrape)}")

    if to_scrape:
        est_min = len(to_scrape) * DELAY * 2.5 / 60
        print(f"Оценка времени: ~{est_min:.1f} мин")
        resp = input("Продолжить? (y/n): ").strip().lower()
        if resp != "y":
            print("Прервано.")
            return
        cache = scrape_all(to_scrape, cache)

    # Применяем рейтинги
    ratings_col = []
    dates_col = []

    for _, row in df.iterrows():
        isin = row.get("ISIN")
        placement = row.get("Начало размещения")
        target = None
        if pd.notna(placement):
            target = pd.Timestamp(placement).strftime("%Y-%m-%d")

        rating, date = None, None
        if isin and isinstance(isin, str) and isin in cache:
            history = cache[isin].get("history", [])
            rating, date = get_rating_on_date(history, target)

        ratings_col.append(rating)
        dates_col.append(date)

    df["expertra_rating"] = ratings_col
    df["expertra_rating_date"] = dates_col

    found = sum(1 for r in ratings_col if r is not None)
    print(f"\nНайдено рейтингов: {found}/{len(df)}")

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Сохранено: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
